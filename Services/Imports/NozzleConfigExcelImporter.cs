using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI.Services.Imports;

public sealed class NozzleConfigExcelImporter
{
    private readonly SqlConnectionFactory _connectionFactory;

    public enum CleanedImportMode
    {
        UpsertOnly = 0,
        ReplaceAll = 1
    }

    public NozzleConfigExcelImporter(SqlConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
    }

    public async Task<ImportResult> ImportAsync(
        string xlsxPath,
        string? sheetName,
        CleanedImportMode cleanedMode = CleanedImportMode.UpsertOnly,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(xlsxPath)) throw new ArgumentException("Path cannot be empty.", nameof(xlsxPath));
        if (!File.Exists(xlsxPath)) throw new FileNotFoundException("Excel file not found.", xlsxPath);

        var errors = new List<string>();
        var read = 0;
        var ok = 0;
        var failed = 0;

        // Import strategy:
        // - If the workbook contains a "cleaned" tabular sheet with headers like:
        //       MachineCode(or MachineID) | HeadLabel | SlotIndex | NozzleCode
        //   we import it in one of two modes:
        //     * UpsertOnly (default): update/insert only the rows present in the sheet (no deletes).
        //     * ReplaceAll: DELETE FROM dbo.MachineNozzleConfig; then import rows.
        // - Otherwise we fall back to the matrix-style parsing used for the original NozzleConfig workbook.
        //
        // Matrix-style sheets:
        // - Source: NozzleConfig Excel (e.g. "NozzleConfig-Jun-2023 1.xlsx").
        // - Many sheets are NOT tabular; they are visual matrices:
        //     * A large title at the top describes the machine.
        //     * Each "Changer N" block contains repeated pairs of rows:
        //           Row A: slot indices (e.g. 1..70, sometimes descending, sometimes split blocks like 45-48 and 33-40)
        //           Row B: nozzle code per slot (e.g. 3220, 3430, H076, FJ461)
        // - Machine matching:
        //     * We match each worksheet to an existing dbo.Machine.MachineCode by containment:
        //       if the worksheet name contains a known MachineCode (case-insensitive), we use it.
        //       Example: "FZ120 GC6A" -> "GC6A", "FUZION2 14(GX3)" -> "GX3".
        // - Behavior:
        //   * We import per worksheet + changer:
        //       - detect "Changer N" rows
        //       - within each changer, scan rows; whenever a row looks like a "slot header row",
        //         build a map column->slotIndex; then read the next row as nozzle codes for those slots.
        //   * Nozzle types are upserted into dbo.NozzleType keyed by NozzleCode.
        //   * Machine nozzle slots are upserted into dbo.MachineNozzleConfig keyed by (MachineId, SlotIndex):
        //     existing rows are updated; otherwise inserted. No rows are deleted.
        //   * HeadLabel is stored as text (e.g. "Changer01"). This avoids text-sorting issues
        //     and supports non-changer sections in cleaned datasets (e.g. "InLine07").

        await using var conn = _connectionFactory.Open();
        await using var tx = (SqlTransaction)await conn.BeginTransactionAsync(ct);

        string? tempCopyPath = null;
        try
        {
            tempCopyPath = await CopyToTempAsync(xlsxPath, ct);

            using var wb = new XLWorkbook(tempCopyPath);

            if (TryFindCleanedTableSheet(wb, out var cleanWs, out var headerMap))
            {
                if (cleanedMode == CleanedImportMode.ReplaceAll)
                    await ReplaceAllNozzleConfigAsync(conn, tx, ct);
                var (tRead, tOk, tFailed, tErrors) = await ImportWorksheetCleanTable(cleanWs, headerMap, conn, tx, ct);
                read += tRead;
                ok += tOk;
                failed += tFailed;
                errors.AddRange(tErrors);
                await tx.CommitAsync(ct);
                return new ImportResult(Path.GetFileName(xlsxPath), read, ok, failed, errors);
            }

            var machineCodeSet = await LoadMachineCodesAsync(conn, tx, ct);

            var worksheets = ResolveWorksheets(wb, sheetName);
            foreach (var ws in worksheets)
            {
                ct.ThrowIfCancellationRequested();

                var wsName = ws.Name?.Trim() ?? "";
                var machineCode = TryResolveMachineCodeFromSheetName(wsName, machineCodeSet);
                if (machineCode is null)
                {
                    failed++;
                    errors.Add($"Worksheet '{wsName}': could not match to any existing MachineCode.");
                    continue;
                }

                var machineId = await FindMachineIdAsync(conn, tx, lineCode: null, machineCode, ct);
                if (machineId is null)
                {
                    failed++;
                    errors.Add($"Worksheet '{wsName}': MachineCode '{machineCode}' not found in dbo.Machine.");
                    continue;
                }

                var (wsRead, wsOk, wsFailed, wsErrors) = await ImportWorksheetMatrix(ws, machineId.Value, conn, tx, ct);
                read += wsRead;
                ok += wsOk;
                failed += wsFailed;
                errors.AddRange(wsErrors);
            }

            await tx.CommitAsync(ct);
        }
        catch
        {
            await tx.RollbackAsync(ct);
            throw;
        }
        finally
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(tempCopyPath) && File.Exists(tempCopyPath))
                    File.Delete(tempCopyPath);
            }
            catch
            {
                // ignore temp cleanup failure
            }
        }

        return new ImportResult(Path.GetFileName(xlsxPath), read, ok, failed, errors);
    }

    private static bool TryFindCleanedTableSheet(XLWorkbook wb, out IXLWorksheet ws, out Dictionary<string, int> header)
    {
        ws = null!;
        header = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        static string NormalizeHeaderName(string s)
        {
            // Make "Serial Number/Computer Name", "SerialNumber ComputerName", "Slot Index" etc comparable.
            // Keep digits + letters only.
            var chars = s.Trim().ToUpperInvariant().Where(ch => char.IsLetterOrDigit(ch)).ToArray();
            return new string(chars);
        }

        foreach (var candidate in wb.Worksheets)
        {
            // Search for a header row near the top-left.
            // Some files put the cleaned-table header at B/C/D/E and may include
            // a few title rows above; expand the scan window for robustness.
            var maxRow = Math.Min(50, candidate.LastRowUsed()?.RowNumber() ?? 50);
            var maxCol = Math.Min(25, candidate.LastColumnUsed()?.ColumnNumber() ?? 25);

            for (var r = 1; r <= maxRow; r++)
            {
                var normalized = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                for (var c = 1; c <= maxCol; c++)
                {
                    var name = candidate.Cell(r, c).GetString();
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    var norm = NormalizeHeaderName(name);
                    if (string.IsNullOrWhiteSpace(norm)) continue;
                    if (!normalized.ContainsKey(norm))
                        normalized[norm] = c;
                }

                // After normalization, we match by "contains" to tolerate:
                //   - different delimiters
                //   - suffixes like "HeadLabel (xxx)"
                //   - variations like "Machine ID"
                static int PickColumn(Dictionary<string, int> map, Func<string, bool> predicate)
                {
                    foreach (var (k, v) in map)
                    {
                        if (predicate(k)) return v;
                    }
                    return -1;
                }

                var machineCol =
                    PickColumn(normalized, k => k.Contains("MACHINE") && (k.Contains("CODE") || k.Contains("ID") || k.Contains("NAME")));
                var headCol = PickColumn(normalized, k => k.Contains("HEAD") && k.Contains("LABEL"));
                var slotCol = PickColumn(normalized, k => k.Contains("SLOT") && k.Contains("INDEX"));
                var nozzleCol = PickColumn(normalized, k => k.Contains("NOZZLE") && k.Contains("CODE"));

                if (machineCol < 0 || headCol < 0 || slotCol < 0 || nozzleCol < 0)
                    continue;

                header["MachineCode"] = machineCol;
                header["HeadLabel"] = headCol;
                header["SlotIndex"] = slotCol;
                header["NozzleCode"] = nozzleCol;

                // Optional line code/name
                var lineCol =
                    PickColumn(normalized, k => k.Contains("LINE") && (k.Contains("CODE") || k.Contains("NAME")));
                if (lineCol > 0)
                    header["LineCode"] = lineCol;

                header["__HEADER_ROW__"] = r;
                ws = candidate;
                return true;
            }
        }

        return false;
    }

    private static async Task ReplaceAllNozzleConfigAsync(SqlConnection conn, SqlTransaction tx, CancellationToken ct)
    {
        // Cleaned table is the source of truth; clear existing nozzle configuration to prevent pollution.
        await using var cmd = new SqlCommand("DELETE FROM dbo.MachineNozzleConfig;", conn, tx);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    private async Task<(int read, int ok, int failed, List<string> errors)> ImportWorksheetCleanTable(
        IXLWorksheet ws,
        Dictionary<string, int> headerMap,
        SqlConnection conn,
        SqlTransaction tx,
        CancellationToken ct)
    {
        var errors = new List<string>();
        var read = 0;
        var ok = 0;
        var failed = 0;

        var headerRow = headerMap.TryGetValue("__HEADER_ROW__", out var hr) ? hr : 1;
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? headerRow;

        // Excel often uses merged cells for Machine ID; ClosedXML may return empty for merged "shadow" cells.
        // We only carry forward MachineCode — never inherit HeadLabel/SlotIndex from the previous row, otherwise
        // intentionally blank Head/Slot cells incorrectly reuse values like "Changer04" / "48" from above.
        // For Head/Slot merges, read the merged range's first cell (Excel stores the value only there).
        string? lastMachineCode = null;
        static string CellToString(IXLCell cell)
        {
            // Resolve to the top-left cell of a merge so shadow cells get the same text as Excel shows.
            var c = cell;
            if (c.IsMerged())
                c = c.MergedRange().FirstCell();

            var s = c.GetString()?.Trim();
            if (!string.IsNullOrWhiteSpace(s)) return s;
            return c.Value.ToString()?.Trim() ?? "";
        }

        for (var r = headerRow + 1; r <= lastRow; r++)
        {
            ct.ThrowIfCancellationRequested();

            var machineCodeCell = CellToString(ws.Cell(r, headerMap["MachineCode"]));
            var headLabelCell = CellToString(ws.Cell(r, headerMap["HeadLabel"]));
            var slotRaw = CellToString(ws.Cell(r, headerMap["SlotIndex"]));
            var nozzleCode = NormalizeNozzleCode(CellToString(ws.Cell(r, headerMap["NozzleCode"])));

            if (!string.IsNullOrWhiteSpace(machineCodeCell))
                lastMachineCode = machineCodeCell;

            var machineCode = string.IsNullOrWhiteSpace(machineCodeCell) ? lastMachineCode ?? "" : machineCodeCell;
            var headLabel = headLabelCell;

            // SlotIndex: parse only what this row actually has (after merged-cell resolution). No inheritance from previous rows.
            int? slotIndexCandidate = null;
            if (!string.IsNullOrWhiteSpace(slotRaw))
            {
                if (int.TryParse(slotRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
                    slotIndexCandidate = parsed;
                else
                {
                    // Sometimes SlotIndex is formatted like "Slot 12" or "12-13"; pick first number.
                    var m = Regex.Match(slotRaw, @"\d+");
                    if (m.Success && int.TryParse(m.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var extracted))
                        slotIndexCandidate = extracted;
                }
            }

            // Not a nozzle row: no slot (section titles, blank spacer rows, or truly empty Head/Slot).
            if (slotIndexCandidate is null)
                continue;
            if (string.IsNullOrWhiteSpace(headLabel))
                continue;

            read++;
            try
            {
                if (string.IsNullOrWhiteSpace(machineCode))
                    throw new FormatException("MachineCode is required.");
                
                var slotIndex = slotIndexCandidate.Value;

                // Strict matching: MachineID/MachineCode must already exist in dbo.Machine.
                var machineId = await FindMachineIdStrictAsync(conn, tx, machineCode, ct);
                if (machineId is null)
                    throw new InvalidDataException($"MachineID '{machineCode}' not found (or not unique) in dbo.Machine.");

                var nozzleTypeId = await EnsureNozzleTypeAsync(conn, tx, nozzleCode, ct);
                await UpsertMachineNozzleAsync(conn, tx, machineId.Value, nozzleTypeId, headLabel, slotIndex, quantity: 1, ct);
                ok++;
            }
            catch (Exception ex)
            {
                failed++;
                errors.Add($"Row {r}: {ex.Message}");
            }
        }

        return (read, ok, failed, errors);
    }

    private static async Task<int?> FindMachineIdStrictAsync(SqlConnection conn, SqlTransaction tx, string machineCode, CancellationToken ct)
    {
        await using var cmd = new SqlCommand("""
SELECT COUNT(1) AS Cnt, MIN(MachineId) AS AnyId
FROM dbo.Machine
WHERE MachineCode = @code;
""", conn, tx);
        cmd.Parameters.AddWithValue("@code", machineCode);
        await using var reader = await cmd.ExecuteReaderAsync(ct);
        if (!await reader.ReadAsync(ct)) return null;
        var cnt = reader.GetInt32(0);
        if (cnt != 1) return null;
        return reader.IsDBNull(1) ? null : reader.GetInt32(1);
    }

    private static async Task<string> CopyToTempAsync(string sourcePath, CancellationToken ct)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "SmtLineAllocationUI");
        Directory.CreateDirectory(tempDir);

        var tempPath = Path.Combine(
            tempDir,
            $"{Path.GetFileNameWithoutExtension(sourcePath)}_nozzle_{DateTime.Now:yyyyMMdd_HHmmss_fff}{Path.GetExtension(sourcePath)}"
        );

        await using var src = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        await using var dst = new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
        await src.CopyToAsync(dst, 1024 * 1024, ct);
        return tempPath;
    }

    private static List<IXLWorksheet> ResolveWorksheets(XLWorkbook wb, string? sheetName)
    {
        if (!string.IsNullOrWhiteSpace(sheetName))
            return new List<IXLWorksheet> { wb.Worksheet(sheetName) };

        var list = new List<IXLWorksheet>();
        foreach (var ws in wb.Worksheets)
        {
            var n = ws.Name?.Trim();
            if (string.IsNullOrWhiteSpace(n)) continue;
            if (string.Equals(n, "Nozzle Part Number", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(n, "Line 3", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(n, "Compatibility Report", StringComparison.OrdinalIgnoreCase)) continue;
            list.Add(ws);
        }

        return list;
    }

    private static async Task<HashSet<string>> LoadMachineCodesAsync(SqlConnection conn, SqlTransaction tx, CancellationToken ct)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        await using var cmd = new SqlCommand("SELECT MachineCode FROM dbo.Machine;", conn, tx);
        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            if (reader[0] is string s && !string.IsNullOrWhiteSpace(s))
                set.Add(s.Trim());
        }
        return set;
    }

    private static string? TryResolveMachineCodeFromSheetName(string wsName, HashSet<string> knownMachineCodes)
    {
        if (string.IsNullOrWhiteSpace(wsName)) return null;

        // Prefer exact token containment matches, longest-first.
        // Example: "FZ120 GC6A" contains "GC6A"; "FUZION2 14(GX3)" contains "GX3".
        foreach (var code in knownMachineCodes.OrderByDescending(x => x.Length))
        {
            if (wsName.Contains(code, StringComparison.OrdinalIgnoreCase))
                return code;
        }

        return null;
    }

    private async Task<(int read, int ok, int failed, List<string> errors)> ImportWorksheetMatrix(
        IXLWorksheet ws,
        int machineId,
        SqlConnection conn,
        SqlTransaction tx,
        CancellationToken ct
    )
    {
        var read = 0;
        var ok = 0;
        var failed = 0;
        var errors = new List<string>();

        var wsName = ws.Name ?? "(unnamed)";
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 1;

        // Find changer blocks; if none, parse whole sheet as one block.
        var changerStarts = new List<(int row, int headIndex)>();
        for (var r = 1; r <= lastRow; r++)
        {
            var row = ws.Row(r);
            // Scan a few more columns because captions such as
            // "Changer 4(Rear of Machine)" are often placed in the
            // middle of the sheet (e.g. column F / 6).
            var maxColToScan = Math.Min(15, lastCol);
            for (var c = 1; c <= maxColToScan; c++)
            {
                var text = row.Cell(c).GetString();
                if (TryParseChanger(text, out var head))
                {
                    changerStarts.Add((r, head));
                    break;
                }
            }
        }

        if (changerStarts.Count == 0)
            changerStarts.Add((row: 1, headIndex: 1));

        for (var i = 0; i < changerStarts.Count; i++)
        {
            var startRow = changerStarts[i].row;
            var headIndex = changerStarts[i].headIndex;
            var headLabel = ToChangerLabel(headIndex);
            var endRow = (i + 1 < changerStarts.Count ? changerStarts[i + 1].row - 1 : lastRow);

            var slotMap = new Dictionary<int, int>(); // col -> slotIndex
            for (var r = startRow; r <= endRow; r++)
            {
                ct.ThrowIfCancellationRequested();

                var firstCellText = ws.Row(r).Cell(1).GetString().Trim();
                // GX2 special case: after the two changer blocks there is an "In Line 7"
                // sub‑table that should not be mixed into the changer configuration.
                // Once we hit that area, we stop processing rows for this machine.
                if (firstCellText.StartsWith("In Line", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                if (TryBuildSlotHeaderMap(ws.Row(r), lastCol, out var newMap))
                {
                    slotMap = newMap;
                    // Next row is expected to contain nozzle codes for these slots.
                    var nozzleRowIndex = r + 1;
                    if (nozzleRowIndex > endRow) continue;

                    var nozzleRow = ws.Row(nozzleRowIndex);
                    foreach (var (col, slotIndex) in slotMap)
                    {
                        var nozzleCode = GetNozzleCodeCell(nozzleRow.Cell(col));
                        // null = header/label; empty string = valid "empty nozzle" slot.
                        if (nozzleCode is null) continue;

                        read++;
                        try
                        {
                            var nozzleTypeId = await EnsureNozzleTypeAsync(conn, tx, nozzleCode, ct);
                            await UpsertMachineNozzleAsync(conn, tx, machineId, nozzleTypeId, headLabel, slotIndex, quantity: 1, ct);
                            ok++;
                        }
                        catch (Exception ex)
                        {
                            failed++;
                            errors.Add($"Worksheet '{wsName}' {headLabel} Slot {slotIndex}: {ex.Message}");
                        }
                    }

                    // Skip the nozzle row we just processed.
                    r = nozzleRowIndex;
                }
            }
        }

        return (read, ok, failed, errors);
    }

    private static bool TryParseChanger(string? text, out int headIndex)
    {
        headIndex = 0;
        if (string.IsNullOrWhiteSpace(text)) return false;
        var t = text.Trim();
        if (!t.StartsWith("Changer", StringComparison.OrdinalIgnoreCase)) return false;

        // Support both "Changer 4" and "Changer 4(Rear of Machine)" styles.
        // Extract the first continuous digit segment after the word "Changer".
        var span = t.AsSpan();
        var idx = span.IndexOfAnyInRange('0', '9');
        if (idx < 0) return false;
        var digitSpan = span.Slice(idx);
        var len = 0;
        while (len < digitSpan.Length && char.IsDigit(digitSpan[len])) len++;
        if (len == 0) return false;
        var numSpan = digitSpan.Slice(0, len);
        if (!int.TryParse(numSpan, NumberStyles.Integer, CultureInfo.InvariantCulture, out headIndex)) return false;
        return headIndex > 0;
    }

    private static bool TryBuildSlotHeaderMap(IXLRow row, int lastCol, out Dictionary<int, int> map)
    {
        map = new Dictionary<int, int>();
        var count = 0;
        for (var c = 1; c <= lastCol; c++)
        {
            var cell = row.Cell(c);
            var s = cell.GetString().Trim();
            if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var slot)) continue;
            if (slot <= 0 || slot > 200) continue;
            map[c] = slot;
            count++;
        }

        // A slot header row usually contains multiple slot indices (e.g. >= 4).
        return count >= 4;
    }

    private static string? GetNozzleCodeCell(IXLCell cell)
    {
        // Nozzle codes can be numeric (e.g. 3220) or alphanumeric (e.g. H076, FJ461).
        // We return empty string for "no nozzle" so that the slot is still materialized.
        var s = cell.GetString().Trim();
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        if (IsEmptyNozzlePlaceholder(s)) return string.Empty;
        // Ignore obvious non-data labels.
        if (s.StartsWith("Changer", StringComparison.OrdinalIgnoreCase)) return null;
        if (string.Equals(s, "Name", StringComparison.OrdinalIgnoreCase)) return null;
        // Keep as-is.
        return s;
    }

    private static bool IsEmptyNozzlePlaceholder(string value)
    {
        // Some Excel exports show empty nozzle as a literal placeholder like:
        //   "_EMPTY_NOZZLI_" / "__EMPTY_NOZZLE__" / etc.
        // We treat any token containing both "EMPTY" and "NOZZ" as empty.
        var t = value.Trim();
        if (string.IsNullOrWhiteSpace(t)) return true;
        if (t.Equals("__EMPTY_NOZZLE__", StringComparison.OrdinalIgnoreCase)) return true;

        var u = t.ToUpperInvariant();
        return u.Contains("EMPTY", StringComparison.OrdinalIgnoreCase) && u.Contains("NOZZ", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeNozzleCode(string? nozzleCode)
    {
        if (string.IsNullOrWhiteSpace(nozzleCode)) return string.Empty;
        var t = nozzleCode.Trim();
        if (IsEmptyNozzlePlaceholder(t)) return string.Empty;
        return t;
    }

    private static async Task<int?> FindMachineIdAsync(SqlConnection conn, SqlTransaction tx, string? lineCode, string machineCode, CancellationToken ct)
    {
        if (!string.IsNullOrWhiteSpace(lineCode))
        {
            await using var cmd = new SqlCommand("""
SELECT m.MachineId
FROM dbo.Machine m
JOIN dbo.Line l ON l.LineId = m.LineId
WHERE l.LineCode = @line AND m.MachineCode = @code;
""", conn, tx);
            cmd.Parameters.AddWithValue("@line", lineCode);
            cmd.Parameters.AddWithValue("@code", machineCode);
            var existing = await cmd.ExecuteScalarAsync(ct);
            if (existing is int idByLine) return idByLine;
        }

        await using (var cmd = new SqlCommand("""
SELECT MachineId FROM dbo.Machine WHERE MachineCode = @code;
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@code", machineCode);
            var existing = await cmd.ExecuteScalarAsync(ct);
            if (existing is int id) return id;
        }

        return null;
    }

    private static async Task<int> EnsureNozzleTypeAsync(SqlConnection conn, SqlTransaction tx, string nozzleCode, CancellationToken ct)
    {
        // Map empty/whitespace nozzle to a special placeholder code so that
        // the DB row can exist while UI can render it as blank.
        nozzleCode = NormalizeNozzleCode(nozzleCode);
        var code = string.IsNullOrWhiteSpace(nozzleCode) ? "__EMPTY_NOZZLE__" : nozzleCode;

        await using (var cmd = new SqlCommand("""
SELECT NozzleTypeId FROM dbo.NozzleType WHERE NozzleCode = @code;
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@code", code);
            var existing = await cmd.ExecuteScalarAsync(ct);
            if (existing is int id) return id;
        }

        await using (var cmd = new SqlCommand("""
INSERT INTO dbo.NozzleType (NozzleCode, NozzleModel, Description)
VALUES (@code, @model, @desc);
SELECT CAST(SCOPE_IDENTITY() AS int);
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@code", code);
            cmd.Parameters.AddWithValue("@model", code);
            cmd.Parameters.AddWithValue("@desc", "Imported from NozzleConfig Excel");
            var id = await cmd.ExecuteScalarAsync(ct);
            return (int)id!;
        }
    }

    private static async Task UpsertMachineNozzleAsync(
        SqlConnection conn,
        SqlTransaction tx,
        int machineId,
        int nozzleTypeId,
        string? headLabel,
        int slotIndex,
        int quantity,
        CancellationToken ct
    )
    {
        await using (var cmd = new SqlCommand("""
SELECT MachineNozzleConfigId FROM dbo.MachineNozzleConfig
WHERE MachineId=@mid AND SlotIndex=@slot
  AND ISNULL(HeadLabel,'') = ISNULL(@headLabel,'');
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@mid", machineId);
            cmd.Parameters.AddWithValue("@slot", slotIndex);
            cmd.Parameters.AddWithValue("@headLabel", (object?)headLabel ?? DBNull.Value);
            var existing = await cmd.ExecuteScalarAsync(ct);
            if (existing is int id)
            {
                await using var up = new SqlCommand("""
UPDATE dbo.MachineNozzleConfig
SET NozzleTypeId = @nt,
    HeadLabel    = @headLabel,
    Quantity     = @qty,
    UpdatedAt    = SYSDATETIME()
WHERE MachineNozzleConfigId = @id;
""", conn, tx);
                up.Parameters.AddWithValue("@id", id);
                up.Parameters.AddWithValue("@nt", nozzleTypeId);
                up.Parameters.AddWithValue("@headLabel", (object?)headLabel ?? DBNull.Value);
                up.Parameters.AddWithValue("@qty", quantity);
                await up.ExecuteNonQueryAsync(ct);
                return;
            }
        }

        await using (var cmd = new SqlCommand("""
INSERT INTO dbo.MachineNozzleConfig (MachineId, NozzleTypeId, HeadLabel, SlotIndex, Quantity)
VALUES (@mid, @nt, @headLabel, @slot, @qty);
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@mid", machineId);
            cmd.Parameters.AddWithValue("@nt", nozzleTypeId);
            cmd.Parameters.AddWithValue("@headLabel", (object?)headLabel ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@slot", slotIndex);
            cmd.Parameters.AddWithValue("@qty", quantity);
            await cmd.ExecuteNonQueryAsync(ct);
        }
    }

    private static string ToChangerLabel(int headIndex)
        => $"Changer{headIndex:00}";
}

