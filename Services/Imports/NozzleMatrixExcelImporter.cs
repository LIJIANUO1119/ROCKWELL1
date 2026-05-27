using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using SmtLineAllocationUI.DataAccess;
using SmtLineAllocationUI.Services;

namespace SmtLineAllocationUI.Services.Imports;

/// <summary>
/// Imports wide-format workbook: column A = Machine ID, columns B+ = nozzle type codes with quantities.
/// </summary>
public sealed class NozzleMatrixExcelImporter
{
    private readonly SqlConnectionFactory _connectionFactory;

    public NozzleMatrixExcelImporter(SqlConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
    }

    public async Task<ImportResult> ImportAsync(string xlsxPath, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(xlsxPath)) throw new ArgumentException("Path cannot be empty.", nameof(xlsxPath));
        if (!File.Exists(xlsxPath)) throw new FileNotFoundException("Excel file not found.", xlsxPath);

        var errors = new List<string>();
        var read = 0;
        var ok = 0;

        string? tempCopyPath = null;
        try
        {
            tempCopyPath = await CopyToTempAsync(xlsxPath, ct);
            using var wb = new XLWorkbook(tempCopyPath);

            List<string> nozzleCodes = new();
            List<(string MachineCode, IReadOnlyDictionary<string, int> Quantities)> rows = new();
            string? parseError = null;

            var sheets = wb.Worksheets.Where(w => w.Visibility == XLWorksheetVisibility.Visible).ToList();
            if (sheets.Count == 0) sheets.Add(wb.Worksheet(1));

            var parsed = false;
            foreach (var ws in sheets)
            {
                if (!TryParseWorksheet(ws, out nozzleCodes, out rows, out parseError)) continue;
                parsed = true;
                break;
            }

            if (!parsed)
            {
                errors.Add(parseError ?? "Could not parse worksheet.");
                return new ImportResult(Path.GetFileName(xlsxPath), 0, 0, 1, errors);
            }

            read = rows.Count;
            var deduped = DeduplicateMachineRows(rows, errors);
            await MachineNozzleMatrixService.ReplaceAllFromImportAsync(_connectionFactory, nozzleCodes, deduped);
            ok = deduped.Count;

            return new ImportResult(Path.GetFileName(xlsxPath), read, ok, Math.Max(0, read - ok), errors);
        }
        catch (Exception ex)
        {
            errors.Add(ex.Message);
            return new ImportResult(Path.GetFileName(xlsxPath), read, ok, Math.Max(1, read - ok), errors);
        }
        finally
        {
            if (tempCopyPath is not null)
            {
                try { File.Delete(tempCopyPath); } catch { /* ignore */ }
            }
        }
    }

    internal static bool TryParseWorksheet(
        IXLWorksheet ws,
        out List<string> nozzleCodes,
        out List<(string MachineCode, IReadOnlyDictionary<string, int> Quantities)> rows,
        out string? error)
    {
        nozzleCodes = new List<string>();
        rows = new List<(string, IReadOnlyDictionary<string, int>)>();
        error = null;

        if (!TryFindMatrixHeader(ws, out var headerRow, out nozzleCodes, out var findError))
        {
            error = findError;
            return false;
        }

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? headerRow;
        for (var r = headerRow + 1; r <= lastRow; r++)
        {
            var machineCode = GetCellText(ws, r, 1);
            if (string.IsNullOrWhiteSpace(machineCode)) continue;
            if (IsMachineIdColumnHeader(machineCode)) continue;

            var quantities = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < nozzleCodes.Count; i++)
            {
                quantities[nozzleCodes[i]] = ParseCellQuantity(ws.Cell(r, i + 2));
            }

            rows.Add((machineCode, quantities));
        }

        if (rows.Count == 0)
        {
            error = "No data rows found under the header.";
            return false;
        }

        return true;
    }

    /// <summary>Excel may list the same Machine ID on multiple rows; keep the last row per ID.</summary>
    internal static List<(string MachineCode, IReadOnlyDictionary<string, int> Quantities)> DeduplicateMachineRows(
        IReadOnlyList<(string MachineCode, IReadOnlyDictionary<string, int> Quantities)> rows,
        IList<string> warnings)
    {
        var map = new Dictionary<string, (string MachineCode, Dictionary<string, int> Quantities)>(
            StringComparer.OrdinalIgnoreCase);

        foreach (var row in rows)
        {
            var code = row.MachineCode.Trim();
            if (string.IsNullOrWhiteSpace(code)) continue;

            if (map.ContainsKey(code))
            {
                warnings.Add(
                    $"Duplicate Machine ID \"{code}\" in the file; the last row was used.");
            }

            map[code] = (code, new Dictionary<string, int>(row.Quantities, StringComparer.OrdinalIgnoreCase));
        }

        return map.Values
            .Select(v => (v.MachineCode, (IReadOnlyDictionary<string, int>)v.Quantities))
            .ToList();
    }

    /// <summary>
    /// Finds the row with nozzle type codes (1220, 1240, …) in columns B+.
    /// Column A may be "Machine ID" or empty when Excel uses merged cells above.
    /// </summary>
    internal static bool TryFindMatrixHeader(
        IXLWorksheet ws,
        out int headerRow,
        out List<string> nozzleCodes,
        out string? error)
    {
        headerRow = -1;
        nozzleCodes = new List<string>();
        error = null;

        var scanRows = Math.Min(ws.LastRowUsed()?.RowNumber() ?? 50, 50);
        var bestCount = 0;
        var bestHasMachineIdLabel = false;

        for (var r = 1; r <= scanRows; r++)
        {
            var codes = ReadNozzleTypeHeaders(ws, r);
            if (codes.Count < 2) continue;
            if (!HasMachineDataRowsBelow(ws, r)) continue;

            var colA = GetCellText(ws, r, 1);
            var hasMachineIdLabel = IsMachineIdColumnHeader(colA);

            var better = codes.Count > bestCount
                         || (codes.Count == bestCount && hasMachineIdLabel && !bestHasMachineIdLabel);

            if (!better) continue;

            bestCount = codes.Count;
            bestHasMachineIdLabel = hasMachineIdLabel;
            headerRow = r;
            nozzleCodes = codes;
        }

        if (headerRow < 1 || nozzleCodes.Count == 0)
        {
            error =
                "Could not find a header row with nozzle type columns (1220, 1240, …). " +
                "Expected column A = Machine ID on that row (or merged), with nozzle codes in columns B onward.";
            return false;
        }

        return true;
    }

    private static bool HasMachineDataRowsBelow(IXLWorksheet ws, int headerRow)
    {
        var lastRow = Math.Min(ws.LastRowUsed()?.RowNumber() ?? headerRow + 1, headerRow + 15);
        for (var r = headerRow + 1; r <= lastRow; r++)
        {
            var a = GetCellText(ws, r, 1);
            if (LooksLikeMachineCode(a)) return true;
        }

        return false;
    }

    private static bool LooksLikeMachineCode(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var t = text.Trim();
        if (IsMachineIdColumnHeader(t)) return false;
        if (IsNozzleTypeBannerText(t)) return false;
        if (t.Contains("machine", StringComparison.OrdinalIgnoreCase)) return false;
        return t.Length <= 40;
    }

    private static bool IsMachineIdColumnHeader(string text)
    {
        var t = text.Trim();
        if (string.IsNullOrEmpty(t)) return false;
        if (t.Length > 40) return false;
        if (t.Contains("configuration", StringComparison.OrdinalIgnoreCase)) return false;

        var compact = Regex.Replace(t, @"[\s_\-]+", "", RegexOptions.CultureInvariant);
        if (compact.Equals("MachineID", StringComparison.OrdinalIgnoreCase)) return true;
        if (compact.Equals("MachineNo", StringComparison.OrdinalIgnoreCase)) return true;

        return t.Contains("machine", StringComparison.OrdinalIgnoreCase)
               && t.Contains("id", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNozzleTypeColumnHeader(string header)
    {
        if (string.IsNullOrWhiteSpace(header)) return false;
        var h = header.Trim();
        if (IsNozzleTypeBannerText(h)) return false;
        if (h.Length > 12) return false;
        return Regex.IsMatch(h, @"^[0-9A-Za-z]{2,10}$");
    }

    private static string GetCellText(IXLWorksheet ws, int row, int col)
    {
        var cell = ws.Cell(row, col);
        if (cell.IsEmpty()) return "";

        var s = cell.GetString();
        if (!string.IsNullOrWhiteSpace(s)) return s.Trim();

        if (cell.TryGetValue(out string str) && !string.IsNullOrWhiteSpace(str))
            return str.Trim();

        return cell.Value.ToString()?.Trim() ?? "";
    }

    private static bool IsNozzleTypeBannerText(string text)
    {
        return text.Contains("nozzle", StringComparison.OrdinalIgnoreCase)
               && text.Contains("type", StringComparison.OrdinalIgnoreCase)
               && (text.Contains("quantity", StringComparison.OrdinalIgnoreCase) || text.Contains('&'));
    }

    private static List<string> ReadNozzleTypeHeaders(IXLWorksheet ws, int headerRow)
    {
        var codes = new List<string>();
        var lastCol = GetLastNonEmptyColumnInRow(ws, headerRow, startCol: 2);
        for (var c = 2; c <= lastCol; c++)
        {
            var header = GetCellText(ws, headerRow, c);
            if (IsNozzleTypeColumnHeader(header))
                codes.Add(header);
        }

        return codes;
    }

    private static int GetLastNonEmptyColumnInRow(IXLWorksheet ws, int row, int startCol)
    {
        var maxScan = Math.Max(ws.LastColumnUsed()?.ColumnNumber() ?? startCol, 60);
        var last = startCol - 1;
        for (var c = startCol; c <= maxScan; c++)
        {
            if (!string.IsNullOrWhiteSpace(GetCellText(ws, row, c)))
                last = c;
        }

        return Math.Max(last, startCol);
    }

    private static int ParseCellQuantity(IXLCell cell)
    {
        if (cell.IsEmpty()) return 0;
        if (cell.TryGetValue(out int i)) return i;
        if (cell.TryGetValue(out double d)) return (int)Math.Round(d);
        var s = cell.GetString().Trim();
        if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)) return n;
        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d2)) return (int)Math.Round(d2);
        return 0;
    }

    private static async Task<string> CopyToTempAsync(string sourcePath, CancellationToken ct)
    {
        var dest = Path.Combine(Path.GetTempPath(), $"nozzle_matrix_{Guid.NewGuid():N}.xlsx");
        await using var src = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        await using var dst = new FileStream(dest, FileMode.Create, FileAccess.Write, FileShare.None);
        await src.CopyToAsync(dst, ct);
        return dest;
    }
}
