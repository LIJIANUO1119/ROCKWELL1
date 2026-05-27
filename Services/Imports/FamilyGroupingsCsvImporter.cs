using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;
using SmtLineAllocationUI.Services;

namespace SmtLineAllocationUI.Services.Imports;

/// <summary>
/// Imports SGP_FAMILY_GROUPINGS_DETAIL (CSV or Excel).
/// Persists ASY, PCB, Family, Family#, Top/Bot priorities (1–3), cycle times, and # Circuits.
/// Priority 2/3 columns are stored as NULL when absent from the file.
/// </summary>
public sealed class FamilyGroupingsCsvImporter
{
    private readonly SqlConnectionFactory _connectionFactory;

    public FamilyGroupingsCsvImporter(SqlConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
    }

    public async Task<ImportResult> ImportAsync(string filePath, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("Path cannot be empty.", nameof(filePath));
        if (!File.Exists(filePath)) throw new FileNotFoundException("File not found.", filePath);

        var ext = Path.GetExtension(filePath);
        List<ParsedRow> rows;
        List<string> errors;
        int read;

        if (ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            (rows, errors, read) = ReadExcel(filePath);
        }
        else if (ext.Equals(".csv", StringComparison.OrdinalIgnoreCase))
        {
            (rows, errors, read) = await ReadCsvAsync(filePath, ct);
        }
        else
        {
            throw new NotSupportedException("Supported formats: .csv, .xlsx, .xlsm.");
        }

        if (rows.Count == 0 && errors.Count > 0)
            return new ImportResult(Path.GetFileName(filePath), read, 0, errors.Count, errors);

        var dedupedByAsy = rows
            .Where(r => !string.IsNullOrWhiteSpace(r.AssemblyNo))
            .GroupBy(r => NormalizeAssembly(r.AssemblyNo), StringComparer.Ordinal)
            .Select(g => g.Last())
            .ToList();

        await using var conn = _connectionFactory.Open();
        await using var tx = (SqlTransaction)await conn.BeginTransactionAsync(ct);

        var inserted = 0;
        var updated = 0;
        var skippedEmptyAsy = rows.Count(r => string.IsNullOrWhiteSpace(r.AssemblyNo));
        try
        {
            await FamilyGroupingViewData.EnsureTableAsync(conn, tx, ct);
            var idByAsy = await LoadAssemblyIdMapAsync(conn, tx, ct);

            var sourceFile = Path.GetFileName(filePath);
            foreach (var row in dedupedByAsy)
            {
                ct.ThrowIfCancellationRequested();
                var key = NormalizeAssembly(row.AssemblyNo);
                if (idByAsy.TryGetValue(key, out var existingId))
                {
                    await UpdateRowAsync(conn, tx, existingId, row, sourceFile, ct);
                    updated++;
                }
                else
                {
                    await InsertRowAsync(conn, tx, row, sourceFile, ct);
                    inserted++;
                    idByAsy[key] = -1;
                }
            }

            await tx.CommitAsync(ct);
        }
        catch
        {
            await tx.RollbackAsync(ct);
            throw;
        }

        var ok = inserted + updated;
        var summary =
            $"Rows read: {read}\n" +
            $"Inserted: {inserted}\n" +
            $"Updated (matching ASY): {updated}\n" +
            $"Skipped (empty ASY): {skippedEmptyAsy}";
        var messages = new List<string> { summary };
        messages.AddRange(errors);

        return new ImportResult(Path.GetFileName(filePath), read, ok, errors.Count, messages);
    }

    private static async Task<(List<ParsedRow> Rows, List<string> Errors, int Read)> ReadCsvAsync(string csvPath, CancellationToken ct)
    {
        var errors = new List<string>();
        var read = 0;
        var rows = new List<ParsedRow>();

        var cfg = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true,
            BadDataFound = null,
            MissingFieldFound = null,
            HeaderValidated = null,
            TrimOptions = TrimOptions.Trim
        };

        using var stream = File.OpenRead(csvPath);
        using var reader = new StreamReader(stream);
        using var csv = new CsvReader(reader, cfg);

        await csv.ReadAsync();
        csv.ReadHeader();

        var map = BuildColumnMap(csv.HeaderRecord ?? Array.Empty<string>());
        var missing = RequiredColumnsMissing(map);
        if (missing.Count > 0)
        {
            errors.Add("Missing required column(s): " + string.Join(", ", missing));
            return (rows, errors, 0);
        }

        while (await csv.ReadAsync())
        {
            ct.ThrowIfCancellationRequested();
            read++;
            try
            {
                rows.Add(ParseRow(
                    i => csv.GetField(i),
                    map,
                    read + 1));
            }
            catch (Exception ex)
            {
                errors.Add($"Line {read + 1}: {ex.Message}");
            }
        }

        return (rows, errors, read);
    }

    private static (List<ParsedRow> Rows, List<string> Errors, int Read) ReadExcel(string xlsxPath)
    {
        var errors = new List<string>();
        var rows = new List<ParsedRow>();
        var read = 0;

        using var wb = new XLWorkbook(xlsxPath);
        var ws = wb.Worksheets.FirstOrDefault(w => w.Visibility == XLWorksheetVisibility.Visible)
                 ?? wb.Worksheets.First();
        var range = ws.RangeUsed();
        if (range is null)
            return (rows, errors, 0);

        var firstRow = range.FirstRow().RowNumber();
        var lastRow = range.LastRow().RowNumber();
        var firstCol = range.FirstColumn().ColumnNumber();
        var lastCol = range.LastColumn().ColumnNumber();

        var headers = new List<string>();
        for (var c = firstCol; c <= lastCol; c++)
            headers.Add(ws.Cell(firstRow, c).GetFormattedString().Trim());

        var map = BuildColumnMap(headers);
        var missing = RequiredColumnsMissing(map);
        if (missing.Count > 0)
        {
            errors.Add("Missing required column(s): " + string.Join(", ", missing));
            return (rows, errors, 0);
        }

        for (var r = firstRow + 1; r <= lastRow; r++)
        {
            read++;
            var rowNum = r;
            try
            {
                string? Field(int col) =>
                    col < 0 ? null : ws.Cell(r, col + firstCol).GetFormattedString().Trim();

                rows.Add(ParseRow(
                    i => i < 0 ? null : Field(i),
                    map,
                    rowNum));
            }
            catch (Exception ex)
            {
                errors.Add($"Row {rowNum}: {ex.Message}");
            }
        }

        return (rows, errors, read);
    }

    private static List<string> RequiredColumnsMissing(ColumnMap map)
    {
        var missing = new List<string>();
        if (map.AssemblyNo < 0) missing.Add("ASY (Asy)");
        if (map.Pcb < 0) missing.Add("PCB");
        if (map.Family < 0) missing.Add("Family");
        if (map.FamilyNumber < 0) missing.Add("Family#");
        if (map.TopPriority1 < 0) missing.Add("Top Priority 1");
        if (map.TopCycleTime < 0) missing.Add("Top CycleTime");
        if (map.BotPriority1 < 0) missing.Add("Bot Priority 1");
        if (map.BotCycleTime < 0) missing.Add("Bot CycleTime");
        if (map.CircuitCount < 0) missing.Add("# Circuits");
        return missing;
    }

    private static ParsedRow ParseRow(Func<int, string?> getField, ColumnMap map, int sourceRow)
    {
        static string? T(Func<int, string?> get, int idx) =>
            idx < 0 ? null : NullIfEmpty(get(idx));

        return new ParsedRow(
            T(getField, map.AssemblyNo) ?? "",
            T(getField, map.Pcb),
            T(getField, map.Family),
            T(getField, map.FamilyNumber),
            T(getField, map.TopPriority1),
            T(getField, map.TopPriority2),
            T(getField, map.TopPriority3),
            ParseDecimal(T(getField, map.TopCycleTime)),
            T(getField, map.BotPriority1),
            T(getField, map.BotPriority2),
            T(getField, map.BotPriority3),
            ParseDecimal(T(getField, map.BotCycleTime)),
            ParseInt(T(getField, map.CircuitCount)),
            sourceRow);
    }

    private static ColumnMap BuildColumnMap(IEnumerable<string> headers)
    {
        var list = headers.ToList();
        int Idx(params string[] names)
        {
            for (var i = 0; i < list.Count; i++)
            {
                var h = list[i];
                foreach (var n in names)
                {
                    if (HeaderEquals(h, n)) return i;
                }
            }

            return -1;
        }

        return new ColumnMap(
            Idx("Asy", "ASY", "AssemblyNo", "Assembly No"),
            Idx("PCB"),
            Idx("Family"),
            Idx("Family #", "Family#", "Family No", "Family No."),
            Idx("Top Priority 1", "Top Priority1"),
            Idx("Top Priority 2", "Top Priority2"),
            Idx("Top Priority 3", "Top Priority3"),
            Idx("Top CycleTime", "Top Cycle Time"),
            Idx("Bot Priority 1", "Bot Priority1"),
            Idx("Bot Priority 2", "Bot Priority2"),
            Idx("Bot Priority 3", "Bot Priority3"),
            Idx("Bot CycleTime", "Bot  CycleTime", "Bot Cycle Time"),
            Idx("# Circuits", "Circuits", "#Circuits"));
    }

    private static bool HeaderEquals(string? header, string expected) =>
        string.Equals(NormalizeHeader(header), NormalizeHeader(expected), StringComparison.OrdinalIgnoreCase);

    private static string NormalizeHeader(string? header) =>
        string.IsNullOrWhiteSpace(header) ? "" : Regex.Replace(header.Trim(), @"\s+", " ");

    private static string? NullIfEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    private static decimal? ParseDecimal(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : null;
    }

    private static int? ParseInt(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;
    }

    private sealed record ParsedRow(
        string AssemblyNo,
        string? Pcb,
        string? FamilyName,
        string? FamilyNumber,
        string? TopPriority1,
        string? TopPriority2,
        string? TopPriority3,
        decimal? TopCycleTimeSec,
        string? BotPriority1,
        string? BotPriority2,
        string? BotPriority3,
        decimal? BotCycleTimeSec,
        int? CircuitCount,
        int SourceRow);

    private sealed record ColumnMap(
        int AssemblyNo,
        int Pcb,
        int Family,
        int FamilyNumber,
        int TopPriority1,
        int TopPriority2,
        int TopPriority3,
        int TopCycleTime,
        int BotPriority1,
        int BotPriority2,
        int BotPriority3,
        int BotCycleTime,
        int CircuitCount);

    private static string NormalizeAssembly(string assemblyNo) =>
        assemblyNo.Trim().ToUpperInvariant();

    private static async Task<Dictionary<string, int>> LoadAssemblyIdMapAsync(
        SqlConnection conn,
        SqlTransaction tx,
        CancellationToken ct)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        await using var cmd = new SqlCommand("""
SELECT FamilyGroupingDetailId, AssemblyNo
FROM dbo.FamilyGroupingDetail
WHERE AssemblyNo IS NOT NULL AND LTRIM(RTRIM(AssemblyNo)) <> N'';
""", conn, tx);
        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            var id = reader.GetInt32(0);
            var asy = reader.GetString(1);
            if (string.IsNullOrWhiteSpace(asy)) continue;
            map[NormalizeAssembly(asy)] = id;
        }

        return map;
    }

    private static async Task UpdateRowAsync(
        SqlConnection conn,
        SqlTransaction tx,
        int id,
        ParsedRow row,
        string sourceFileName,
        CancellationToken ct)
    {
        await using var cmd = new SqlCommand("""
UPDATE dbo.FamilyGroupingDetail
SET AssemblyNo = @asy,
    Pcb = @pcb,
    FamilyName = @family,
    FamilyNumber = @familyNum,
    TopPriority1 = @tp1,
    TopPriority2 = @tp2,
    TopPriority3 = @tp3,
    TopCycleTimeSec = @topCt,
    BotPriority1 = @bp1,
    BotPriority2 = @bp2,
    BotPriority3 = @bp3,
    BotCycleTimeSec = @botCt,
    CircuitCount = @circuits,
    SourceFileName = @src
WHERE FamilyGroupingDetailId = @id;
""", conn, tx);

        cmd.Parameters.AddWithValue("@id", id);
        BindRowParameters(cmd, row, sourceFileName);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    private static async Task InsertRowAsync(
        SqlConnection conn,
        SqlTransaction tx,
        ParsedRow row,
        string sourceFileName,
        CancellationToken ct)
    {
        await using var cmd = new SqlCommand("""
INSERT INTO dbo.FamilyGroupingDetail
    (AssemblyNo, Pcb, FamilyName, FamilyNumber,
     TopPriority1, TopPriority2, TopPriority3, TopCycleTimeSec,
     BotPriority1, BotPriority2, BotPriority3, BotCycleTimeSec,
     CircuitCount, SourceFileName)
VALUES
    (@asy, @pcb, @family, @familyNum,
     @tp1, @tp2, @tp3, @topCt,
     @bp1, @bp2, @bp3, @botCt,
     @circuits, @src);
""", conn, tx);

        BindRowParameters(cmd, row, sourceFileName);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    private static void BindRowParameters(SqlCommand cmd, ParsedRow row, string sourceFileName)
    {
        cmd.Parameters.AddWithValue("@asy", row.AssemblyNo.Trim());
        cmd.Parameters.AddWithValue("@pcb", (object?)row.Pcb ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@family", (object?)row.FamilyName ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@familyNum", (object?)row.FamilyNumber ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@tp1", (object?)row.TopPriority1 ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@tp2", (object?)row.TopPriority2 ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@tp3", (object?)row.TopPriority3 ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@topCt", (object?)row.TopCycleTimeSec ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@bp1", (object?)row.BotPriority1 ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@bp2", (object?)row.BotPriority2 ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@bp3", (object?)row.BotPriority3 ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@botCt", (object?)row.BotCycleTimeSec ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@circuits", (object?)row.CircuitCount ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@src", sourceFileName);
    }
}
