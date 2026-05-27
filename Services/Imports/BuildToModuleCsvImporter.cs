using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI.Services.Imports;

/// <summary>
/// Imports SGP_ASSEMBLY_BUILD_TO_MODULE CSV with cleaning:
/// 1) drop rows with empty PCB
/// 2) drop rows where Build Key is not numeric
/// 3) for each Build #, keep only the row with the largest numeric Build Key
/// 4) insert only Build # values not already present in dbo.BuildToModuleMapping (existing rows are kept).
/// </summary>
public sealed class BuildToModuleCsvImporter
{
    private readonly SqlConnectionFactory _connectionFactory;

    public BuildToModuleCsvImporter(SqlConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
    }

    public async Task<ImportResult> ImportAsync(string csvPath, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(csvPath)) throw new ArgumentException("Path cannot be empty.", nameof(csvPath));
        if (!File.Exists(csvPath)) throw new FileNotFoundException("CSV file not found.", csvPath);

        var lineParseErrors = new List<string>();
        var read = 0;
        var skippedEmptyPcb = 0;
        var skippedNonNumericBuildKey = 0;
        var skippedEmptyBuildNumber = 0;

        var cfg = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true,
            BadDataFound = null,
            MissingFieldFound = null,
            HeaderValidated = null,
            TrimOptions = TrimOptions.Trim
        };

        var candidates = new List<CleanRow>();

        using (var stream = File.OpenRead(csvPath))
        using (var reader = new StreamReader(stream))
        using (var csv = new CsvReader(reader, cfg))
        {
            await csv.ReadAsync();
            csv.ReadHeader();

            var hBuild = FindHeader(csv, "Build #", "Build#", "BUILD#");
            var hKey = FindHeader(csv, "Build Key", "BuildKey");
            var hModule = FindHeader(csv, "Module #", "Module#", "MODULE#");
            var hPcb = FindHeader(csv, "PCB");

            if (hBuild is null || hKey is null || hModule is null || hPcb is null)
            {
                lineParseErrors.Add("Missing required column(s). Expected: Build #, Build Key, Module #, PCB.");
                return new ImportResult(Path.GetFileName(csvPath), 0, 0, 1, lineParseErrors);
            }

            while (await csv.ReadAsync())
            {
                ct.ThrowIfCancellationRequested();
                read++;

                try
                {
                    var buildNumber = csv.GetField(hBuild)?.Trim() ?? "";
                    var buildKeyRaw = csv.GetField(hKey)?.Trim() ?? "";
                    var moduleNumber = csv.GetField(hModule)?.Trim() ?? "";
                    var pcb = csv.GetField(hPcb)?.Trim() ?? "";

                    if (string.IsNullOrWhiteSpace(pcb))
                    {
                        skippedEmptyPcb++;
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(buildNumber))
                    {
                        skippedEmptyBuildNumber++;
                        continue;
                    }

                    if (!TryParseNumericBuildKey(buildKeyRaw, out var keyNumeric))
                    {
                        skippedNonNumericBuildKey++;
                        continue;
                    }

                    candidates.Add(new CleanRow(read, buildNumber, buildKeyRaw, keyNumeric, moduleNumber, pcb));
                }
                catch (Exception ex)
                {
                    lineParseErrors.Add($"Line {read + 1}: {ex.Message}");
                }
            }
        }

        // Per Build #, keep single row with maximum numeric Build Key (tie: later line in file wins).
        var deduped = candidates
            .GroupBy(r => r.BuildNumber, StringComparer.Ordinal)
            .Select(g =>
                g.OrderByDescending(x => x.BuildKeyNumeric)
                    .ThenByDescending(x => x.SourceLineNumber)
                    .First())
            .ToList();

        await using var conn = _connectionFactory.Open();
        await using var tx = (SqlTransaction)await conn.BeginTransactionAsync(ct);

        var ok = 0;
        var skippedExistingBuild = 0;
        try
        {
            await EnsureBuildToModuleMappingTableAsync(conn, tx, ct);
            var existingBuilds = await LoadExistingBuildNumbersAsync(conn, tx, ct);

            var toInsert = new List<CleanRow>();
            foreach (var row in deduped)
            {
                if (existingBuilds.Contains(NormalizeBuildNumber(row.BuildNumber)))
                {
                    skippedExistingBuild++;
                    continue;
                }

                toInsert.Add(row);
            }

            foreach (var row in toInsert)
            {
                ct.ThrowIfCancellationRequested();
                await InsertCleanRowAsync(conn, tx, row, Path.GetFileName(csvPath), ct);
                ok++;
            }

            await tx.CommitAsync(ct);
        }
        catch
        {
            await tx.RollbackAsync(ct);
            throw;
        }

        var summary =
            $"Rows read: {read}\n" +
            $"Inserted (new Build # only, after dedupe): {ok}\n" +
            $"Skipped (Build # already in database): {skippedExistingBuild}\n" +
            $"Skipped (empty PCB): {skippedEmptyPcb}\n" +
            $"Skipped (non-numeric Build Key): {skippedNonNumericBuildKey}\n" +
            $"Skipped (empty Build #): {skippedEmptyBuildNumber}";

        var messages = new List<string> { summary };
        messages.AddRange(lineParseErrors);

        return new ImportResult(Path.GetFileName(csvPath), read, ok, lineParseErrors.Count, messages);
    }

    private static string? FindHeader(CsvReader csv, params string[] names)
    {
        if (csv.HeaderRecord is null) return null;
        foreach (var n in names)
        {
            foreach (var h in csv.HeaderRecord)
            {
                if (string.Equals(h?.Trim(), n, StringComparison.OrdinalIgnoreCase))
                    return h;
            }
        }

        return null;
    }

    /// <summary>Accepts integers and decimals; rejects empty, letters-only, hex-like junk.</summary>
    private static bool TryParseNumericBuildKey(string raw, out decimal value)
    {
        value = 0m;
        if (string.IsNullOrWhiteSpace(raw)) return false;
        return decimal.TryParse(raw, NumberStyles.Number, CultureInfo.InvariantCulture, out value);
    }

    private sealed record CleanRow(
        int SourceLineNumber,
        string BuildNumber,
        string BuildKeyText,
        decimal BuildKeyNumeric,
        string ModuleNumber,
        string Pcb);

    private static async Task EnsureBuildToModuleMappingTableAsync(SqlConnection conn, SqlTransaction tx, CancellationToken ct)
    {
        // Defensive: some users may run an older build output that still has an old schema.sql copy.
        // Ensure the required table exists even if schema upgrade didn't run.
        await using var cmd = new SqlCommand("""
IF OBJECT_ID(N'dbo.BuildToModuleMapping', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.BuildToModuleMapping (
        BuildToModuleMappingId INT IDENTITY(1,1) PRIMARY KEY,
        BuildNumber       NVARCHAR(50)  NULL,
        BuildKey          NVARCHAR(50)  NULL,
        ModuleNumber      NVARCHAR(50)  NULL,
        ModuleKey         NVARCHAR(50)  NULL,
        PCB               NVARCHAR(50)  NULL,
        ModulesPerPanel   INT           NULL,
        OptelCreateDate   DATETIME2(0)  NULL,
        LastChanged       DATETIME2(0)  NULL,
        SapBuildDescription  NVARCHAR(200) NULL,
        SapModuleDescription NVARCHAR(200) NULL,
        OptelDescription     NVARCHAR(500) NULL,
        SourceFileName    NVARCHAR(260) NULL,
        CreatedAt         DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME()
    );
END;
""", conn, tx);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    private static string NormalizeBuildNumber(string buildNumber) =>
        buildNumber.Trim().ToUpperInvariant();

    private static async Task<HashSet<string>> LoadExistingBuildNumbersAsync(
        SqlConnection conn,
        SqlTransaction tx,
        CancellationToken ct)
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        await using var cmd = new SqlCommand("""
SELECT DISTINCT BuildNumber
FROM dbo.BuildToModuleMapping
WHERE BuildNumber IS NOT NULL AND LTRIM(RTRIM(BuildNumber)) <> N'';
""", conn, tx);
        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            if (reader.IsDBNull(0)) continue;
            var bn = reader.GetString(0);
            if (!string.IsNullOrWhiteSpace(bn))
                set.Add(NormalizeBuildNumber(bn));
        }

        return set;
    }

    private static async Task InsertCleanRowAsync(
        SqlConnection conn,
        SqlTransaction tx,
        CleanRow row,
        string sourceFileName,
        CancellationToken ct)
    {
        await using var cmd = new SqlCommand("""
INSERT INTO dbo.BuildToModuleMapping
    (BuildNumber, BuildKey, ModuleNumber, ModuleKey, PCB, ModulesPerPanel, OptelCreateDate, LastChanged,
     SapBuildDescription, SapModuleDescription, OptelDescription, SourceFileName)
VALUES
    (@buildNumber, @buildKey, @moduleNumber, NULL, @pcb, NULL, NULL, NULL,
     NULL, NULL, NULL, @src);
""", conn, tx);

        cmd.Parameters.AddWithValue("@buildNumber", row.BuildNumber);
        cmd.Parameters.AddWithValue("@buildKey", row.BuildKeyText);
        cmd.Parameters.AddWithValue("@moduleNumber", string.IsNullOrWhiteSpace(row.ModuleNumber) ? (object)DBNull.Value : row.ModuleNumber);
        cmd.Parameters.AddWithValue("@pcb", row.Pcb);
        cmd.Parameters.AddWithValue("@src", sourceFileName);

        await cmd.ExecuteNonQueryAsync(ct);
    }
}
