using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI.Services.Imports;

public sealed class MachineCycletimeCsvImporter
{
    private readonly SqlConnectionFactory _connectionFactory;

    public MachineCycletimeCsvImporter(SqlConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
    }

    public async Task<ImportResult> ImportAsync(string csvPath, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(csvPath)) throw new ArgumentException("Path cannot be empty.", nameof(csvPath));
        if (!File.Exists(csvPath)) throw new FileNotFoundException("CSV file not found.", csvPath);

        var parseErrors = new List<string>();
        var read = 0;
        var skippedBoardsProcessed = 0;
        var skippedInvalidPanelEndTime = 0;

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

        // 1) Read + filter
        var candidates = new List<Row>();
        await csv.ReadAsync();
        csv.ReadHeader();

        while (await csv.ReadAsync())
        {
            ct.ThrowIfCancellationRequested();
            read++;

            try
            {
                var lineCode = csv.GetField("LINE")?.Trim() ?? "";
                var assemblyNo = csv.GetField("ASSEMBLYNO")?.Trim() ?? "";
                var side = csv.GetField("Side")?.Trim() ?? "";
                var machineName = csv.GetField("MACHINENAME")?.Trim() ?? "";

                var boardsProcessedText = csv.GetField("BOARDSPROCESSED")?.Trim() ?? "";
                if (!int.TryParse(boardsProcessedText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var boardsProcessed) || boardsProcessed < 10)
                {
                    skippedBoardsProcessed++;
                    continue;
                }

                var panelEndTimeRaw = csv.GetField("PANELENDTIME")?.Trim() ?? "";
                if (!TryParsePanelEndTime(panelEndTimeRaw, out var panelEndTime))
                {
                    skippedInvalidPanelEndTime++;
                    continue;
                }

                var minText = csv.GetField("MACHINE_MIN_CYCLETIME")?.Trim() ?? "";
                var medText = csv.GetField("MACHINE_MEDIUM_CYCLETIME")?.Trim() ?? "";
                var curText = csv.GetField("ADM_CURRENT_CYCLETIME")?.Trim() ?? "";

                _ = decimal.TryParse(minText, NumberStyles.Any, CultureInfo.InvariantCulture, out var minSec);
                _ = decimal.TryParse(medText, NumberStyles.Any, CultureInfo.InvariantCulture, out var medSec);
                _ = decimal.TryParse(curText, NumberStyles.Any, CultureInfo.InvariantCulture, out var curSec);

                if (string.IsNullOrWhiteSpace(lineCode) || string.IsNullOrWhiteSpace(assemblyNo) || string.IsNullOrWhiteSpace(machineName))
                    throw new InvalidDataException("Missing required field(s): LINE / ASSEMBLYNO / MACHINENAME.");

                candidates.Add(new Row(read, lineCode, assemblyNo, side, machineName, panelEndTimeRaw, panelEndTime, boardsProcessed, minSec, medSec, curSec));
            }
            catch (Exception ex)
            {
                parseErrors.Add($"Line {read + 1}: {ex.Message}");
            }
        }

        // 2) Dedupe: same (LINE, ASSEMBLYNO, Side, MACHINENAME) keep row with latest PANELENDTIME.
        var deduped = candidates
            .GroupBy(r => (r.LineCode, r.AssemblyNo, r.Side, r.MachineName), RowKeyComparer.Ordinal)
            .Select(g => g.OrderByDescending(x => x.PanelEndTime).ThenByDescending(x => x.SourceLineNumber).First())
            .ToList();

        // 3) Append to snapshot table (existing rows are kept).
        await using var conn = _connectionFactory.Open();
        await using var tx = (SqlTransaction)await conn.BeginTransactionAsync(ct);

        var ok = 0;
        try
        {
            await EnsureSnapshotTableAsync(conn, tx, ct);

            foreach (var row in deduped)
            {
                ct.ThrowIfCancellationRequested();
                await InsertSnapshotAsync(conn, tx, row, Path.GetFileName(csvPath), ct);
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
            $"Inserted (appended after filter/dedupe): {ok}\n" +
            $"Skipped (BOARDSPROCESSED non-numeric or <10): {skippedBoardsProcessed}\n" +
            $"Skipped (PANELENDTIME invalid): {skippedInvalidPanelEndTime}";
        var messages = new List<string> { summary };
        messages.AddRange(parseErrors);

        return new ImportResult(Path.GetFileName(csvPath), read, ok, parseErrors.Count, messages);
    }

    private sealed record Row(
        int SourceLineNumber,
        string LineCode,
        string AssemblyNo,
        string Side,
        string MachineName,
        string PanelEndTimeText,
        DateTime PanelEndTime,
        int BoardsProcessed,
        decimal MachineMinSec,
        decimal MachineMediumSec,
        decimal AdmCurrentSec
    );

    private sealed class RowKeyComparer : IEqualityComparer<(string LineCode, string AssemblyNo, string Side, string MachineName)>
    {
        public static readonly RowKeyComparer Ordinal = new();

        public bool Equals((string LineCode, string AssemblyNo, string Side, string MachineName) x, (string LineCode, string AssemblyNo, string Side, string MachineName) y)
            => string.Equals(x.LineCode, y.LineCode, StringComparison.Ordinal)
               && string.Equals(x.AssemblyNo, y.AssemblyNo, StringComparison.Ordinal)
               && string.Equals(x.Side, y.Side, StringComparison.Ordinal)
               && string.Equals(x.MachineName, y.MachineName, StringComparison.Ordinal);

        public int GetHashCode((string LineCode, string AssemblyNo, string Side, string MachineName) obj)
            => HashCode.Combine(obj.LineCode, obj.AssemblyNo, obj.Side, obj.MachineName);
    }

    private static async Task EnsureSnapshotTableAsync(SqlConnection conn, SqlTransaction tx, CancellationToken ct)
    {
        await using var cmd = new SqlCommand("""
IF OBJECT_ID(N'dbo.MachineCycleTimeSnapshot', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MachineCycleTimeSnapshot (
        MachineCycleTimeSnapshotId INT IDENTITY(1,1) PRIMARY KEY,
        LineCode          NVARCHAR(50)  NOT NULL,
        AssemblyNo        NVARCHAR(100) NOT NULL,
        Side              NVARCHAR(20)  NULL,
        MachineName       NVARCHAR(100) NOT NULL,
        PanelEndTime      DATETIME2(0)  NULL,
        BoardsProcessed   INT           NULL,
        MachineMinCycleTimeSec     DECIMAL(10,3) NULL,
        MachineMediumCycleTimeSec  DECIMAL(10,3) NULL,
        AdmCurrentCycleTimeSec     DECIMAL(10,3) NULL,
        SourceFileName    NVARCHAR(260) NULL,
        CreatedAt         DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME()
    );

    CREATE INDEX IX_MachineCycleTimeSnapshot_Key
    ON dbo.MachineCycleTimeSnapshot(LineCode, AssemblyNo, Side, MachineName);
END;

-- Idempotent upgrade: add PanelEndTime column if missing.
IF OBJECT_ID(N'dbo.MachineCycleTimeSnapshot', N'U') IS NOT NULL
   AND COL_LENGTH('dbo.MachineCycleTimeSnapshot', 'PanelEndTime') IS NULL
BEGIN
    ALTER TABLE dbo.MachineCycleTimeSnapshot ADD PanelEndTime DATETIME2(0) NULL;
END;

-- Idempotent upgrade: remove WorkOrderNo column (no longer used).
IF OBJECT_ID(N'dbo.MachineCycleTimeSnapshot', N'U') IS NOT NULL
   AND COL_LENGTH('dbo.MachineCycleTimeSnapshot', 'WorkOrderNo') IS NOT NULL
BEGIN
    ALTER TABLE dbo.MachineCycleTimeSnapshot DROP COLUMN WorkOrderNo;
END;
""", conn, tx);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    private static async Task InsertSnapshotAsync(SqlConnection conn, SqlTransaction tx, Row row, string sourceFileName, CancellationToken ct)
    {
        await using var cmd = new SqlCommand("""
INSERT INTO dbo.MachineCycleTimeSnapshot
    (LineCode, AssemblyNo, Side, MachineName, PanelEndTime, BoardsProcessed,
     MachineMinCycleTimeSec, MachineMediumCycleTimeSec, AdmCurrentCycleTimeSec,
     SourceFileName)
VALUES
    (@line, @assy, @side, @machine, @pet, @boards, @min, @med, @cur, @src);
""", conn, tx);

        cmd.Parameters.AddWithValue("@line", row.LineCode);
        cmd.Parameters.AddWithValue("@assy", row.AssemblyNo);
        cmd.Parameters.AddWithValue("@side", string.IsNullOrWhiteSpace(row.Side) ? (object)DBNull.Value : row.Side);
        cmd.Parameters.AddWithValue("@machine", row.MachineName);
        cmd.Parameters.AddWithValue("@pet", row.PanelEndTime == default ? (object)DBNull.Value : row.PanelEndTime);
        cmd.Parameters.AddWithValue("@boards", row.BoardsProcessed);
        cmd.Parameters.AddWithValue("@min", row.MachineMinSec == 0m ? (object)DBNull.Value : row.MachineMinSec);
        cmd.Parameters.AddWithValue("@med", row.MachineMediumSec == 0m ? (object)DBNull.Value : row.MachineMediumSec);
        cmd.Parameters.AddWithValue("@cur", row.AdmCurrentSec == 0m ? (object)DBNull.Value : row.AdmCurrentSec);
        cmd.Parameters.AddWithValue("@src", sourceFileName);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    private static bool TryParsePanelEndTime(string raw, out DateTime dt)
    {
        dt = default;
        if (string.IsNullOrWhiteSpace(raw)) return false;

        // Typical example from the file: "10/1/2025 12:29:57 AM" (US-style with AM/PM)
        var formats = new[]
        {
            "M/d/yyyy h:mm:ss tt",
            "M/d/yyyy hh:mm:ss tt",
            "MM/dd/yyyy h:mm:ss tt",
            "MM/dd/yyyy hh:mm:ss tt",
            "M/d/yyyy H:mm:ss",
            "MM/dd/yyyy HH:mm:ss"
        };

        if (DateTime.TryParseExact(raw, formats, CultureInfo.GetCultureInfo("en-US"),
                DateTimeStyles.AllowWhiteSpaces, out dt))
            return true;

        if (DateTime.TryParseExact(raw, formats, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces, out dt))
            return true;

        // Fallback: let .NET parse with current culture if it can.
        return DateTime.TryParse(raw, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out dt);
    }
}

