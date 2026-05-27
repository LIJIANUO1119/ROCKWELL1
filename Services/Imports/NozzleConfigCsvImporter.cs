using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI.Services.Imports;

/// <summary>
/// Simple CSV-based nozzle configuration importer.
/// Expected columns (no header required but recommended):
///   MachineCode,HeadLabel,SlotIndex,NozzleCode
/// Import strategy:
///   - Find MachineId by MachineCode.
///   - Upsert NozzleType by NozzleCode (empty -> __EMPTY_NOZZLE__).
///   - Upsert MachineNozzleConfig by (MachineId, SlotIndex).
/// </summary>
public sealed class NozzleConfigCsvImporter
{
    private readonly SqlConnectionFactory _connectionFactory;

    public NozzleConfigCsvImporter(SqlConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
    }

    public async Task<ImportResult> ImportAsync(string csvPath, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(csvPath))
            throw new ArgumentException("Path cannot be empty.", nameof(csvPath));
        if (!File.Exists(csvPath))
            throw new FileNotFoundException("CSV file not found.", csvPath);

        var errors = new List<string>();
        var read = 0;
        var ok = 0;
        var failed = 0;

        await using var conn = _connectionFactory.Open();
        await using var tx = (SqlTransaction)await conn.BeginTransactionAsync(ct);

        try
        {
            await foreach (var row in ReadRowsAsync(csvPath, ct))
            {
                ct.ThrowIfCancellationRequested();
                read++;

                var lineInfo = $"line {row.LineNumber}";

                try
                {
                    var machineId = await FindMachineIdAsync(conn, tx, row.MachineCode, ct);
                    if (machineId is null)
                    {
                        failed++;
                        errors.Add($"{lineInfo}: MachineCode '{row.MachineCode}' not found.");
                        continue;
                    }

                    var nozzleTypeId = await EnsureNozzleTypeAsync(conn, tx, row.NozzleCode, ct);
                    await UpsertMachineNozzleAsync(
                        conn,
                        tx,
                        machineId.Value,
                        nozzleTypeId,
                        row.HeadLabel,
                        row.SlotIndex,
                        quantity: 1,
                        ct);

                    ok++;
                }
                catch (Exception ex)
                {
                    failed++;
                    errors.Add($"{lineInfo}: {ex.Message}");
                }
            }

            await tx.CommitAsync(ct);
        }
        catch
        {
            await tx.RollbackAsync(ct);
            throw;
        }

        return new ImportResult(Path.GetFileName(csvPath), read, ok, failed, errors);
    }

    private readonly struct CsvRow
    {
        public CsvRow(int lineNumber, string machineCode, string headLabel, int slotIndex, string nozzleCode)
        {
            LineNumber = lineNumber;
            MachineCode = machineCode;
            HeadLabel = headLabel;
            SlotIndex = slotIndex;
            NozzleCode = nozzleCode;
        }

        public int LineNumber { get; }
        public string MachineCode { get; }
        public string HeadLabel { get; }
        public int SlotIndex { get; }
        public string NozzleCode { get; }
    }

    private static async IAsyncEnumerable<CsvRow> ReadRowsAsync(string path, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct)
    {
        using var reader = new StreamReader(path);
        var lineNumber = 0;
        var first = true;

        while (!reader.EndOfStream)
        {
            ct.ThrowIfCancellationRequested();
            var line = await reader.ReadLineAsync();
            if (line is null) break;
            lineNumber++;

            if (string.IsNullOrWhiteSpace(line))
                continue;

            // Optional header: if the first line contains "MachineCode" treat it as header.
            if (first)
            {
                first = false;
                if (line.Contains("MachineCode", StringComparison.OrdinalIgnoreCase))
                    continue;
            }

            var parts = line.Split(',');
            if (parts.Length < 4)
                throw new FormatException($"Line {lineNumber}: expected at least 4 comma-separated fields.");

            var machineCode = parts[0].Trim();
            if (string.IsNullOrWhiteSpace(machineCode))
                throw new FormatException($"Line {lineNumber}: MachineCode is required.");

            var headLabel = parts[1].Trim();
            if (string.IsNullOrWhiteSpace(headLabel))
                throw new FormatException($"Line {lineNumber}: HeadLabel is required.");

            if (!int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var slotIndex))
                throw new FormatException($"Line {lineNumber}: invalid SlotIndex '{parts[2]}'.");

            var nozzleCode = parts[3].Trim();

            yield return new CsvRow(lineNumber, machineCode, headLabel, slotIndex, nozzleCode);
        }
    }

    private static async Task<int?> FindMachineIdAsync(SqlConnection conn, SqlTransaction tx, string machineCode, CancellationToken ct)
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

    private static async Task<int> EnsureNozzleTypeAsync(SqlConnection conn, SqlTransaction tx, string nozzleCode, CancellationToken ct)
    {
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
            cmd.Parameters.AddWithValue("@desc", "Imported from nozzle CSV");
            var id = await cmd.ExecuteScalarAsync(ct);
            return (int)id!;
        }
    }

    private static bool IsEmptyNozzlePlaceholder(string value)
    {
        var t = value.Trim();
        if (string.IsNullOrWhiteSpace(t)) return true;
        if (t.Equals("__EMPTY_NOZZLE__", StringComparison.OrdinalIgnoreCase)) return true;

        var u = t.ToUpperInvariant();
        return u.Contains("EMPTY", StringComparison.OrdinalIgnoreCase) && u.Contains("NOZZ", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeNozzleCode(string nozzleCode)
    {
        if (string.IsNullOrWhiteSpace(nozzleCode)) return string.Empty;
        var t = nozzleCode.Trim();
        if (IsEmptyNozzlePlaceholder(t)) return string.Empty;
        return t;
    }

    private static async Task UpsertMachineNozzleAsync(
        SqlConnection conn,
        SqlTransaction tx,
        int machineId,
        int nozzleTypeId,
        string headLabel,
        int slotIndex,
        int quantity,
        CancellationToken ct)
    {
        await using (var cmd = new SqlCommand("""
SELECT MachineNozzleConfigId FROM dbo.MachineNozzleConfig
WHERE MachineId=@mid AND SlotIndex=@slot
  AND ISNULL(HeadLabel,'') = ISNULL(@headLabel,'');
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@mid", machineId);
            cmd.Parameters.AddWithValue("@slot", slotIndex);
            cmd.Parameters.AddWithValue("@headLabel", headLabel);
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
                up.Parameters.AddWithValue("@headLabel", headLabel);
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
            cmd.Parameters.AddWithValue("@headLabel", headLabel);
            cmd.Parameters.AddWithValue("@slot", slotIndex);
            cmd.Parameters.AddWithValue("@qty", quantity);
            await cmd.ExecuteNonQueryAsync(ct);
        }
    }
}

