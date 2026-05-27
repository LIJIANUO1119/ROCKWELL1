using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI.Services;

/// <summary>Wide machine × nozzle-type quantity matrix (Excel import + grid editing).</summary>
public static class MachineNozzleMatrixService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public const string MachineIdColumn = "Machine ID";
    public const string RowIdColumn = "SmtMachineNozzleMatrixRowId";

    public static async Task<DataTable> LoadWideMatrixAsync(SqlConnectionFactory connectionFactory)
    {
        var initializer = new DbInitializer(connectionFactory);
        await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

        var nozzleCodes = (await LoadNozzleColumnCodesAsync(connectionFactory))
            .Where(c => !IsBannerNozzleColumn(c))
            .ToList();
        var rows = await LoadRowsAsync(connectionFactory);

        foreach (var key in rows.SelectMany(r => r.Quantities.Keys))
        {
            if (IsBannerNozzleColumn(key)) continue;
            if (!nozzleCodes.Contains(key, StringComparer.OrdinalIgnoreCase))
                nozzleCodes.Add(key);
        }

        var table = new DataTable();
        table.Columns.Add(RowIdColumn, typeof(int));
        table.Columns.Add(MachineIdColumn, typeof(string));
        foreach (var code in nozzleCodes)
            table.Columns.Add(code, typeof(int));

        foreach (var row in rows)
        {
            var dr = table.NewRow();
            dr[RowIdColumn] = row.RowId;
            dr[MachineIdColumn] = row.MachineCode;
            foreach (var code in nozzleCodes)
            {
                dr[code] = row.Quantities.TryGetValue(code, out var q) ? q : 0;
            }

            table.Rows.Add(dr);
        }

        return table;
    }

    public static async Task SaveWideMatrixAsync(SqlConnectionFactory connectionFactory, DataTable table)
    {
        var initializer = new DbInitializer(connectionFactory);
        await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

        var nozzleCodes = table.Columns.Cast<DataColumn>()
            .Select(c => c.ColumnName)
            .Where(n => n is not RowIdColumn and not MachineIdColumn)
            .Where(n => !IsBannerNozzleColumn(n))
            .ToList();

        await using var conn = connectionFactory.Open();
        await using var tx = await conn.BeginTransactionAsync();
        var sqlTx = (SqlTransaction)tx;

        try
        {
            await ReplaceColumnCatalogAsync(conn, sqlTx, nozzleCodes);

            var seenMachineCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                var machineCode = CellString(row, MachineIdColumn).Trim();
                if (string.IsNullOrWhiteSpace(machineCode)) continue;

                if (!seenMachineCodes.Add(machineCode))
                {
                    throw new InvalidOperationException(
                        $"Duplicate Machine ID \"{machineCode}\" in the grid. Remove or merge duplicate rows before saving.");
                }

                var quantities = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var code in nozzleCodes)
                {
                    quantities[code] = ParseQuantity(row[code]);
                }

                var json = JsonSerializer.Serialize(quantities);
                var hasRowId = row[RowIdColumn] != DBNull.Value && row[RowIdColumn] != null;

                if (hasRowId)
                {
                    var rowId = Convert.ToInt32(row[RowIdColumn], CultureInfo.InvariantCulture);
                    await using var upd = new SqlCommand("""
UPDATE dbo.SmtMachineNozzleMatrixRow
SET MachineCode = @machineCode,
    QuantitiesJson = @json,
    UpdatedAt = SYSDATETIME()
WHERE SmtMachineNozzleMatrixRowId = @rowId;
""", conn, sqlTx);
                    upd.Parameters.AddWithValue("@rowId", rowId);
                    upd.Parameters.AddWithValue("@machineCode", machineCode);
                    upd.Parameters.AddWithValue("@json", json);
                    await upd.ExecuteNonQueryAsync();
                }
                else
                {
                    await using var ins = new SqlCommand("""
IF EXISTS (SELECT 1 FROM dbo.SmtMachineNozzleMatrixRow WHERE MachineCode = @machineCode)
    UPDATE dbo.SmtMachineNozzleMatrixRow
    SET QuantitiesJson = @json, UpdatedAt = SYSDATETIME()
    WHERE MachineCode = @machineCode;
ELSE
    INSERT INTO dbo.SmtMachineNozzleMatrixRow (MachineCode, QuantitiesJson, UpdatedAt)
    VALUES (@machineCode, @json, SYSDATETIME());
""", conn, sqlTx);
                    ins.Parameters.AddWithValue("@machineCode", machineCode);
                    ins.Parameters.AddWithValue("@json", json);
                    await ins.ExecuteNonQueryAsync();
                }
            }

            await tx.CommitAsync();
        }
        catch
        {
            await tx.RollbackAsync();
            throw;
        }
    }

    public static async Task DeleteRowAsync(SqlConnectionFactory connectionFactory, int rowId)
    {
        var initializer = new DbInitializer(connectionFactory);
        await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

        await using var conn = connectionFactory.Open();
        await using var cmd = new SqlCommand(
            "DELETE FROM dbo.SmtMachineNozzleMatrixRow WHERE SmtMachineNozzleMatrixRowId = @id;",
            conn);
        cmd.Parameters.AddWithValue("@id", rowId);
        await cmd.ExecuteNonQueryAsync();
    }

    public static async Task ReplaceAllFromImportAsync(
        SqlConnectionFactory connectionFactory,
        IReadOnlyList<string> nozzleCodes,
        IReadOnlyList<(string MachineCode, IReadOnlyDictionary<string, int> Quantities)> rows)
    {
        var initializer = new DbInitializer(connectionFactory);
        await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

        await using var conn = connectionFactory.Open();
        await using var tx = await conn.BeginTransactionAsync();
        var sqlTx = (SqlTransaction)tx;

        try
        {
            await using (var delRows = new SqlCommand("DELETE FROM dbo.SmtMachineNozzleMatrixRow;", conn, sqlTx))
                await delRows.ExecuteNonQueryAsync();

            await ReplaceColumnCatalogAsync(conn, sqlTx, nozzleCodes);

            foreach (var row in rows)
            {
                if (string.IsNullOrWhiteSpace(row.MachineCode)) continue;

                var json = JsonSerializer.Serialize(row.Quantities);
                await using var ins = new SqlCommand("""
INSERT INTO dbo.SmtMachineNozzleMatrixRow (MachineCode, QuantitiesJson, UpdatedAt)
VALUES (@machineCode, @json, SYSDATETIME());
""", conn, sqlTx);
                ins.Parameters.AddWithValue("@machineCode", row.MachineCode.Trim());
                ins.Parameters.AddWithValue("@json", json);
                await ins.ExecuteNonQueryAsync();
            }

            await tx.CommitAsync();
        }
        catch
        {
            await tx.RollbackAsync();
            throw;
        }
    }

    private static async Task ReplaceColumnCatalogAsync(SqlConnection conn, SqlTransaction tx, IReadOnlyList<string> nozzleCodes)
    {
        await using (var del = new SqlCommand("DELETE FROM dbo.SmtMachineNozzleColumn;", conn, tx))
            await del.ExecuteNonQueryAsync();

        var order = 0;
        foreach (var code in nozzleCodes.Where(c => !string.IsNullOrWhiteSpace(c)))
        {
            await using var ins = new SqlCommand("""
INSERT INTO dbo.SmtMachineNozzleColumn (NozzleCode, SortOrder) VALUES (@code, @sort);
""", conn, tx);
            ins.Parameters.AddWithValue("@code", code.Trim());
            ins.Parameters.AddWithValue("@sort", order++);
            await ins.ExecuteNonQueryAsync();
        }
    }

    private static async Task<List<string>> LoadNozzleColumnCodesAsync(SqlConnectionFactory connectionFactory)
    {
        await using var conn = connectionFactory.Open();
        await using var cmd = new SqlCommand("""
SELECT NozzleCode FROM dbo.SmtMachineNozzleColumn ORDER BY SortOrder, NozzleCode;
""", conn);

        var list = new List<string>();
        await using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
            list.Add(reader.GetString(0));
        return list;
    }

    private sealed record MatrixRow(int RowId, string MachineCode, Dictionary<string, int> Quantities);

    private static async Task<List<MatrixRow>> LoadRowsAsync(SqlConnectionFactory connectionFactory)
    {
        await using var conn = connectionFactory.Open();
        await using var cmd = new SqlCommand("""
SELECT SmtMachineNozzleMatrixRowId, MachineCode, QuantitiesJson
FROM dbo.SmtMachineNozzleMatrixRow
ORDER BY MachineCode;
""", conn);

        var list = new List<MatrixRow>();
        await using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var id = reader.GetInt32(0);
            var machine = reader.GetString(1);
            var json = reader.IsDBNull(2) ? "{}" : reader.GetString(2);
            var dict = JsonSerializer.Deserialize<Dictionary<string, int>>(json, JsonOptions)
                       ?? new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            list.Add(new MatrixRow(id, machine, dict));
        }

        return list;
    }

    internal static bool IsBannerNozzleColumn(string? columnName)
    {
        if (string.IsNullOrWhiteSpace(columnName)) return false;
        var h = columnName.Trim();
        return h.Contains("nozzle", StringComparison.OrdinalIgnoreCase)
               && h.Contains("type", StringComparison.OrdinalIgnoreCase)
               && (h.Contains("quantity", StringComparison.OrdinalIgnoreCase) || h.Contains('&'));
    }

    private static string CellString(DataRow row, string columnName)
    {
        if (!row.Table.Columns.Contains(columnName)) return "";
        var v = row[columnName];
        return v == DBNull.Value ? "" : Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
    }

    private static int ParseQuantity(object value)
    {
        if (value == DBNull.Value || value is null) return 0;
        if (value is int i) return i;
        if (value is long l) return (int)l;
        if (value is double d) return (int)Math.Round(d);
        var s = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        if (int.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)) return n;
        if (double.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var d2)) return (int)Math.Round(d2);
        return 0;
    }
}
