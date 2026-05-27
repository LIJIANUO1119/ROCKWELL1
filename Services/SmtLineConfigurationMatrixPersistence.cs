using System;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI.Services;

/// <summary>Persist edits from the line configuration matrix grid to dbo.Line / dbo.Machine.</summary>
public static class SmtLineConfigurationMatrixPersistence
{
    public static async Task SaveMatrixAsync(SqlConnectionFactory connectionFactory, DataTable table)
    {
        var initializer = new DbInitializer(connectionFactory);
        await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

        await using var conn = connectionFactory.Open();
        await using var tx = await conn.BeginTransactionAsync();
        var sqlTx = (SqlTransaction)tx;

        try
        {
            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                var lineCode = CellString(row, "Process Line");
                var machineCode = CellString(row, "Machine ID");
                var hasMachineId = row["MachineId"] != DBNull.Value && row["MachineId"] != null;

                if (!hasMachineId && string.IsNullOrWhiteSpace(machineCode) && string.IsNullOrWhiteSpace(lineCode))
                    continue;

                if (string.IsNullOrWhiteSpace(lineCode) || string.IsNullOrWhiteSpace(machineCode))
                {
                    throw new InvalidOperationException(
                        "Each machine row needs Process Line and Machine ID. Delete unused blank rows, or fill both columns before saving.");
                }

                var lineId = await GetOrCreateLineIdAsync(conn, sqlTx, lineCode.Trim());

                if (!hasMachineId)
                    await InsertMachineAsync(conn, sqlTx, lineId, row);
                else
                {
                    var machineId = Convert.ToInt32(row["MachineId"], CultureInfo.InvariantCulture);
                    await UpdateMachineAsync(conn, sqlTx, machineId, lineId, row);
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

    public static async Task DeleteMachineRowAsync(SqlConnectionFactory connectionFactory, int machineId)
    {
        var initializer = new DbInitializer(connectionFactory);
        await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

        await using var conn = connectionFactory.Open();
        await using var tx = await conn.BeginTransactionAsync();
        var sqlTx = (SqlTransaction)tx;

        try
        {
            await DeleteMachineCascadeAsync(conn, sqlTx, machineId);
            await tx.CommitAsync();
        }
        catch
        {
            await tx.RollbackAsync();
            throw;
        }
    }

    private static string CellString(DataRow row, string columnName)
    {
        if (!row.Table.Columns.Contains(columnName)) return "";
        var v = row[columnName];
        return v == DBNull.Value ? "" : Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
    }

    private static async Task<int> GetOrCreateLineIdAsync(SqlConnection conn, SqlTransaction tx, string lineCode)
    {
        await using (var sel = new SqlCommand("SELECT LineId FROM dbo.Line WHERE LineCode = @c;", conn, tx))
        {
            sel.Parameters.AddWithValue("@c", lineCode);
            var o = await sel.ExecuteScalarAsync();
            if (o is not null && o != DBNull.Value)
                return Convert.ToInt32(o, CultureInfo.InvariantCulture);
        }

        await using var ins = new SqlCommand("""
INSERT INTO dbo.Line (LineCode, LineName, IsActive)
OUTPUT INSERTED.LineId
VALUES (@c, @c, 1);
""", conn, tx);
        ins.Parameters.AddWithValue("@c", lineCode);
        var idObj = await ins.ExecuteScalarAsync();
        return Convert.ToInt32(idObj, CultureInfo.InvariantCulture);
    }

    private static async Task InsertMachineAsync(SqlConnection conn, SqlTransaction tx, int lineId, DataRow row)
    {
        var machineCode = CellString(row, "Machine ID").Trim();
        await using var cmd = new SqlCommand("""
INSERT INTO dbo.Machine
    (LineId, MachineCode, MachineType, EquipType, IsActive, PositionInLine)
VALUES
    (@lineId, @machineCode, @machineType, @equipType, 1, NULL);
""", conn, tx);

        cmd.Parameters.AddWithValue("@lineId", lineId);
        cmd.Parameters.AddWithValue("@machineCode", machineCode);
        cmd.Parameters.AddWithValue("@machineType", NullIfEmpty(CellString(row, "Machine Model")));
        cmd.Parameters.AddWithValue("@equipType", NullIfEmpty(CellString(row, "Equipment Type")));
        await cmd.ExecuteNonQueryAsync();
    }

    private static async Task UpdateMachineAsync(SqlConnection conn, SqlTransaction tx, int machineId, int lineId, DataRow row)
    {
        var machineCode = CellString(row, "Machine ID").Trim();
        await using var cmd = new SqlCommand("""
UPDATE dbo.Machine
SET LineId = @lineId,
    MachineCode = @machineCode,
    MachineType = @machineType,
    EquipType = @equipType,
    UpdatedAt = SYSDATETIME()
WHERE MachineId = @machineId;
""", conn, tx);

        cmd.Parameters.AddWithValue("@machineId", machineId);
        cmd.Parameters.AddWithValue("@lineId", lineId);
        cmd.Parameters.AddWithValue("@machineCode", machineCode);
        cmd.Parameters.AddWithValue("@machineType", NullIfEmpty(CellString(row, "Machine Model")));
        cmd.Parameters.AddWithValue("@equipType", NullIfEmpty(CellString(row, "Equipment Type")));
        await cmd.ExecuteNonQueryAsync();
    }

    private static object NullIfEmpty(string s)
        => string.IsNullOrWhiteSpace(s) ? DBNull.Value : s.Trim();

    private static async Task DeleteMachineCascadeAsync(SqlConnection conn, SqlTransaction tx, int machineId)
    {
        await using (var c1 = new SqlCommand("DELETE FROM dbo.MachineNozzleConfig WHERE MachineId=@id;", conn, tx))
        {
            c1.Parameters.AddWithValue("@id", machineId);
            await c1.ExecuteNonQueryAsync();
        }

        await using (var c2 = new SqlCommand("DELETE FROM dbo.MachineSoftware WHERE MachineId=@id;", conn, tx))
        {
            c2.Parameters.AddWithValue("@id", machineId);
            await c2.ExecuteNonQueryAsync();
        }

        await using (var c3 = new SqlCommand("DELETE FROM dbo.SmtMachineConstraintOptionValue WHERE MachineId=@id;", conn, tx))
        {
            c3.Parameters.AddWithValue("@id", machineId);
            await c3.ExecuteNonQueryAsync();
        }

        await using (var c4 = new SqlCommand("DELETE FROM dbo.MachineConstraint WHERE MachineId=@id;", conn, tx))
        {
            c4.Parameters.AddWithValue("@id", machineId);
            await c4.ExecuteNonQueryAsync();
        }

        await using (var c5 = new SqlCommand("DELETE FROM dbo.MachineCycleTime WHERE MachineId=@id;", conn, tx))
        {
            c5.Parameters.AddWithValue("@id", machineId);
            await c5.ExecuteNonQueryAsync();
        }

        await using (var c6 = new SqlCommand("UPDATE dbo.ProductionHistory SET MachineId = NULL WHERE MachineId=@id;", conn, tx))
        {
            c6.Parameters.AddWithValue("@id", machineId);
            await c6.ExecuteNonQueryAsync();
        }

        await using (var c7 = new SqlCommand("DELETE FROM dbo.Machine WHERE MachineId=@id;", conn, tx))
        {
            c7.Parameters.AddWithValue("@id", machineId);
            await c7.ExecuteNonQueryAsync();
        }
    }
}
