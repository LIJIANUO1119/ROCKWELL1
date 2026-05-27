using System;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI.Services;

public static class FamilyGroupingPersistence
{
    public static async Task SaveGridAsync(SqlConnectionFactory connectionFactory, DataTable table)
    {
        await FamilyGroupingViewData.EnsureTableAsync(connectionFactory);

        await using var conn = connectionFactory.Open();
        await using var tx = await conn.BeginTransactionAsync();
        var sqlTx = (SqlTransaction)tx;

        try
        {
            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted)
                {
                    var id = row["FamilyGroupingDetailId", DataRowVersion.Original];
                    if (id != DBNull.Value && id != null)
                    {
                        await DeleteByIdAsync(conn, sqlTx, Convert.ToInt32(id, CultureInfo.InvariantCulture));
                    }

                    continue;
                }

                var asy = CellString(row, "ASY").Trim();
                if (string.IsNullOrWhiteSpace(asy))
                {
                    var hasId = row["FamilyGroupingDetailId"] != DBNull.Value && row["FamilyGroupingDetailId"] != null;
                    if (!hasId) continue;
                    throw new InvalidOperationException("Each row needs ASY. Delete blank rows or fill ASY before saving.");
                }

                var hasRowId = row["FamilyGroupingDetailId"] != DBNull.Value && row["FamilyGroupingDetailId"] != null;
                if (hasRowId)
                {
                    var detailId = Convert.ToInt32(row["FamilyGroupingDetailId"], CultureInfo.InvariantCulture);
                    await UpdateRowAsync(conn, sqlTx, detailId, row, sourceFileName: null);
                }
                else
                    await InsertFromGridRowAsync(conn, sqlTx, row, sourceFileName: "Manual edit");
            }

            await tx.CommitAsync();
        }
        catch
        {
            await tx.RollbackAsync();
            throw;
        }
    }

    public static async Task DeleteRowAsync(SqlConnectionFactory connectionFactory, int familyGroupingDetailId)
    {
        await FamilyGroupingViewData.EnsureTableAsync(connectionFactory);

        await using var conn = connectionFactory.Open();
        await using var cmd = new SqlCommand(
            "DELETE FROM dbo.FamilyGroupingDetail WHERE FamilyGroupingDetailId = @id;",
            conn);
        cmd.Parameters.AddWithValue("@id", familyGroupingDetailId);
        await cmd.ExecuteNonQueryAsync();
    }

    private static async Task DeleteByIdAsync(SqlConnection conn, SqlTransaction tx, int id)
    {
        await using var cmd = new SqlCommand(
            "DELETE FROM dbo.FamilyGroupingDetail WHERE FamilyGroupingDetailId = @id;",
            conn, tx);
        cmd.Parameters.AddWithValue("@id", id);
        await cmd.ExecuteNonQueryAsync();
    }

    private static async Task InsertFromGridRowAsync(
        SqlConnection conn,
        SqlTransaction tx,
        DataRow row,
        string sourceFileName)
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
        await cmd.ExecuteNonQueryAsync();
    }

    private static async Task UpdateRowAsync(
        SqlConnection conn,
        SqlTransaction tx,
        int id,
        DataRow row,
        string? sourceFileName)
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
    CircuitCount = @circuits
WHERE FamilyGroupingDetailId = @id;
""", conn, tx);

        cmd.Parameters.AddWithValue("@id", id);
        BindRowParameters(cmd, row, sourceFileName);
        await cmd.ExecuteNonQueryAsync();
    }

    private static void BindRowParameters(SqlCommand cmd, DataRow row, string? sourceFileName)
    {
        cmd.Parameters.AddWithValue("@asy", CellString(row, "ASY").Trim());
        cmd.Parameters.AddWithValue("@pcb", NullIfEmpty(CellString(row, "PCB")));
        cmd.Parameters.AddWithValue("@family", NullIfEmpty(CellString(row, "Family")));
        cmd.Parameters.AddWithValue("@familyNum", NullIfEmpty(CellString(row, "Family#")));
        cmd.Parameters.AddWithValue("@tp1", NullIfEmpty(CellString(row, "Top Priority 1")));
        cmd.Parameters.AddWithValue("@tp2", NullIfEmpty(CellString(row, "Top Priority 2")));
        cmd.Parameters.AddWithValue("@tp3", NullIfEmpty(CellString(row, "Top Priority 3")));
        cmd.Parameters.AddWithValue("@topCt", ParseDecimalCell(row, "Top CycleTime"));
        cmd.Parameters.AddWithValue("@bp1", NullIfEmpty(CellString(row, "Bot Priority 1")));
        cmd.Parameters.AddWithValue("@bp2", NullIfEmpty(CellString(row, "Bot Priority 2")));
        cmd.Parameters.AddWithValue("@bp3", NullIfEmpty(CellString(row, "Bot Priority 3")));
        cmd.Parameters.AddWithValue("@botCt", ParseDecimalCell(row, "Bot CycleTime"));
        cmd.Parameters.AddWithValue("@circuits", ParseIntCell(row, "# Circuits"));
        cmd.Parameters.AddWithValue("@src", (object?)sourceFileName ?? DBNull.Value);
    }

    private static string CellString(DataRow row, string columnName)
    {
        if (!row.Table.Columns.Contains(columnName)) return "";
        var v = row[columnName];
        return v == DBNull.Value ? "" : Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
    }

    private static object NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? DBNull.Value : s.Trim();

    private static object ParseDecimalCell(DataRow row, string columnName)
    {
        var s = CellString(row, columnName);
        if (string.IsNullOrWhiteSpace(s)) return DBNull.Value;
        return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)
            ? v
            : DBNull.Value;
    }

    private static object ParseIntCell(DataRow row, string columnName)
    {
        var s = CellString(row, columnName);
        if (string.IsNullOrWhiteSpace(s)) return DBNull.Value;
        return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v)
            ? v
            : DBNull.Value;
    }
}
