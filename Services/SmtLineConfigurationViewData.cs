using System.Data;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI.Services;

/// <summary>Queries for full-page SMT line / machine nozzle views.</summary>
public static class SmtLineConfigurationViewData
{
    public static async Task<DataTable> LoadSmtLineConfigurationMatrixAsync(SqlConnectionFactory connectionFactory)
    {
        var initializer = new DbInitializer(connectionFactory);
        await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

        const string sql = """
SELECT
    l.LineId,
    m.MachineId,
    l.LineCode AS [Process Line],
    m.MachineCode AS [Machine ID],
    m.MachineType AS [Machine Model],
    m.EquipType AS [Equipment Type]
FROM dbo.Line l
LEFT JOIN dbo.Machine m ON m.LineId = l.LineId
ORDER BY l.LineCode, m.MachineCode;
""";

        await using var conn = connectionFactory.Open();
        await using var cmd = new SqlCommand(sql, conn);
        await using var reader = await cmd.ExecuteReaderAsync();
        var table = new DataTable();
        table.Load(reader);
        return table;
    }

    /// <summary>DataGrid binding breaks if column alias contains '/'; fix display header for serial column.</summary>
    public static void ApplySerialColumnHeader(System.Windows.Controls.DataGrid grid)
    {
        grid.UpdateLayout();
        foreach (var col in grid.Columns)
        {
            var sortPath = col.SortMemberPath;
            var headerText = col.Header?.ToString();
            if (sortPath == "SerialNumberOrComputerName_Display" || headerText == "SerialNumberOrComputerName_Display")
                col.Header = "Serial Number/Computer Name";
        }
    }
}
