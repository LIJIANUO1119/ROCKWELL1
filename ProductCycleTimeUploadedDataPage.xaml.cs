using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;
using SmtLineAllocationUI.Services;

namespace SmtLineAllocationUI;

public partial class ProductCycleTimeUploadedDataPage : Page
{
    private readonly SqlConnectionFactory _connectionFactory = SqlConnectionFactory.FromAppSettings();

    public ProductCycleTimeUploadedDataPage()
    {
        InitializeComponent();
        Loaded += (_, _) => _ = LoadDataAsync();
    }

    private async Task LoadDataAsync()
    {
        try
        {
            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var counts = await QueryCountsAsync();
            var table = await QueryAsync(ProductCycleTimeViewData.JoinedUploadedDataSql);

            GridData.ItemsSource = table.DefaultView;
            ApplyGridColumnLayout();

            TxtSubtitle.Text =
                $"Joined on Module# = AssemblyNo · {table.Rows.Count} row(s) · " +
                $"build-to-module: {counts.BuildRows} · machine cycletime: {counts.MachineRows}";

            if (table.Rows.Count == 0)
            {
                TxtSubtitle.Text +=
                    " · Upload both datasets on the Product Cycletime page, then open this view again.";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Load failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task<(int BuildRows, int MachineRows)> QueryCountsAsync()
    {
        var table = await QueryAsync("""
SELECT
    (SELECT COUNT(1) FROM dbo.BuildToModuleMapping) AS BuildRows,
    (SELECT COUNT(1) FROM dbo.MachineCycleTimeSnapshot) AS MachineRows;
""");
        if (table.Rows.Count == 0) return (0, 0);
        return (
            Convert.ToInt32(table.Rows[0]["BuildRows"]),
            Convert.ToInt32(table.Rows[0]["MachineRows"])
        );
    }

    private void ApplyGridColumnLayout()
    {
        if (GridData.Columns.Count == 0) return;

        foreach (var col in GridData.Columns)
        {
            var header = col.Header?.ToString() ?? "";
            if (header is "Build" or "Module#" or "PCB" or "LINE" or "MACHINENAME")
            {
                col.Width = new DataGridLength(100);
                col.MinWidth = 72;
            }
            else if (header is "Build Key" or "Side" or "AssemblyNo")
            {
                col.Width = new DataGridLength(88);
                col.MinWidth = 64;
            }
            else if (header.Contains("CYCLETIME", StringComparison.OrdinalIgnoreCase)
                     || header == "Bottleneck"
                     || header == "BOARDSPROCESSED")
            {
                col.Width = new DataGridLength(88);
                col.MinWidth = 72;
            }
            else if (header.Contains("At", StringComparison.OrdinalIgnoreCase)
                     || header == "PANELENDTIME"
                     || header.Contains("Source", StringComparison.OrdinalIgnoreCase))
            {
                col.Width = new DataGridLength(140);
                col.MinWidth = 100;
            }
            else
            {
                col.Width = new DataGridLength(90);
            }

            col.CanUserResize = true;
        }

        GridData.FrozenColumnCount = Math.Min(3, GridData.Columns.Count);
    }

    private async Task<DataTable> QueryAsync(string sql)
    {
        await using var conn = _connectionFactory.Open();
        await using var cmd = new SqlCommand(sql, conn);
        await using var reader = await cmd.ExecuteReaderAsync();
        var table = new DataTable();
        table.Load(reader);
        return table;
    }

    private void BtnBack_OnClick(object sender, RoutedEventArgs e)
    {
        if (NavigationService?.CanGoBack == true)
            NavigationService.GoBack();
        else
            NavigationService?.Navigate(new ProductCycleTimePage());
    }
}
