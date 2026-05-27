using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI;

public partial class ProductCycleTimeQueryResultsPage : Page
{
    private readonly SqlConnectionFactory _connectionFactory = SqlConnectionFactory.FromAppSettings();
    private readonly string _filterField;
    private readonly string _filterValue;

    public ProductCycleTimeQueryResultsPage(string filterField, string filterValue)
    {
        _filterField = filterField ?? throw new ArgumentNullException(nameof(filterField));
        _filterValue = filterValue ?? throw new ArgumentNullException(nameof(filterValue));

        InitializeComponent();
        Loaded += async (_, _) => await LoadAsync();
    }

    private async Task LoadAsync()
    {
        try
        {
            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            TxtSubtitle.Text = $"{_filterField} = {_filterValue}";

            var sql = BuildSql(_filterField);
            var table = await QueryAsync(sql, new SqlParameter("@v", _filterValue));
            GridData.Columns.Clear();
            GridData.ItemsSource = table.DefaultView;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Query failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string BuildSql(string filterField)
    {
        // Join key: BuildToModuleMapping.ModuleNumber == MachineCycleTimeSnapshot.AssemblyNo (case/space insensitive).
        var where = filterField switch
        {
            "Build" => "WHERE b.BuildNumber = @v",
            "Module #" => "WHERE b.ModuleNumber = @v",
            "PCB" => "WHERE b.PCB = @v",
            _ => throw new InvalidOperationException("Unsupported filter field.")
        };

        return $"""
WITH joined AS (
    SELECT
        b.BuildNumber AS [Build],
        b.ModuleNumber AS [Module#],
        b.PCB AS [PCB],
        s.LineCode AS [LINE],
        s.Side AS [Side],
        s.MachineName AS [MACHINENAME],
        s.MachineMinCycleTimeSec AS [MACHINE_MIN_CYCLETIME],
        s.MachineMediumCycleTimeSec AS [MACHINE_MEDIUM_CYCLETIME],
        s.AdmCurrentCycleTimeSec AS [ADM_CURRENT_CYCLETIME]
    FROM dbo.BuildToModuleMapping b
    INNER JOIN dbo.MachineCycleTimeSnapshot s
        ON UPPER(LTRIM(RTRIM(b.ModuleNumber))) = UPPER(LTRIM(RTRIM(s.AssemblyNo)))
    {where}
),
withBn AS (
    SELECT
        [Build],
        [Module#],
        [PCB],
        [LINE],
        [Side],
        [MACHINENAME],
        [MACHINE_MIN_CYCLETIME],
        [MACHINE_MEDIUM_CYCLETIME],
        [ADM_CURRENT_CYCLETIME],
        MAX([MACHINE_MEDIUM_CYCLETIME]) OVER (PARTITION BY [Build], [LINE], [Module#], [Side]) AS [Bottleneck]
    FROM joined
)
SELECT
    [Build],
    [Module#],
    [PCB],
    [LINE],
    [Side],
    [MACHINENAME],
    [MACHINE_MIN_CYCLETIME],
    [MACHINE_MEDIUM_CYCLETIME],
    [ADM_CURRENT_CYCLETIME],
    [Bottleneck]
FROM withBn
ORDER BY [Build], [LINE], [Module#], [Side], [MACHINENAME];
""";
    }

    private async Task<DataTable> QueryAsync(string sql, params SqlParameter[] parameters)
    {
        await using var conn = _connectionFactory.Open();
        await using var cmd = new SqlCommand(sql, conn);
        if (parameters is { Length: > 0 })
            cmd.Parameters.AddRange(parameters);

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

