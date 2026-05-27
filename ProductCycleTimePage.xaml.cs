using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using SmtLineAllocationUI.DataAccess;
using SmtLineAllocationUI.Services;
using SmtLineAllocationUI.Services.Imports;

namespace SmtLineAllocationUI;

public partial class ProductCycleTimePage : Page
{
    private readonly CycleTimeStatisticsService _stats = new();
    private readonly SqlConnectionFactory _connectionFactory = SqlConnectionFactory.FromAppSettings();

    public ProductCycleTimePage()
    {
        InitializeComponent();
    }

    private async void BtnQueryAllCycleTime_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            BtnQueryAllCycleTime.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            if (!await EnsureCycleTimePrerequisitesAsync())
                return;

            var option = ShowChoice("Query by", new[]
            {
                "Build",
                "Module #",
                "PCB"
            });
            if (option is null) return;

            var (ok, value) = ShowInput($"Enter {option}", defaultValue: "");
            if (!ok) return;
            value = (value ?? "").Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                MessageBox.Show("Input cannot be empty.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            NavigationService?.Navigate(new ProductCycleTimeQueryResultsPage(option, value));
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Query failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnQueryAllCycleTime.IsEnabled = true;
        }
    }

    private async void BtnShowNotMeetingTarget_OnClick(object sender, RoutedEventArgs e)
    {
        await ExportNotMeetingTargetAsync();
    }

    private async void BtnShowExceedTarget_OnClick(object sender, RoutedEventArgs e)
    {
        await ExportExceedTargetAsync();
    }

    private async Task ExportNotMeetingTargetAsync()
    {
        var dlg = new SaveFileDialog
        {
            Title = "Export not meeting target to CSV",
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = "REPORT_NOT_MEETING_TARGET.csv"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            BtnShowNotMeetingTarget.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            if (!await EnsureCycleTimePrerequisitesAsync())
                return;

            var table = await QueryAsync(NotMeetingTargetExportSql());
            WriteCsv(table, dlg.FileName);

            MessageBox.Show($"Export completed.\nRows: {table.Rows.Count}", "Export", MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnShowNotMeetingTarget.IsEnabled = true;
        }
    }

    private async Task ExportExceedTargetAsync()
    {
        var dlg = new SaveFileDialog
        {
            Title = "Export exceed target to CSV",
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = "REPORT_EXCEED_TARGET.csv"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            BtnShowExceedTarget.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            if (!await EnsureCycleTimePrerequisitesAsync())
                return;

            var table = await QueryAsync(ExceedTargetExportSql());
            WriteCsv(table, dlg.FileName);

            MessageBox.Show($"Export completed.\nRows: {table.Rows.Count}", "Export", MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnShowExceedTarget.IsEnabled = true;
        }
    }

    private async Task RunReportToGridAsync(string filter)
    {
        try
        {
            BtnShowNotMeetingTarget.IsEnabled = false;
            BtnShowExceedTarget.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            if (!await EnsureCycleTimePrerequisitesAsync())
                return;

            // Grid removed per UI requirement; query is still executed as a quick health-check.
            _ = await QueryAsync(CycleTimeQuerySql(allOnly: false, filter));
            MessageBox.Show("OK. Report query executed.", "Product cycletime", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Query failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnShowNotMeetingTarget.IsEnabled = true;
            BtnShowExceedTarget.IsEnabled = true;
        }
    }

    private void BtnViewUploadedData_OnClick(object sender, RoutedEventArgs e)
    {
        NavigationService?.Navigate(new ProductCycleTimeUploadedDataPage());
    }

    private async void BtnExportAllProductCycletime_OnClick(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog
        {
            Title = "Export cycle time report to CSV",
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = "ALL_PRODUCT_CYCLTIME.csv"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            BtnExportAllProductCycletime.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            if (!await EnsureCycleTimePrerequisitesAsync())
                return;

            var table = await QueryAsync(JoinedDetailExportSql());
            WriteCsv(table, dlg.FileName);
            MessageBox.Show("Export completed.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnExportAllProductCycletime.IsEnabled = true;
        }
    }

    private static string JoinedDetailExportSql()
    {
        // Same join as the on-screen "View uploaded data" grid (export omits traceability columns).
        return """
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
        s.AdmCurrentCycleTimeSec AS [ADM_CURRENT_CYCLETIME],
        s.PanelEndTime AS [PANELENDTIME],
        s.CreatedAt AS [SnapshotCreatedAt]
    FROM dbo.BuildToModuleMapping b
    INNER JOIN dbo.MachineCycleTimeSnapshot s
        ON UPPER(LTRIM(RTRIM(b.ModuleNumber))) = UPPER(LTRIM(RTRIM(s.AssemblyNo)))
),
withBn AS (
    SELECT
        *,
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
    [Bottleneck],
    [PANELENDTIME],
    [SnapshotCreatedAt]
FROM withBn
ORDER BY [Build], [LINE], [Module#], [Side], [MACHINENAME];
""";
    }

    private static string NotMeetingTargetExportSql()
    {
        // Requirement:
        // - Export rows where ADM_CURRENT_CYCLETIME < 90% of MACHINE_MEDIUM_CYCLETIME.
        // - If either column is missing (NULL) then the row is NOT included.
        return """
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
        s.AdmCurrentCycleTimeSec AS [ADM_CURRENT_CYCLETIME],
        s.PanelEndTime AS [PANELENDTIME],
        s.CreatedAt AS [SnapshotCreatedAt]
    FROM dbo.BuildToModuleMapping b
    INNER JOIN dbo.MachineCycleTimeSnapshot s
        ON UPPER(LTRIM(RTRIM(b.ModuleNumber))) = UPPER(LTRIM(RTRIM(s.AssemblyNo)))
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
    [PANELENDTIME],
    [SnapshotCreatedAt]
FROM joined
WHERE [ADM_CURRENT_CYCLETIME] IS NOT NULL
  AND [MACHINE_MEDIUM_CYCLETIME] IS NOT NULL
  AND [MACHINE_MEDIUM_CYCLETIME] <> 0
  AND [ADM_CURRENT_CYCLETIME] < (0.9 * [MACHINE_MEDIUM_CYCLETIME])
ORDER BY [Build], [LINE], [Module#], [Side], [MACHINENAME];
""";
    }

    private static string ExceedTargetExportSql()
    {
        // Requirement:
        // - Export rows where ADM_CURRENT_CYCLETIME > MACHINE_MEDIUM_CYCLETIME.
        // - If either column is missing (NULL) then the row is NOT included.
        return """
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
        s.AdmCurrentCycleTimeSec AS [ADM_CURRENT_CYCLETIME],
        s.PanelEndTime AS [PANELENDTIME],
        s.CreatedAt AS [SnapshotCreatedAt]
    FROM dbo.BuildToModuleMapping b
    INNER JOIN dbo.MachineCycleTimeSnapshot s
        ON UPPER(LTRIM(RTRIM(b.ModuleNumber))) = UPPER(LTRIM(RTRIM(s.AssemblyNo)))
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
    [PANELENDTIME],
    [SnapshotCreatedAt]
FROM joined
WHERE [ADM_CURRENT_CYCLETIME] IS NOT NULL
  AND [MACHINE_MEDIUM_CYCLETIME] IS NOT NULL
  AND [ADM_CURRENT_CYCLETIME] > [MACHINE_MEDIUM_CYCLETIME]
ORDER BY [Build], [LINE], [Module#], [Side], [MACHINENAME];
""";
    }

    private async void BtnUploadMachineCycletime_OnClick(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select SGP_GEM_MACHINE_CYCLETIME CSV",
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            InitialDirectory = GetDefaultDataDirectory()
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            BtnUploadMachineCycletime.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var importer = new MachineCycletimeCsvImporter(_connectionFactory);
            var result = await importer.ImportAsync(dlg.FileName);

            var detail = result.Errors.Count > 0
                ? "\n\n" + string.Join("\n", result.Errors)
                : "";
            MessageBox.Show(
                $"Imported: {result.FileName}\nRead: {result.ReadCount}\nInserted (after clean/dedupe): {result.SuccessCount}\nParse failures: {result.FailureCount}{detail}",
                "Import result",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Import failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnUploadMachineCycletime.IsEnabled = true;
        }
    }

    private async void BtnUploadBuildToModule_OnClick(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select SGP_ASSEMBLY_BUILD_TO_MODULE.csv (Build #, Build Key, Module #, PCB)",
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            InitialDirectory = GetDefaultDataDirectory()
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            BtnUploadBuildToModule.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var importer = new BuildToModuleCsvImporter(_connectionFactory);
            var result = await importer.ImportAsync(dlg.FileName);

            var detail = result.Errors.Count > 0
                ? "\n\n" + string.Join("\n", result.Errors)
                : "";
            MessageBox.Show(
                $"Imported: {result.FileName}\nRead: {result.ReadCount}\nInserted (after clean/dedupe): {result.SuccessCount}\nParse failures: {result.FailureCount}{detail}",
                "Import result",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Import failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnUploadBuildToModule.IsEnabled = true;
        }
    }

    private static string GetDefaultDataDirectory()
    {
        try
        {
            var baseDir = AppContext.BaseDirectory;
            var dir = new DirectoryInfo(baseDir);
            while (dir is not null)
            {
                var candidate = Path.Combine(dir.FullName, "data");
                if (Directory.Exists(candidate)) return candidate;
                dir = dir.Parent;
            }
        }
        catch
        {
            // ignore
        }

        return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
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

    private async Task<bool> EnsureCycleTimePrerequisitesAsync()
    {
        // Requirement: system must detect BOTH tables before working.
        // Link key: MachineCycleTimeSnapshot.AssemblyNo == BuildToModuleMapping.ModuleNumber.
        var table = await QueryAsync("""
SELECT
    CASE WHEN OBJECT_ID(N'dbo.MachineCycleTimeSnapshot', N'U') IS NULL THEN 0 ELSE (SELECT COUNT(1) FROM dbo.MachineCycleTimeSnapshot) END AS MachineCycleTimeRows,
    CASE WHEN OBJECT_ID(N'dbo.BuildToModuleMapping', N'U') IS NULL THEN 0 ELSE (SELECT COUNT(1) FROM dbo.BuildToModuleMapping) END AS BuildToModuleRows;
""");

        var machineRows = table.Rows.Count == 0 ? 0 : Convert.ToInt32(table.Rows[0]["MachineCycleTimeRows"]);
        var buildRows = table.Rows.Count == 0 ? 0 : Convert.ToInt32(table.Rows[0]["BuildToModuleRows"]);

        if (machineRows > 0 && buildRows > 0)
            return true;

        var missing = new System.Collections.Generic.List<string>();
        if (machineRows == 0) missing.Add("machine cycletime (Upload machine cycletime)");
        if (buildRows == 0) missing.Add("build-to-module (Upload build-to-module)");

        MessageBox.Show(
            "Missing required data:\n- " + string.Join("\n- ", missing),
            "Product cycletime prerequisites",
            MessageBoxButton.OK,
            MessageBoxImage.Warning
        );
        return false;
    }

    private static string? ShowChoice(string title, System.Collections.Generic.IReadOnlyList<string> options)
    {
        var win = new Window
        {
            Title = title,
            SizeToContent = SizeToContent.WidthAndHeight,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
            Owner = Application.Current?.MainWindow
        };

        string? selected = null;
        var root = new StackPanel { Margin = new Thickness(16) };
        foreach (var opt in options)
        {
            var btn = new Button
            {
                Content = opt,
                MinWidth = 260,
                Height = 34,
                Margin = new Thickness(0, 0, 0, 8)
            };
            btn.Click += (_, _) =>
            {
                selected = opt;
                win.DialogResult = true;
            };
            root.Children.Add(btn);
        }

        var cancel = new Button { Content = "Cancel", MinWidth = 260, Height = 34 };
        cancel.Click += (_, _) => win.DialogResult = false;
        root.Children.Add(cancel);

        win.Content = root;
        _ = win.ShowDialog();
        return selected;
    }

    private static (bool ok, string value) ShowInput(string title, string defaultValue)
    {
        var win = new Window
        {
            Title = title,
            SizeToContent = SizeToContent.WidthAndHeight,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
            Owner = Application.Current?.MainWindow
        };

        var root = new StackPanel { Margin = new Thickness(16) };
        var tb = new TextBox { MinWidth = 320, Text = defaultValue ?? "", Margin = new Thickness(0, 0, 0, 12) };
        root.Children.Add(tb);

        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        var ok = new Button { Content = "OK", Width = 90, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
        var cancel = new Button { Content = "Cancel", Width = 90, IsCancel = true };
        buttons.Children.Add(ok);
        buttons.Children.Add(cancel);
        root.Children.Add(buttons);

        ok.Click += (_, _) => win.DialogResult = true;

        win.Content = root;
        var result = win.ShowDialog() == true;
        return (result, tb.Text ?? "");
    }

    private static void WriteCsv(DataTable table, string path)
    {
        using var writer = new StreamWriter(path);

        for (var i = 0; i < table.Columns.Count; i++)
        {
            if (i > 0) writer.Write(',');
            writer.Write(EscapeCsv(table.Columns[i].ColumnName));
        }
        writer.WriteLine();

        foreach (DataRow row in table.Rows)
        {
            for (var i = 0; i < table.Columns.Count; i++)
            {
                if (i > 0) writer.Write(',');
                var val = row[i] == DBNull.Value ? "" : Convert.ToString(row[i]) ?? "";
                writer.Write(EscapeCsv(val));
            }
            writer.WriteLine();
        }
    }

    private static string EscapeCsv(string s)
    {
        if (s.Contains('"') || s.Contains(',') || s.Contains('\n') || s.Contains('\r'))
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        return s;
    }

    private static string CycleTimeQuerySql(bool allOnly, string? filter = null)
    {
        // Definition:
        // - Link datasets by: MachineCycleTimeSnapshot.AssemblyNo == BuildToModuleMapping.ModuleNumber
        //   Unmatched rows from either side are ignored (INNER JOIN).
        // - BuildLineBottleneckSec: max MACHINE_MEDIUM_CYCLETIME across machines in the line for that build.
        // - TargetCycleTimeSec: dbo.Product.TargetCycleTimeSec (if present) keyed by PCB (BuildToModuleMapping.PCB).
        // - NOT_MEETING: bottleneck > target (and target exists).
        // - EXCEED: same as NOT_MEETING for now (separate button/report name).
        // This can be refined later when the factory-specific threshold rule is finalized.

        var where = "";
        if (!allOnly && string.Equals(filter, "NOT_MEETING", StringComparison.OrdinalIgnoreCase))
            where = "WHERE TargetCycleTimeSec IS NOT NULL AND ProductLineBottleneckSec > TargetCycleTimeSec";
        else if (!allOnly && string.Equals(filter, "EXCEED", StringComparison.OrdinalIgnoreCase))
            where = "WHERE TargetCycleTimeSec IS NOT NULL AND ProductLineBottleneckSec > TargetCycleTimeSec";

        return $"""
WITH base AS (
    SELECT
        b.BuildNumber AS BuildNumber,
        b.BuildKey AS BuildKey,
        b.ModuleNumber AS ModuleNumber,
        b.PCB AS ProductCode,
        p.TargetCycleTimeSec,
        s.LineCode,
        s.MachineName AS MachineCode,
        s.Side AS StageName,
        s.MachineMediumCycleTimeSec AS CycleTimeSec,
        s.CreatedAt AS MeasuredAt,
        s.AssemblyNo AS AssemblyNo
    FROM dbo.MachineCycleTimeSnapshot s
    INNER JOIN dbo.BuildToModuleMapping b
        ON UPPER(LTRIM(RTRIM(b.ModuleNumber))) = UPPER(LTRIM(RTRIM(s.AssemblyNo)))
    LEFT JOIN dbo.Product p ON p.ProductCode = b.PCB
    WHERE s.MachineMediumCycleTimeSec IS NOT NULL
),
lineAgg AS (
    SELECT
        BuildNumber,
        BuildKey,
        ModuleNumber,
        ProductCode,
        LineCode,
        MAX(TargetCycleTimeSec) AS TargetCycleTimeSec,
        MAX(CycleTimeSec) AS ProductLineBottleneckSec,
        MAX(MeasuredAt) AS LastUpdatedAt
    FROM base
    GROUP BY BuildNumber, BuildKey, ModuleNumber, ProductCode, LineCode
)
SELECT
    BuildNumber,
    BuildKey,
    ModuleNumber AS [Module#],
    ProductCode,
    LineCode,
    TargetCycleTimeSec,
    ProductLineBottleneckSec,
    LastUpdatedAt
FROM lineAgg
{where}
ORDER BY BuildNumber, LineCode;
""";
    }
}

