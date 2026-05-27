using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using SmtLineAllocationUI.DataAccess;
using SmtLineAllocationUI.Services.Imports;

namespace SmtLineAllocationUI;

public partial class SmtLineConfigurationPage : Page
{
    private const string OptionKeyGlue = "GLUE_DISPENSER";
    private const string OptionKeyLargeBoard = "LARGE_BOARD_10IN";
    private const string OptionKeyReflow = "REFLOW_CENTER_SUPPORT";

    private readonly SqlConnectionFactory _connectionFactory = SqlConnectionFactory.FromAppSettings();
    private bool _constraintsLoading;
    private bool _constraintsUiReady;

    public SmtLineConfigurationPage()
    {
        InitializeComponent();
    }

    private async void Page_Loaded(object sender, RoutedEventArgs e)
    {
        _constraintsLoading = true;
        _constraintsUiReady = false;
        try
        {
            await LoadConstraintLinesAsync();
            await RebuildConstraintOptionPanelsAsync();

            if (CmbConstraintLine.SelectedItem is DataRowView drvl)
                await LoadLineConstraintChecksAsync(Convert.ToString(drvl.Row["LineCode"]) ?? "");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Load failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _constraintsLoading = false;
            _constraintsUiReady = true;
        }
    }

    private void BtnViewLines_OnClick(object sender, RoutedEventArgs e)
    {
        NavigationService?.Navigate(new SmtLineConfigurationMatrixPage());
    }

    private void BtnViewMachines_OnClick(object sender, RoutedEventArgs e)
    {
        NavigationService?.Navigate(new MachineNozzleConfigurationPage());
    }

    private void BtnDataActions_OnClick(object sender, RoutedEventArgs e)
    {
        var action = ShowChoice("Data actions", new[]
        {
            "Import machine list (update)",
            "Import nozzle matrix (update)",
            "Export configuration"
        });

        if (action is null) return;
        if (action == "Import machine list (update)") BtnImportMachineSheet_OnClick(sender, e);
        else if (action == "Import nozzle matrix (update)") BtnImportNozzleUpdate_OnClick(sender, e);
        else if (action == "Export configuration") BtnExportAllConfig_OnClick(sender, e);
    }

    private async Task LoadConstraintLinesAsync()
    {
        try
        {
            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var table = await QueryAsync("""
SELECT LineCode, LineName
FROM dbo.Line
WHERE IsActive = 1
ORDER BY LineCode;
""");

            CmbConstraintLine.DisplayMemberPath = "LineCode";
            CmbConstraintLine.SelectedValuePath = "LineCode";
            CmbConstraintLine.ItemsSource = table.DefaultView;

            if (CmbConstraintLine.Items.Count > 0)
                CmbConstraintLine.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Load failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task RebuildConstraintOptionPanelsAsync()
    {
        var initializer = new DbInitializer(_connectionFactory);
        await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

        var options = await QueryAsync("""
SELECT SmtConstraintOptionId AS ConstraintOptionId, OptionKey, DisplayName, SortOrder
FROM dbo.SmtConstraintOption
WHERE IsActive = 1
ORDER BY SortOrder, DisplayName;
""");

        PnlLineConstraintChecks.Children.Clear();

        foreach (DataRow row in options.Rows)
        {
            var id = Convert.ToInt32(row["ConstraintOptionId"], CultureInfo.InvariantCulture);
            var label = Convert.ToString(row["DisplayName"]) ?? "";

            var chkLine = new CheckBox
            {
                Content = label,
                Tag = id,
                Margin = new Thickness(0, 0, 0, 6)
            };
            PnlLineConstraintChecks.Children.Add(chkLine);
        }
    }

    private async void BtnManageConstraintOptions_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            BtnManageConstraintOptions.IsEnabled = false;
            var win = new ConstraintOptionsManageWindow(_connectionFactory)
            {
                Owner = Application.Current?.MainWindow
            };
            _ = win.ShowDialog();

            await RebuildConstraintOptionPanelsAsync();

            if (CmbConstraintLine.SelectedItem is DataRowView drvl)
                await LoadLineConstraintChecksAsync(Convert.ToString(drvl.Row["LineCode"]) ?? "");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Manage options failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnManageConstraintOptions.IsEnabled = true;
        }
    }

    private async void CmbConstraintLine_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_constraintsUiReady || _constraintsLoading) return;
        if (CmbConstraintLine.SelectedItem is not DataRowView drv) return;
        var lineCode = Convert.ToString(drv.Row["LineCode"]) ?? "";
        if (string.IsNullOrWhiteSpace(lineCode)) return;
        await LoadLineConstraintChecksAsync(lineCode);
    }

    private async Task<int?> GetLineIdByCodeAsync(string lineCode)
    {
        var table = await QueryAsync("""
SELECT LineId FROM dbo.Line WHERE LineCode = @lineCode;
""", new SqlParameter("@lineCode", lineCode));
        if (table.Rows.Count == 0) return null;
        return Convert.ToInt32(table.Rows[0]["LineId"], CultureInfo.InvariantCulture);
    }

    private async Task LoadLineConstraintChecksAsync(string lineCode)
    {
        try
        {
            _constraintsLoading = true;
            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var lineId = await GetLineIdByCodeAsync(lineCode);
            if (lineId is null)
            {
                SetAllPanelChecks(PnlLineConstraintChecks, isChecked: false);
                TxtConstraintRemark.Text = "";
                return;
            }

            var valueTable = await QueryAsync("""
SELECT co.SmtConstraintOptionId AS ConstraintOptionId, CAST(COALESCE(v.BitValue, 0) AS BIT) AS BitValue
FROM dbo.SmtConstraintOption co
LEFT JOIN dbo.SmtLineConstraintOptionValue v
    ON v.SmtConstraintOptionId = co.SmtConstraintOptionId AND v.LineId = @lineId
WHERE co.IsActive = 1;
""", new SqlParameter("@lineId", lineId.Value));

            ApplyValuesToPanel(PnlLineConstraintChecks, valueTable, "ConstraintOptionId", "BitValue");

            var remarkTable = await QueryAsync("""
SELECT lc.Remark
FROM dbo.LineConstraint lc
WHERE lc.LineId = @lineId;
""", new SqlParameter("@lineId", lineId.Value));

            if (remarkTable.Rows.Count == 0)
                TxtConstraintRemark.Text = "";
            else
                TxtConstraintRemark.Text = remarkTable.Rows[0]["Remark"] == DBNull.Value ? "" : Convert.ToString(remarkTable.Rows[0]["Remark"]) ?? "";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Load failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _constraintsLoading = false;
        }
    }

    private static void ApplyValuesToPanel(Panel panel, DataTable valueTable, string idColumn, string bitColumn)
    {
        var map = new Dictionary<int, bool>();
        foreach (DataRow r in valueTable.Rows)
        {
            var id = Convert.ToInt32(r[idColumn], CultureInfo.InvariantCulture);
            var bit = r[bitColumn] != DBNull.Value && (bool)r[bitColumn];
            map[id] = bit;
        }

        foreach (var child in panel.Children)
        {
            if (child is not CheckBox chk || chk.Tag is not int optId) continue;
            chk.IsChecked = map.TryGetValue(optId, out var b) && b;
        }
    }

    private static void SetAllPanelChecks(Panel panel, bool isChecked)
    {
        foreach (var child in panel.Children)
        {
            if (child is CheckBox chk)
                chk.IsChecked = isChecked;
        }
    }

    private async void BtnSaveConstraints_OnClick(object sender, RoutedEventArgs e)
    {
        if (CmbConstraintLine.SelectedItem is not DataRowView drv)
        {
            MessageBox.Show("Select a line first.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var lineCode = Convert.ToString(drv.Row["LineCode"]) ?? "";
        if (string.IsNullOrWhiteSpace(lineCode)) return;

        try
        {
            BtnSaveConstraints.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var lineId = await GetLineIdByCodeAsync(lineCode);
            if (lineId is null)
            {
                MessageBox.Show("Line not found.", "Save failed", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            await using var conn = _connectionFactory.Open();
            await using var tx = await conn.BeginTransactionAsync();

            try
            {
                await UpsertPanelOptionValuesAsync(conn, (SqlTransaction)tx, isLine: true, lineId.Value, machineId: null, PnlLineConstraintChecks);

                var machineTable = await QueryAsyncWithTransactionAsync(conn, (SqlTransaction)tx, """
SELECT MachineId
FROM dbo.Machine
WHERE LineId = @lineId AND IsActive = 1;
""", new SqlParameter("@lineId", lineId.Value));
                foreach (DataRow mr in machineTable.Rows)
                {
                    var mid = Convert.ToInt32(mr["MachineId"], CultureInfo.InvariantCulture);
                    await UpsertPanelOptionValuesAsync(conn, (SqlTransaction)tx, isLine: false, lineId: null, mid, PnlLineConstraintChecks);
                }

                var keyTable = await QueryAsyncWithTransactionAsync(conn, (SqlTransaction)tx, """
SELECT SmtConstraintOptionId AS ConstraintOptionId, OptionKey
FROM dbo.SmtConstraintOption
WHERE IsActive = 1;
""");

                bool BitForKey(string key)
                {
                    foreach (DataRow r in keyTable.Rows)
                    {
                        if (r["OptionKey"] == DBNull.Value) continue;
                        var k = Convert.ToString(r["OptionKey"]) ?? "";
                        if (!string.Equals(k, key, StringComparison.OrdinalIgnoreCase)) continue;
                        var id = Convert.ToInt32(r["ConstraintOptionId"], CultureInfo.InvariantCulture);
                        return GetPanelCheckState(PnlLineConstraintChecks, id) == true;
                    }

                    return false;
                }

                var glue = BitForKey(OptionKeyGlue);
                var large = BitForKey(OptionKeyLargeBoard);
                var reflow = BitForKey(OptionKeyReflow);

                await using (var cmdLc = new SqlCommand("""
IF EXISTS (SELECT 1 FROM dbo.LineConstraint WHERE LineId=@lineId)
BEGIN
    UPDATE dbo.LineConstraint
    SET GlueDispenserRequired=@glue,
        LargeBoardOver10InchSupported=@large,
        ReflowNonRetractableCenterSupport=@reflow,
        Remark=@remark,
        UpdatedAt=SYSDATETIME()
    WHERE LineId=@lineId;
END
ELSE
BEGIN
    INSERT INTO dbo.LineConstraint
        (LineId, GlueDispenserRequired, LargeBoardOver10InchSupported, ReflowNonRetractableCenterSupport, Remark)
    VALUES
        (@lineId, @glue, @large, @reflow, @remark);
END
""", conn, (SqlTransaction)tx))
                {
                    cmdLc.Parameters.AddWithValue("@lineId", lineId.Value);
                    cmdLc.Parameters.AddWithValue("@glue", glue);
                    cmdLc.Parameters.AddWithValue("@large", large);
                    cmdLc.Parameters.AddWithValue("@reflow", reflow);
                    cmdLc.Parameters.AddWithValue("@remark", string.IsNullOrWhiteSpace(TxtConstraintRemark.Text) ? (object)DBNull.Value : TxtConstraintRemark.Text.Trim());
                    await cmdLc.ExecuteNonQueryAsync();
                }

                await tx.CommitAsync();
            }
            catch
            {
                await tx.RollbackAsync();
                throw;
            }

            MessageBox.Show("Constraints saved (line and all machines on this line).", "Constraints", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Save failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnSaveConstraints.IsEnabled = true;
        }
    }

    private static bool? GetPanelCheckState(Panel panel, int optionId)
    {
        foreach (var child in panel.Children)
        {
            if (child is not CheckBox chk || chk.Tag is not int id || id != optionId) continue;
            return chk.IsChecked;
        }

        return false;
    }

    private static async Task UpsertPanelOptionValuesAsync(
        SqlConnection conn,
        SqlTransaction tx,
        bool isLine,
        int? lineId,
        int? machineId,
        Panel panel)
    {
        foreach (var child in panel.Children)
        {
            if (child is not CheckBox chk || chk.Tag is not int optionId) continue;
            var bit = chk.IsChecked == true;

            if (isLine)
            {
                await using var cmd = new SqlCommand("""
IF EXISTS (SELECT 1 FROM dbo.SmtLineConstraintOptionValue WHERE LineId=@lineId AND SmtConstraintOptionId=@optId)
    UPDATE dbo.SmtLineConstraintOptionValue SET BitValue=@bit, UpdatedAt=SYSDATETIME() WHERE LineId=@lineId AND SmtConstraintOptionId=@optId;
ELSE
    INSERT INTO dbo.SmtLineConstraintOptionValue (LineId, SmtConstraintOptionId, BitValue, UpdatedAt) VALUES (@lineId, @optId, @bit, SYSDATETIME());
""", conn, tx);
                cmd.Parameters.AddWithValue("@lineId", lineId!.Value);
                cmd.Parameters.AddWithValue("@optId", optionId);
                cmd.Parameters.AddWithValue("@bit", bit);
                await cmd.ExecuteNonQueryAsync();
            }
            else
            {
                await using var cmd = new SqlCommand("""
IF EXISTS (SELECT 1 FROM dbo.SmtMachineConstraintOptionValue WHERE MachineId=@machineId AND SmtConstraintOptionId=@optId)
    UPDATE dbo.SmtMachineConstraintOptionValue SET BitValue=@bit, UpdatedAt=SYSDATETIME() WHERE MachineId=@machineId AND SmtConstraintOptionId=@optId;
ELSE
    INSERT INTO dbo.SmtMachineConstraintOptionValue (MachineId, SmtConstraintOptionId, BitValue, UpdatedAt) VALUES (@machineId, @optId, @bit, SYSDATETIME());
""", conn, tx);
                cmd.Parameters.AddWithValue("@machineId", machineId!.Value);
                cmd.Parameters.AddWithValue("@optId", optionId);
                cmd.Parameters.AddWithValue("@bit", bit);
                await cmd.ExecuteNonQueryAsync();
            }
        }
    }

    private static async Task<DataTable> QueryAsyncWithTransactionAsync(SqlConnection conn, SqlTransaction tx, string sql, params SqlParameter[] parameters)
    {
        await using var cmd = new SqlCommand(sql, conn, tx);
        if (parameters is { Length: > 0 })
            cmd.Parameters.AddRange(parameters);

        await using var reader = await cmd.ExecuteReaderAsync();
        var table = new DataTable();
        table.Load(reader);
        return table;
    }

    private async void BtnViewAllConstraints_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            BtnViewAllConstraints.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var table = await QueryAsync("""
SELECT
    l.LineCode AS Line,
    lc.Remark AS Remark,
    co.DisplayName AS ConstraintOption,
    CAST(COALESCE(v.BitValue, 0) AS BIT) AS Enabled
FROM dbo.Line l
INNER JOIN dbo.SmtConstraintOption co ON co.IsActive = 1
LEFT JOIN dbo.SmtLineConstraintOptionValue v
    ON v.LineId = l.LineId AND v.SmtConstraintOptionId = co.SmtConstraintOptionId
LEFT JOIN dbo.LineConstraint lc ON lc.LineId = l.LineId
ORDER BY l.LineCode, co.SortOrder, co.DisplayName;
""");

            var popup = new Window
            {
                Title = "All line constraints",
                Width = 960,
                Height = 560,
                Owner = Application.Current?.MainWindow,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };
            var grid = new DataGrid
            {
                ItemsSource = table.DefaultView,
                IsReadOnly = true,
                AutoGenerateColumns = true,
                Margin = new Thickness(12),
                CanUserAddRows = false
            };
            popup.Content = grid;
            popup.ShowDialog();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Query failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnViewAllConstraints.IsEnabled = true;
        }
    }

    private static string? ShowChoice(string title, IReadOnlyList<string> options)
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

    private async void BtnImportNozzleUpdate_OnClick(object sender, RoutedEventArgs e)
    {
        await ImportNozzleConfigAsync();
    }

    private async Task ImportNozzleConfigAsync()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select SMT Machine nozzle configuration Excel",
            Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
            InitialDirectory = GetDefaultDataDirectory()
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            BtnDataActions.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var importer = new NozzleMatrixExcelImporter(_connectionFactory);
            var result = await importer.ImportAsync(dlg.FileName);

            MessageBox.Show(
                $"Imported: {result.FileName}\nRead: {result.ReadCount}\nSuccess: {result.SuccessCount}\nFailed: {result.FailureCount}",
                "Import result",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );

            if (result.FailureCount > 0 && result.Errors.Count > 0)
            {
                var preview = string.Join("\n", result.Errors.Take(10));
                MessageBox.Show(
                    $"Import failed rows (first 10):\n{preview}",
                    "Import errors",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Import failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnDataActions.IsEnabled = true;
        }
    }

    private async void BtnImportMachineSheet_OnClick(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select SMT line manager Project Statement Of Work.xlsx",
            Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
            InitialDirectory = GetDefaultDataDirectory()
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            BtnDataActions.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var importer = new MachineSoftwareVerExcelImporter(_connectionFactory);
            var result = await importer.ImportAsync(dlg.FileName, sheetName: "3) Machine software ver");

            MessageBox.Show(
                $"Imported: {result.FileName}\nRead: {result.ReadCount}\nSuccess: {result.SuccessCount}\nFailed: {result.FailureCount}",
                "Import result",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );

            if (result.FailureCount > 0 && result.Errors.Count > 0)
            {
                var preview = string.Join("\n", result.Errors.Take(10));
                MessageBox.Show(
                    $"Import failed rows (first 10):\n{preview}",
                    "Import errors",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Import failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnDataActions.IsEnabled = true;
        }
    }

    private async void BtnExportAllConfig_OnClick(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog
        {
            Title = "Export configuration to CSV",
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = "smt_line_configuration_export.csv"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            BtnDataActions.IsEnabled = false;

            var table = await QueryAsync("""
SELECT
    l.LineCode AS [Process Line],
    m.MachineCode AS [Machine ID],
    m.MachineType AS [Machine Model],
    m.EquipType AS [Equipment Type]
FROM dbo.Line l
LEFT JOIN dbo.Machine m ON m.LineId = l.LineId
ORDER BY l.LineCode, m.MachineCode;
""");

            WriteCsv(table, dlg.FileName);
            MessageBox.Show("Export completed.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnDataActions.IsEnabled = true;
        }
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
                var val = row[i] == DBNull.Value ? "" : Convert.ToString(row[i], CultureInfo.InvariantCulture) ?? "";
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
}

