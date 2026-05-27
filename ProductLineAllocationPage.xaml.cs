using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using SmtLineAllocationUI.DataAccess;
using SmtLineAllocationUI.Services;
using SmtLineAllocationUI.Services.Imports;

namespace SmtLineAllocationUI;

public partial class ProductLineAllocationPage : Page
{
    private readonly SqlConnectionFactory _connectionFactory = SqlConnectionFactory.FromAppSettings();

    public ProductLineAllocationPage()
    {
        InitializeComponent();
    }

    private async void BtnQueryAllocation_OnClick(object sender, RoutedEventArgs e)
    {
        var option = ShowChoice("Query family groupings by", new[] { "ASY", "Family" });
        if (option is null) return;

        var (ok, value) = ShowInput($"Enter {option}", "");
        if (!ok) return;
        value = (value ?? "").Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            MessageBox.Show("Input cannot be empty.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            BtnQueryAllocation.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            NavigationService?.Navigate(new FamilyGroupingQueryResultsPage(option, value));
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Query failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnQueryAllocation.IsEnabled = true;
        }
    }

    private async void BtnExportAllocation_OnClick(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog
        {
            Title = "Export family groupings to CSV",
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = "FAMILY_GROUPINGS.csv"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            BtnExportAllocation.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");
            await FamilyGroupingViewData.EnsureTableAsync(_connectionFactory);

            var table = await QueryAsync(FamilyGroupingViewData.ExportSql);
            WriteCsv(table, dlg.FileName);

            MessageBox.Show(
                $"Export completed.\nRows: {table.Rows.Count}",
                "Export",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnExportAllocation.IsEnabled = true;
        }
    }

    private async void BtnUploadFamilyGroupings_OnClick(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select SGP_FAMILY_GROUPINGS_DETAIL (CSV or Excel)",
            Filter = "CSV and Excel|*.csv;*.xlsx;*.xlsm|CSV files (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*",
            InitialDirectory = GetDefaultDataDirectory()
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            BtnUploadFamilyGroupings.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var importer = new FamilyGroupingsCsvImporter(_connectionFactory);
            var result = await importer.ImportAsync(dlg.FileName);

            var detail = result.Errors.Count > 0
                ? "\n\n" + string.Join("\n", result.Errors.Take(20))
                : "";
            if (result.Errors.Count > 20)
                detail += $"\n... and {result.Errors.Count - 20} more.";

            MessageBox.Show(
                $"Imported: {result.FileName}\nRead: {result.ReadCount}\nSaved: {result.SuccessCount}\nIssues: {result.FailureCount}{detail}",
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
            BtnUploadFamilyGroupings.IsEnabled = true;
        }
    }

    private async void BtnViewFamilyGroupings_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            BtnViewFamilyGroupings.IsEnabled = false;

            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            NavigationService?.Navigate(new FamilyGroupingConfigurationPage());
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Open failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnViewFamilyGroupings.IsEnabled = true;
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
        using var writer = new StreamWriter(path, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

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

    private static (bool ok, string? value) ShowInput(string title, string defaultValue)
    {
        var win = new Window
        {
            Title = title,
            SizeToContent = SizeToContent.WidthAndHeight,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
            Owner = Application.Current?.MainWindow
        };

        var root = new Grid { Margin = new Thickness(16) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var form = new Grid();
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(320) });

        var tb = new TextBox { Text = defaultValue ?? "", Margin = new Thickness(0, 0, 0, 8), MinWidth = 280 };
        Grid.SetRow(tb, 0);
        Grid.SetColumn(tb, 1);
        form.Children.Add(tb);

        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 8, 0, 0) };
        var btnOk = new Button { Content = "Confirm", Width = 110, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
        var btnCancel = new Button { Content = "Cancel", Width = 110, IsCancel = true };
        buttons.Children.Add(btnOk);
        buttons.Children.Add(btnCancel);
        btnOk.Click += (_, _) => win.DialogResult = true;

        Grid.SetRow(form, 0);
        Grid.SetRow(buttons, 1);
        root.Children.Add(form);
        root.Children.Add(buttons);
        win.Content = root;

        var result = win.ShowDialog() == true;
        return (result, tb.Text);
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
