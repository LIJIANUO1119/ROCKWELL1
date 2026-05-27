using System;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using SmtLineAllocationUI.DataAccess;
using SmtLineAllocationUI.Services;

namespace SmtLineAllocationUI;

public partial class FamilyGroupingConfigurationPage : Page
{
    private readonly SqlConnectionFactory _connectionFactory = SqlConnectionFactory.FromAppSettings();
    private DataTable? _table;

    public FamilyGroupingConfigurationPage()
    {
        InitializeComponent();
        Loaded += (_, _) => _ = LoadDataAsync();
    }

    private static void PrepareTableForEditing(DataTable table)
    {
        table.DefaultView.AllowNew = true;
        table.DefaultView.AllowDelete = true;

        foreach (DataColumn col in table.Columns)
        {
            col.AllowDBNull = true;
            if (col.ColumnName == "FamilyGroupingDetailId")
                col.ReadOnly = true;
        }
    }

    private async Task LoadDataAsync()
    {
        try
        {
            _table = await FamilyGroupingViewData.LoadGridAsync(_connectionFactory);
            PrepareTableForEditing(_table);
            GridData.ItemsSource = _table.DefaultView;
            TxtSubtitle.Text = $"{_table.Rows.Count} row(s)";
            SchedulePostBindColumnFix();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Load failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SchedulePostBindColumnFix()
    {
        void Handler(object? sender, EventArgs e)
        {
            GridData.LayoutUpdated -= Handler;
            HideIdColumn();
            ApplyColumnLayout();
        }

        GridData.LayoutUpdated += Handler;
    }

    private void HideIdColumn()
    {
        foreach (var col in GridData.Columns)
        {
            if (col.Header?.ToString() == "FamilyGroupingDetailId")
                col.Visibility = Visibility.Collapsed;
        }
    }

    private void ApplyColumnLayout()
    {
        if (GridData.Columns.Count == 0) return;

        foreach (var col in GridData.Columns)
        {
            var header = col.Header?.ToString() ?? "";
            if (header == "FamilyGroupingDetailId") continue;

            if (header is "ASY" or "PCB" or "Family")
            {
                col.Width = new DataGridLength(100);
                col.MinWidth = 72;
            }
            else if (header is "Family#")
            {
                col.Width = new DataGridLength(72);
            }
            else if (header.Contains("Priority", StringComparison.OrdinalIgnoreCase))
            {
                col.Width = new DataGridLength(88);
                col.MinWidth = 72;
            }
            else if (header.Contains("CycleTime", StringComparison.OrdinalIgnoreCase)
                     || header == "# Circuits")
            {
                col.Width = new DataGridLength(80);
            }
            else
            {
                col.Width = new DataGridLength(90);
            }

            col.CanUserResize = true;
        }

        var visible = 0;
        foreach (var col in GridData.Columns)
            if (col.Visibility == Visibility.Visible) visible++;
        GridData.FrozenColumnCount = Math.Min(2, visible);
    }

    private async void BtnSave_OnClick(object sender, RoutedEventArgs e)
    {
        if (_table is null) return;

        try
        {
            BtnSave.IsEnabled = false;
            _ = GridData.CommitEdit(DataGridEditingUnit.Row, true);

            await FamilyGroupingPersistence.SaveGridAsync(_connectionFactory, _table);
            MessageBox.Show("Saved.", "Family groupings", MessageBoxButton.OK, MessageBoxImage.Information);
            await LoadDataAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Save failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnSave.IsEnabled = true;
        }
    }

    private void BtnAddRow_OnClick(object sender, RoutedEventArgs e)
    {
        if (_table is null) return;

        try
        {
            var r = _table.NewRow();
            foreach (DataColumn c in _table.Columns)
                r[c] = DBNull.Value;

            _table.Rows.Add(r);

            if (_table.DefaultView.Count > 0)
            {
                var last = _table.DefaultView[_table.DefaultView.Count - 1];
                GridData.SelectedItem = last;
                GridData.ScrollIntoView(last);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Add row failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void BtnDeleteRow_OnClick(object sender, RoutedEventArgs e)
    {
        if (_table is null) return;
        if (GridData.SelectedItem is not DataRowView drv)
        {
            MessageBox.Show("Select a row first.", "Delete", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var row = drv.Row;
        var hasId = row["FamilyGroupingDetailId"] != DBNull.Value && row["FamilyGroupingDetailId"] != null;

        if (hasId)
        {
            var confirm = MessageBox.Show(
                "Delete this row from the database?",
                "Confirm delete",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;

            try
            {
                BtnDeleteRow.IsEnabled = false;
                var id = Convert.ToInt32(row["FamilyGroupingDetailId"], CultureInfo.InvariantCulture);
                await FamilyGroupingPersistence.DeleteRowAsync(_connectionFactory, id);
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Delete failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnDeleteRow.IsEnabled = true;
            }

            return;
        }

        row.Delete();
    }

    private void BtnBack_OnClick(object sender, RoutedEventArgs e)
    {
        if (NavigationService?.CanGoBack == true)
            NavigationService.GoBack();
        else
            NavigationService?.Navigate(new ProductLineAllocationPage());
    }
}
