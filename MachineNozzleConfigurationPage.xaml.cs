using System;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using SmtLineAllocationUI.DataAccess;
using SmtLineAllocationUI.Services;

namespace SmtLineAllocationUI;

public partial class MachineNozzleConfigurationPage : Page
{
    private readonly SqlConnectionFactory _connectionFactory = SqlConnectionFactory.FromAppSettings();
    private DataTable? _matrixTable;

    public MachineNozzleConfigurationPage()
    {
        InitializeComponent();
        Loaded += (_, _) => _ = LoadDataAsync();
    }

    private static void PrepareMatrixTableForEditing(DataTable table)
    {
        table.DefaultView.AllowNew = true;
        table.DefaultView.AllowDelete = true;

        foreach (DataColumn col in table.Columns)
        {
            col.AllowDBNull = true;

            if (col.ColumnName == MachineNozzleMatrixService.RowIdColumn)
            {
                col.AutoIncrement = false;
                col.ReadOnly = false;
            }
        }
    }

    private async Task LoadDataAsync()
    {
        try
        {
            _matrixTable = await MachineNozzleMatrixService.LoadWideMatrixAsync(_connectionFactory);
            PrepareMatrixTableForEditing(_matrixTable);
            GridData.ItemsSource = _matrixTable.DefaultView;
            SchedulePostBindColumnFix();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Query failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SchedulePostBindColumnFix()
    {
        void Handler(object? sender, EventArgs e)
        {
            if (GridData.Columns.Count == 0) return;

            GridData.LayoutUpdated -= Handler;
            ApplyGridColumnLayout();
        }

        GridData.LayoutUpdated += Handler;
    }

    private void ApplyGridColumnLayout()
    {
        HideIdColumn();

        var machineColIndex = -1;
        for (var i = 0; i < GridData.Columns.Count; i++)
        {
            var col = GridData.Columns[i];
            if (col.Visibility != Visibility.Visible) continue;

            var header = col.Header?.ToString() ?? "";

            if (header == MachineNozzleMatrixService.MachineIdColumn)
            {
                machineColIndex = i;
                col.Width = new DataGridLength(112);
                col.MinWidth = 88;
                col.CanUserResize = true;
                continue;
            }

            // Fixed width per nozzle type so many columns scroll horizontally instead of squeezing.
            col.Width = new DataGridLength(60);
            col.MinWidth = 52;
            col.MaxWidth = 80;
            col.CanUserResize = true;
        }

        // Keep Machine ID visible while scrolling nozzle columns left/right.
        GridData.FrozenColumnCount = machineColIndex >= 0 ? machineColIndex + 1 : 0;
    }

    private void HideIdColumn()
    {
        foreach (var col in GridData.Columns)
        {
            if (col.Header?.ToString() == MachineNozzleMatrixService.RowIdColumn)
                col.Visibility = Visibility.Collapsed;
        }
    }

    private async void BtnSaveMatrix_OnClick(object sender, RoutedEventArgs e)
    {
        if (_matrixTable is null) return;

        try
        {
            BtnSaveMatrix.IsEnabled = false;
            _ = GridData.CommitEdit(DataGridEditingUnit.Row, true);

            await MachineNozzleMatrixService.SaveWideMatrixAsync(_connectionFactory, _matrixTable);
            MessageBox.Show("Saved.", "Nozzle configuration", MessageBoxButton.OK, MessageBoxImage.Information);
            await LoadDataAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Save failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnSaveMatrix.IsEnabled = true;
        }
    }

    private void BtnAddMatrixRow_OnClick(object sender, RoutedEventArgs e)
    {
        if (_matrixTable is null) return;

        try
        {
            var r = _matrixTable.NewRow();
            foreach (DataColumn c in _matrixTable.Columns)
            {
                if (c.ColumnName == MachineNozzleMatrixService.RowIdColumn)
                    r[c] = DBNull.Value;
                else if (c.ColumnName == MachineNozzleMatrixService.MachineIdColumn)
                    r[c] = DBNull.Value;
                else if (c.DataType == typeof(int))
                    r[c] = 0;
                else
                    r[c] = DBNull.Value;
            }

            _matrixTable.Rows.Add(r);

            if (_matrixTable.DefaultView.Count > 0)
            {
                var last = _matrixTable.DefaultView[_matrixTable.DefaultView.Count - 1];
                GridData.SelectedItem = last;
                GridData.ScrollIntoView(last);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Add row failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void BtnDeleteMatrixRow_OnClick(object sender, RoutedEventArgs e)
    {
        if (_matrixTable is null) return;
        if (GridData.SelectedItem is not DataRowView drv)
        {
            MessageBox.Show("Select a row first.", "Delete", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var row = drv.Row;
        var hasRowId = row[MachineNozzleMatrixService.RowIdColumn] != DBNull.Value
                       && row[MachineNozzleMatrixService.RowIdColumn] != null;

        if (hasRowId)
        {
            var confirm = MessageBox.Show(
                "Delete this machine row from the database?",
                "Confirm delete",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning
            );
            if (confirm != MessageBoxResult.Yes) return;

            try
            {
                BtnDeleteMatrixRow.IsEnabled = false;
                var rowId = Convert.ToInt32(row[MachineNozzleMatrixService.RowIdColumn], CultureInfo.InvariantCulture);
                await MachineNozzleMatrixService.DeleteRowAsync(_connectionFactory, rowId);
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Delete failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnDeleteMatrixRow.IsEnabled = true;
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
            NavigationService?.Navigate(new SmtLineConfigurationPage());
    }
}
