using System;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using SmtLineAllocationUI.DataAccess;
using SmtLineAllocationUI.Services;

namespace SmtLineAllocationUI;

public partial class SmtLineConfigurationMatrixPage : Page
{
    private readonly SqlConnectionFactory _connectionFactory = SqlConnectionFactory.FromAppSettings();
    private DataTable? _matrixTable;

    public SmtLineConfigurationMatrixPage()
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

            if (col.ColumnName is "LineId" or "MachineId")
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
            _matrixTable = await SmtLineConfigurationViewData.LoadSmtLineConfigurationMatrixAsync(_connectionFactory);
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
            GridData.LayoutUpdated -= Handler;
            HideIdColumns();
        }

        GridData.LayoutUpdated += Handler;
    }

    private void HideIdColumns()
    {
        foreach (var col in GridData.Columns)
        {
            var h = col.Header?.ToString();
            if (h is "LineId" or "MachineId")
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

            await SmtLineConfigurationMatrixPersistence.SaveMatrixAsync(_connectionFactory, _matrixTable);
            MessageBox.Show("Saved.", "Line configuration", MessageBoxButton.OK, MessageBoxImage.Information);
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
                r[c] = DBNull.Value;

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
        var hasMachine = row["MachineId"] != DBNull.Value && row["MachineId"] != null;

        if (hasMachine)
        {
            var confirm = MessageBox.Show(
                "Delete this machine from the database? Related nozzle/software rows for this machine will be removed.",
                "Confirm delete",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning
            );
            if (confirm != MessageBoxResult.Yes) return;

            try
            {
                BtnDeleteMatrixRow.IsEnabled = false;
                var machineId = Convert.ToInt32(row["MachineId"], CultureInfo.InvariantCulture);
                await SmtLineConfigurationMatrixPersistence.DeleteMachineRowAsync(_connectionFactory, machineId);
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
