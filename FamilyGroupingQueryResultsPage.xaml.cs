using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;
using SmtLineAllocationUI.Services;

namespace SmtLineAllocationUI;

public partial class FamilyGroupingQueryResultsPage : Page
{
    private readonly SqlConnectionFactory _connectionFactory = SqlConnectionFactory.FromAppSettings();
    private readonly string _filterField;
    private readonly string _filterValue;

    public FamilyGroupingQueryResultsPage(string filterField, string filterValue)
    {
        _filterField = filterField ?? throw new ArgumentNullException(nameof(filterField));
        _filterValue = (filterValue ?? "").Trim();

        InitializeComponent();
        TxtSubtitle.Text = $"{_filterField} = {_filterValue}";
        Loaded += (_, _) => _ = LoadAsync();
    }

    private async Task LoadAsync()
    {
        if (string.IsNullOrWhiteSpace(_filterValue))
        {
            MessageBox.Show("Search value cannot be empty.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            GridData.ItemsSource = null;
            return;
        }

        try
        {
            await FamilyGroupingViewData.EnsureTableAsync(_connectionFactory);

            var sql = FamilyGroupingViewData.BuildFilterSql(_filterField);
            var table = await QueryAsync(sql, new SqlParameter("@v", _filterValue));
            GridData.ItemsSource = table.DefaultView;
            SchedulePostBindColumnFix();

            if (table.Rows.Count == 0)
            {
                MessageBox.Show(
                    $"No records found for {_filterField} '{_filterValue}'.",
                    "Query result",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
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
            ApplyColumnLayout();
        }

        GridData.LayoutUpdated += Handler;
    }

    private void ApplyColumnLayout()
    {
        if (GridData.Columns.Count == 0) return;

        foreach (var col in GridData.Columns)
        {
            var header = col.Header?.ToString() ?? "";
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

        GridData.FrozenColumnCount = Math.Min(2, GridData.Columns.Count);
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
            NavigationService?.Navigate(new ProductLineAllocationPage());
    }
}
