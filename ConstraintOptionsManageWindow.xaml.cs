using System;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI;

public partial class ConstraintOptionsManageWindow : Window
{
    private readonly SqlConnectionFactory _connectionFactory;

    public ConstraintOptionsManageWindow(SqlConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        InitializeComponent();
        _ = LoadGridAsync();
    }

    private async Task LoadGridAsync()
    {
        try
        {
            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            var table = await QueryAsync("""
SELECT
    SmtConstraintOptionId AS ConstraintOptionId,
    DisplayName,
    SortOrder,
    IsActive
FROM dbo.SmtConstraintOption
ORDER BY SortOrder, DisplayName;
""");
            GridOptions.ItemsSource = table.DefaultView;
            if (table.Rows.Count > 0)
                GridOptions.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Load failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void BtnAdd_OnClick(object sender, RoutedEventArgs e)
    {
        if (!PromptOptionFields(isEdit: false, out var displayName, out var sortOrder, out var isActive))
            return;

        try
        {
            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            await using var conn = _connectionFactory.Open();
            await using var cmd = new SqlCommand("""
INSERT INTO dbo.SmtConstraintOption (OptionKey, DisplayName, SortOrder, IsActive, UpdatedAt)
VALUES (NULL, @name, @sort, @active, SYSDATETIME());
""", conn);
            cmd.Parameters.AddWithValue("@name", displayName.Trim());
            cmd.Parameters.AddWithValue("@sort", sortOrder);
            cmd.Parameters.AddWithValue("@active", isActive);
            await cmd.ExecuteNonQueryAsync();
            await LoadGridAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Add failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void BtnEdit_OnClick(object sender, RoutedEventArgs e)
    {
        if (GridOptions.SelectedItem is not DataRowView drv)
        {
            MessageBox.Show("Select a row first.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var id = Convert.ToInt32(drv.Row["ConstraintOptionId"], CultureInfo.InvariantCulture);
        var currentName = Convert.ToString(drv.Row["DisplayName"]) ?? "";
        var currentSort = Convert.ToInt32(drv.Row["SortOrder"], CultureInfo.InvariantCulture);
        var currentActive = drv.Row["IsActive"] != DBNull.Value && (bool)drv.Row["IsActive"];

        if (!PromptOptionFields(isEdit: true, out var displayName, out var sortOrder, out var isActive, currentName, currentSort, currentActive))
            return;

        try
        {
            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            await using var conn = _connectionFactory.Open();
            await using var cmd = new SqlCommand("""
UPDATE dbo.SmtConstraintOption
SET DisplayName = @name,
    SortOrder = @sort,
    IsActive = @active,
    UpdatedAt = SYSDATETIME()
WHERE SmtConstraintOptionId = @id;
""", conn);
            cmd.Parameters.AddWithValue("@id", id);
            cmd.Parameters.AddWithValue("@name", displayName.Trim());
            cmd.Parameters.AddWithValue("@sort", sortOrder);
            cmd.Parameters.AddWithValue("@active", isActive);
            await cmd.ExecuteNonQueryAsync();
            await LoadGridAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Update failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void BtnDelete_OnClick(object sender, RoutedEventArgs e)
    {
        if (GridOptions.SelectedItem is not DataRowView drv)
        {
            MessageBox.Show("Select a row first.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var id = Convert.ToInt32(drv.Row["ConstraintOptionId"], CultureInfo.InvariantCulture);
        var name = Convert.ToString(drv.Row["DisplayName"]) ?? "";

        var confirm = MessageBox.Show(
            $"Delete constraint option \"{name}\"?\n\nLine and machine values for this option will be removed.",
            "Confirm delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning
        );
        if (confirm != MessageBoxResult.Yes) return;

        try
        {
            var initializer = new DbInitializer(_connectionFactory);
            await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql");

            await using var conn = _connectionFactory.Open();
            await using var cmd = new SqlCommand("DELETE FROM dbo.SmtConstraintOption WHERE SmtConstraintOptionId = @id;", conn);
            cmd.Parameters.AddWithValue("@id", id);
            await cmd.ExecuteNonQueryAsync();
            await LoadGridAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Delete failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnClose_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private static bool PromptOptionFields(
        bool isEdit,
        out string displayName,
        out int sortOrder,
        out bool isActive,
        string initialName = "",
        int initialSort = 100,
        bool initialActive = true)
    {
        displayName = "";
        sortOrder = initialSort;
        isActive = initialActive;

        var win = new Window
        {
            Title = isEdit ? "Edit constraint option" : "Add constraint option",
            Width = 420,
            SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
            Owner = Application.Current?.MainWindow
        };

        var nameBox = new System.Windows.Controls.TextBox
        {
            Text = initialName,
            Margin = new Thickness(0, 4, 0, 0),
            Height = 28
        };
        var sortBox = new System.Windows.Controls.TextBox
        {
            Text = initialSort.ToString(CultureInfo.InvariantCulture),
            Margin = new Thickness(0, 4, 0, 0),
            Height = 28
        };
        var activeChk = new System.Windows.Controls.CheckBox
        {
            Content = "Active (shown when assigning to lines/machines)",
            IsChecked = initialActive,
            Margin = new Thickness(0, 12, 0, 0)
        };

        var root = new System.Windows.Controls.StackPanel { Margin = new Thickness(16) };
        root.Children.Add(new System.Windows.Controls.TextBlock { Text = "Display name", FontWeight = FontWeights.SemiBold });
        root.Children.Add(nameBox);
        root.Children.Add(new System.Windows.Controls.TextBlock { Text = "Sort order (lower appears first)", Margin = new Thickness(0, 12, 0, 0), FontWeight = FontWeights.SemiBold });
        root.Children.Add(sortBox);
        root.Children.Add(activeChk);

        var btnRow = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 16, 0, 0)
        };
        var ok = new System.Windows.Controls.Button { Content = "OK", Width = 88, Height = 30, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
        var cancel = new System.Windows.Controls.Button { Content = "Cancel", Width = 88, Height = 30, IsCancel = true };
        btnRow.Children.Add(ok);
        btnRow.Children.Add(cancel);
        root.Children.Add(btnRow);

        win.Content = root;

        var pendingName = "";
        var pendingSort = initialSort;
        var pendingActive = initialActive;
        var accepted = false;
        ok.Click += (_, _) =>
        {
            pendingName = nameBox.Text ?? "";
            if (string.IsNullOrWhiteSpace(pendingName))
            {
                _ = MessageBox.Show("Display name is required.", "Validation", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!int.TryParse((sortBox.Text ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedSort))
            {
                _ = MessageBox.Show("Sort order must be an integer.", "Validation", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            pendingSort = parsedSort;
            pendingActive = activeChk.IsChecked == true;
            accepted = true;
            win.DialogResult = true;
        };
        cancel.Click += (_, _) => { win.DialogResult = false; };

        _ = win.ShowDialog();
        if (!accepted)
        {
            displayName = "";
            sortOrder = initialSort;
            isActive = initialActive;
            return false;
        }

        displayName = pendingName;
        sortOrder = pendingSort;
        isActive = pendingActive;
        return true;
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
}
