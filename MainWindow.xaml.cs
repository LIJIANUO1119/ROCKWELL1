using System.Windows;
using System.Windows.Controls;

namespace SmtLineAllocationUI;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        NavigateToSmtLineConfiguration();
    }

    private void NavigateToSmtLineConfiguration()
    {
        ContentFrame.Navigate(new SmtLineConfigurationPage());
    }

    private void NavigateToProductCycleTime()
    {
        ContentFrame.Navigate(new ProductCycleTimePage());
    }

    private void NavigateToProductLineAllocation()
    {
        ContentFrame.Navigate(new ProductLineAllocationPage());
    }

    private void BtnSmtLineConfig_OnClick(object sender, RoutedEventArgs e)
    {
        NavigateToSmtLineConfiguration();
    }

    private void BtnProductCycleTime_OnClick(object sender, RoutedEventArgs e)
    {
        NavigateToProductCycleTime();
    }

    private void BtnProductLineAllocation_OnClick(object sender, RoutedEventArgs e)
    {
        NavigateToProductLineAllocation();
    }
}

