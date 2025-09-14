using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Hardcodet.Wpf.TaskbarNotification;
using System.Management.Automation;
using System.ServiceProcess;
using System.Diagnostics; // Added for Debug.WriteLine

namespace Guifender;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private TaskbarIcon _notifyIcon;
    const int MajorVersion = 0;

    //*************************************************************************
    // MainWindow()
    //*************************************************************************
    public MainWindow()
    {
        InitializeComponent();

        // /!\ The VersionInfo class is auto-generated during build /!\
        this.Title = $"Guifender {MajorVersion}.{VersionInfo.CommitCount}-{VersionInfo.CommitHash}";

        // Diagnostic: Get resource stream once and check its validity
        Debug.WriteLine($"Attempting to get resource stream for: pack://application:,,,/app_icon.png");
        var resourceInfo = System.Windows.Application.GetResourceStream(new Uri("pack://application:,,,/app_icon.png"));
        if (resourceInfo == null || resourceInfo.Stream == null)
        {
            Debug.WriteLine("Resource 'app_icon.png' not found or stream is null.");
            System.Windows.MessageBox.Show("Resource 'app_icon.png' not found or stream is null.", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            return; // Exit if resource is not found
        }

        using (System.IO.Stream iconStream = resourceInfo.Stream)
        {
            Debug.WriteLine($"Resource stream found. Stream length: {iconStream.Length}");
            if (iconStream.Length == 0)
            {
                Debug.WriteLine("Resource stream for app_icon.png is empty.");
                System.Windows.MessageBox.Show("Resource stream for app_icon.png is empty.", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return; // Exit if resource stream is empty
            }

            // Diagnostic: Save the stream to a temporary file to inspect its content
            string tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "temp_app_icon.png");
            Debug.WriteLine($"Saving resource stream to: {tempFilePath}");
            // Ensure stream is at the beginning before copying
            iconStream.Seek(0, System.IO.SeekOrigin.Begin);
            using (System.IO.FileStream fs = new System.IO.FileStream(tempFilePath, System.IO.FileMode.Create))
            {
                iconStream.CopyTo(fs);
            }
            // Reset stream position for subsequent BitmapImage and Icon loading
            iconStream.Seek(0, System.IO.SeekOrigin.Begin);

            // Load app_icon.png for the main window icon
            BitmapImage mainWindowIcon = new BitmapImage();
            mainWindowIcon.BeginInit();
            mainWindowIcon.StreamSource = iconStream;
            mainWindowIcon.CacheOption = BitmapCacheOption.OnLoad;
            mainWindowIcon.EndInit();
            Icon = mainWindowIcon; // Assign the created BitmapImage to the Icon property
            iconStream.Seek(0, System.IO.SeekOrigin.Begin); // Reset stream for next use

            _notifyIcon = new TaskbarIcon();
            // Load app_icon.png for the system tray icon
            System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(iconStream);
            _notifyIcon.Icon = System.Drawing.Icon.FromHandle(bitmap.GetHicon());
        }
        _notifyIcon.ToolTipText = this.Title;
        _notifyIcon.TrayLeftMouseDown += (s, e) =>
        {
            Show();
            WindowState = WindowState.Normal;
        };

        var contextMenu = new ContextMenu();
        var hideShowMenuItem = new MenuItem { Header = "Hide / Show" };
        hideShowMenuItem.Click += (s, e) =>
        {
            if (this.IsVisible)
            {
                this.Hide();
            }
            else
            {
                this.Show();
                this.WindowState = WindowState.Normal;
            }
        };
        contextMenu.Items.Add(hideShowMenuItem);

        var exitMenuItem = new MenuItem { Header = "Exit" };
        exitMenuItem.Click += (s, e) => { System.Windows.Application.Current.Shutdown(); };
        contextMenu.Items.Add(exitMenuItem);
        _notifyIcon.ContextMenu = contextMenu;

        StateChanged += (s, e) =>
        {
            if (WindowState == WindowState.Minimized)
            {
                Hide();
            }
        };
        _ = InitializeDataAsync(); // Call the async helper method
    }

    //*************************************************************************
    // InitializeDataAsync()
    //*************************************************************************
    private async Task InitializeDataAsync()
    {
        ComputerStatusGrid.RowDefinitions.Clear(); // Clear existing rows once
        ComputerStatusGrid.Children.Clear(); // Clear existing children once
        await GetComputerStatus();
        GetServicesStatus();
    }

    //*************************************************************************
    // GetComputerStatus()
    //*************************************************************************
    private async Task GetComputerStatus()
    {
        Debug.WriteLine("GetComputerStatus() started.");
        try
        {
            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                PowerShellInstance.AddScript("Set-ExecutionPolicy RemoteSigned; Import-Module ConfigDefender; Get-MpComputerStatus | Select-Object AntivirusSignatureLastUpdated, AntivirusSignatureVersion, AMEngineVersion, AMProductVersion, AMRunningMode, AMServiceEnabled, AMServiceVersion, AntispywareSignatureVersion, AntispywareEnabled, AntispywareSignatureLastUpdated, RealTimeProtectionEnabled, RebootRequired");
                // PowerShellInstance.AddScript("Get-MpComputerStatus");
                var psOutput = await Task.Factory.FromAsync(PowerShellInstance.BeginInvoke(), PowerShellInstance.EndInvoke);
                StringBuilder sb = new StringBuilder();
                if (PowerShellInstance.Streams.Error.Count > 0)
                {
                    foreach (var error in PowerShellInstance.Streams.Error)
                    {
                        sb.AppendLine(error.ToString());
                    }
                    Debug.WriteLine($"PowerShell Errors: {sb.ToString()}");
                }
                else if (psOutput != null)
                {
                    foreach (var outputItem in psOutput)
                    {
                        if (outputItem != null)
                        {
                            foreach (var prop in outputItem.Properties)
                            {
                                sb.AppendLine($"{prop.Name}: {prop.Value}");
                            }
                        }
                    }
                    Debug.WriteLine($"PowerShell Output: {sb.ToString()}");
                }
                await Dispatcher.InvokeAsync(() =>
                {
                    Debug.WriteLine("Dispatcher.InvokeAsync in GetComputerStatus() started.");
                    int row = ComputerStatusGrid.RowDefinitions.Count; // Start adding from current row count
                    foreach (var line in sb.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var parts = line.Split(new[] { ':' }, 2);
                        if (parts.Length == 2)
                        {
                            ComputerStatusGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                            var label = new TextBlock { Text = parts[0].Trim() + ":", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 2, 5, 2) };
                            Grid.SetRow(label, row);
                            Grid.SetColumn(label, 0);
                            ComputerStatusGrid.Children.Add(label);

                            var value = new TextBlock { Text = parts[1].Trim(), Margin = new Thickness(0, 2, 0, 2) };
                            Grid.SetRow(value, row);
                            Grid.SetColumn(value, 1);
                            ComputerStatusGrid.Children.Add(value);

                            row++;
                        }
                    }
                    Debug.WriteLine("Dispatcher.InvokeAsync in GetComputerStatus() finished.");
                });
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Exception in GetComputerStatus(): {ex.Message}");
            await Dispatcher.InvokeAsync(() =>
            {
                ComputerStatusGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var errorText = new TextBlock { Text = $"Error: {ex.Message}", Foreground = System.Windows.Media.Brushes.Red, Margin = new Thickness(0, 2, 0, 2) };
                Grid.SetRow(errorText, ComputerStatusGrid.RowDefinitions.Count - 1);
                Grid.SetColumn(errorText, 0);
                Grid.SetColumnSpan(errorText, 2);
                ComputerStatusGrid.Children.Add(errorText);
            });
        }
        Debug.WriteLine("GetComputerStatus() finished.");
    }

    //*************************************************************************
    // GetAntiVirusStatus()
    //*************************************************************************
    private async Task GetAntiVirusStatus()
    {
        try
        {
            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                PowerShellInstance.AddScript("Set-ExecutionPolicy RemoteSigned; Import-Module ConfigDefender; Get-MpComputerStatus | Select-Object AntivirusSignatureLastUpdated, AntivirusSignatureVersion");
                // PowerShellInstance.AddScript("Get-MpComputerStatus");
                var psOutput = await Task.Factory.FromAsync(PowerShellInstance.BeginInvoke(), PowerShellInstance.EndInvoke);
                StringBuilder sb = new StringBuilder();
                if (PowerShellInstance.Streams.Error.Count > 0)
                {
                    foreach (var error in PowerShellInstance.Streams.Error)
                    {
                        sb.AppendLine(error.ToString());
                    }
                }
                else if (psOutput != null)
                {
                    foreach (var outputItem in psOutput)
                    {
                        if (outputItem != null)
                        {
                            foreach (var prop in outputItem.Properties)
                            {
                                sb.AppendLine($"{prop.Name}: {prop.Value}");
                            }
                        }
                    }
                }
                await Dispatcher.InvokeAsync(() =>
                {
                    int row = ComputerStatusGrid.RowDefinitions.Count; // Start adding from current row count
                    foreach (var line in sb.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var parts = line.Split(new[] { ':' }, 2);
                        if (parts.Length == 2)
                        {
                            ComputerStatusGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                            var label = new TextBlock { Text = parts[0].Trim() + ":", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 2, 5, 2) };
                            Grid.SetRow(label, row);
                            Grid.SetColumn(label, 0);
                            ComputerStatusGrid.Children.Add(label);

                            var value = new TextBlock { Text = parts[1].Trim(), Margin = new Thickness(0, 2, 0, 2) };
                            Grid.SetRow(value, row);
                            Grid.SetColumn(value, 1);
                            ComputerStatusGrid.Children.Add(value);

                            row++;
                        }
                    }
                });
            }
        }
        catch (Exception ex)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                ComputerStatusGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var errorText = new TextBlock { Text = $"Error: {ex.Message}", Foreground = System.Windows.Media.Brushes.Red, Margin = new Thickness(0, 2, 0, 2) };
                Grid.SetRow(errorText, ComputerStatusGrid.RowDefinitions.Count - 1);
                Grid.SetColumn(errorText, 0);
                Grid.SetColumnSpan(errorText, 2);
                ComputerStatusGrid.Children.Add(errorText);
            });
        }
    }

    //*************************************************************************
    // GetServicesStatus()
    //*************************************************************************
    private void GetServicesStatus()
    {
        StringBuilder sb = new StringBuilder();
        try
        {
            // These are the key Defender services to check
            string[] services = {"MDCoreSvc","mpssvc","Sense","WdNisSvc","WinDefend"};
            foreach (string serviceName in services)
            {
                string displayName;
                string status;
                (displayName,status) = GetServiceStatus(serviceName);
                sb.AppendLine($"Service '{displayName}' status: {status}");
            }
        }
        catch (Exception ex)
        {
            sb.AppendLine(ex.Message);
        }

        // Append Defender status to the ComputerStatusGrid
        Dispatcher.Invoke(() =>
        {
            ComputerStatusGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var headerLabel = new TextBlock { Text = "Defender Services:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 2) };
            Grid.SetRow(headerLabel, ComputerStatusGrid.RowDefinitions.Count - 1);
            Grid.SetColumn(headerLabel, 0);
            Grid.SetColumnSpan(headerLabel, 2);
            ComputerStatusGrid.Children.Add(headerLabel);

            foreach (var line in sb.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
            {
                var parts = line.Split(new[] { ':' }, 2);
                if (parts.Length == 2)
                {
                    string serviceName = parts[0].Trim(); // This is actually the display name, not the service name
                    string statusText = parts[1].Trim();

                    ComputerStatusGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                    var label = new TextBlock { Text = serviceName + ":", FontWeight = FontWeights.Normal, Margin = new Thickness(10, 2, 5, 2) };
                    Grid.SetRow(label, ComputerStatusGrid.RowDefinitions.Count - 1);
                    Grid.SetColumn(label, 0);
                    ComputerStatusGrid.Children.Add(label);

                    var value = new TextBlock { Text = statusText, Margin = new Thickness(0, 2, 0, 2) };

                    // Set color based on status
                    if (statusText.Equals("Running", StringComparison.OrdinalIgnoreCase))
                    {
                        value.Foreground = System.Windows.Media.Brushes.Green;
                    }
                    else
                    {
                        value.Foreground = System.Windows.Media.Brushes.Red;
                    }

                    Grid.SetRow(value, ComputerStatusGrid.RowDefinitions.Count - 1);
                    Grid.SetColumn(value, 1);
                    ComputerStatusGrid.Children.Add(value);
                }
            }
        });
    }

    //*************************************************************************
    // OnClosing()
    //*************************************************************************
    protected override void OnClosing(CancelEventArgs e)
    {
        e.Cancel = true; // Prevent the window from closing
        Hide();
    }

    //*************************************************************************
    // OnClosed()
    //*************************************************************************
    protected override void OnClosed(EventArgs e)
    {
        _notifyIcon.Dispose();
        base.OnClosed(e);
    }

    //*************************************************************************
    // GetServiceStatus()
    //*************************************************************************
    static protected (string,string) GetServiceStatus(String serviceName)
    {
        try
        {
            ServiceController sc = new ServiceController(serviceName);
            return (sc.DisplayName.ToString(),sc.Status.ToString());
        }
        catch (Exception e)
        {
            return ("Not found","Unknown");
        }

    }
}