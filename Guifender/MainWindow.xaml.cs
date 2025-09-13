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

namespace Guifender;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private TaskbarIcon _notifyIcon;

    //*************************************************************************
    // MainWindow()
    //*************************************************************************
    public MainWindow()
    {
        InitializeComponent();
        _ = GetComputerStatus();
        GetDefenderStatus();

        Icon = Imaging.CreateBitmapSourceFromHIcon(
            SystemIcons.Shield.Handle,
            Int32Rect.Empty,
            BitmapSizeOptions.FromEmptyOptions());

        _notifyIcon = new TaskbarIcon();
        _notifyIcon.Icon = SystemIcons.Shield;
        _notifyIcon.ToolTipText = "Guifender";
        _notifyIcon.TrayLeftMouseDown += (s, e) =>
        {
            Show();
            WindowState = WindowState.Normal;
        };

        var contextMenu = new ContextMenu();
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
    }

    //*************************************************************************
    // GetComputerStatus()
    //*************************************************************************
    private async Task GetComputerStatus()
    {
        try
        {
            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                PowerShellInstance.AddScript("Set-ExecutionPolicy RemoteSigned; Import-Module ConfigDefender; Get-MpComputerStatus");
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
                    ComputerStatusTextBlock.Text += sb.ToString();
                });
            }
        }
        catch (Exception e)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                ComputerStatusTextBlock.Text = e.Message;
            });
        }
    }

    //*************************************************************************
    // GetDefenderStatus()
    //*************************************************************************
    private void GetDefenderStatus()
    {
        StringBuilder sb = new StringBuilder();
        try
        {
            string[] services = {"MDCoreSvc","mpssvc","Sense","WdNisSvc","WinDefend"};
            foreach (string serviceName in services)
            {
                string displayName;
                string status;
                (displayName,status) = GetServiceStatus(serviceName);
                sb.AppendLine($"Service '{displayName}' status: {status}");
                // ComputerStatusTextBlock.Text = sb.ToString();
            }
        }
        catch (Exception e)
        {
            sb.AppendLine(e.Message);
        }
        ComputerStatusTextBlock.Text = sb.ToString();
    }

    //*************************************************************************
    // OnClosing()
    //*************************************************************************
    protected override void OnClosing(CancelEventArgs e)
    {
        Hide();
        base.OnClosing(e);
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
    protected (string,string) GetServiceStatus(String serviceName)
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