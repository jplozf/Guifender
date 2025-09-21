using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Management.Automation;
using System.Runtime.CompilerServices;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Hardcodet.Wpf.TaskbarNotification;
using Newtonsoft.Json;

namespace Guifender
{
    public class AppSettings
    {
        public bool ConfirmOnExit { get; set; } = true;
        public bool MinimizeToTrayOnClose { get; set; } = true;
        public string SelectedUpdateSource { get; set; } = "MicrosoftUpdateServer";
        public int ThrottleLimit { get; set; } = 0;
        public bool AsJob { get; set; } = false;
        public double WindowTop { get; set; } = 100;
        public double WindowLeft { get; set; } = 100;
        public double WindowHeight { get; set; } = 500;
        public double WindowWidth { get; set; } = 850;
        public WindowState WindowState { get; set; } = WindowState.Normal;
        public int SelectedTabIndex { get; set; } = 0;
    }

    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private TaskbarIcon _notifyIcon;
        const int MajorVersion = 0;
        private AppSettings _settings;
        private readonly string _settingsFilePath;
        private CancellationTokenSource _statusClearTokenSource = new CancellationTokenSource();
        private bool _isExiting = false;

        #region Properties
        private string _statusText;
        public string StatusText
        {
            get => _statusText;
            set { _statusText = value; OnPropertyChanged(); }
        }

        private bool _confirmOnExit;
        public bool ConfirmOnExit
        {
            get => _confirmOnExit;
            set { _confirmOnExit = value; OnPropertyChanged(); }
        }

        private bool _minimizeToTrayOnClose;
        public bool MinimizeToTrayOnClose
        {
            get => _minimizeToTrayOnClose;
            set { _minimizeToTrayOnClose = value; OnPropertyChanged(); }
        }

        private List<string> _updateSources;
        public List<string> UpdateSources
        {
            get => _updateSources;
            set { _updateSources = value; OnPropertyChanged(); }
        }

        private string _selectedUpdateSource;
        public string SelectedUpdateSource
        {
            get => _selectedUpdateSource;
            set { _selectedUpdateSource = value; OnPropertyChanged(); }
        }

        private int _throttleLimit;
        public int ThrottleLimit
        {
            get => _throttleLimit;
            set { _throttleLimit = value; OnPropertyChanged(); }
        }

        private bool _asJob;
        public bool AsJob
        {
            get => _asJob;
            set { _asJob = value; OnPropertyChanged(); }
        }

        private string _updateOutputText;
        public string UpdateOutputText
        {
            get => _updateOutputText;
            set { _updateOutputText = value; OnPropertyChanged(); }
        }
        #endregion

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        public MainWindow()
        {
            InitializeComponent();
            UpdateSources = new List<string> { "MicrosoftUpdateServer", "InternalDefinitionUpdateServer", "MMPC" };

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = Path.Combine(appDataPath, "Guifender");
            Directory.CreateDirectory(settingsDir);
            _settingsFilePath = Path.Combine(settingsDir, "settings.json");

            LoadSettings();

            this.Title = $"Guifender {MajorVersion}.{VersionInfo.CommitCount}-{VersionInfo.CommitHash}";
            SetupTaskbarIcon();
            SetStatus("Ready", clearAfter: false);

            _ = InitializeDataAsync();
        }

        private void SetupTaskbarIcon()
        {
            var resourceInfo = System.Windows.Application.GetResourceStream(new Uri("pack://application:,,,/app_icon.png"));
            if (resourceInfo == null || resourceInfo.Stream == null)
            {
                System.Windows.MessageBox.Show("Resource 'app_icon.png' not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            using (Stream iconStream = resourceInfo.Stream)
            {
                BitmapImage mainWindowIcon = new BitmapImage();
                mainWindowIcon.BeginInit();
                mainWindowIcon.StreamSource = iconStream;
                mainWindowIcon.CacheOption = BitmapCacheOption.OnLoad;
                mainWindowIcon.EndInit();
                Icon = mainWindowIcon;
                iconStream.Seek(0, SeekOrigin.Begin);

                _notifyIcon = new TaskbarIcon();
                System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(iconStream);
                _notifyIcon.Icon = System.Drawing.Icon.FromHandle(bitmap.GetHicon());
            }

            _notifyIcon.ToolTipText = this.Title;
            _notifyIcon.TrayLeftMouseDown += (s, e) => { Show(); WindowState = WindowState.Normal; };

            var contextMenu = new ContextMenu();
            var hideShowMenuItem = new MenuItem { Header = "Hide / Show" };
            hideShowMenuItem.Click += (s, e) =>
            {
                if (IsVisible) Hide();
                else { Show(); WindowState = WindowState.Normal; }
            };
            contextMenu.Items.Add(hideShowMenuItem);

            var exitMenuItem = new MenuItem { Header = "Exit" };
            exitMenuItem.Click += (s, e) => Dispatcher.Invoke(RequestExit);
            contextMenu.Items.Add(exitMenuItem);
            _notifyIcon.ContextMenu = contextMenu;

            StateChanged += (s, e) => { if (WindowState == WindowState.Minimized) Hide(); };
        }

        private void RequestExit()
        {
            bool wasHidden = !this.IsVisible;
            if (wasHidden) Show();
            this.Activate();

            if (ConfirmOnExit)
            {
                var result = System.Windows.MessageBox.Show("Are you sure you want to exit?", "Confirm Exit", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                {
                    if (wasHidden) Hide();
                    return;
                }
            }

            _isExiting = true;
            SetStatus("Exiting...", clearAfter: false);
            SaveSettings();
            System.Windows.Application.Current.Shutdown();
        }

        #region Settings and Status
        private async void SetStatus(string message, bool clearAfter = true)
        {
            StatusText = message;

            if (clearAfter)
            {
                _statusClearTokenSource.Cancel();
                _statusClearTokenSource = new CancellationTokenSource();
                var token = _statusClearTokenSource.Token;

                await Task.Delay(5000);

                if (!token.IsCancellationRequested)
                {
                    StatusText = "Ready";
                }
            }
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    string json = File.ReadAllText(_settingsFilePath);
                    _settings = JsonConvert.DeserializeObject<AppSettings>(json) ?? new AppSettings();
                }
                else
                {
                    _settings = new AppSettings();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading settings: {ex.Message}");
                _settings = new AppSettings();
            }

            ConfirmOnExit = _settings.ConfirmOnExit;
            MinimizeToTrayOnClose = _settings.MinimizeToTrayOnClose;
            SelectedUpdateSource = _settings.SelectedUpdateSource;
            ThrottleLimit = _settings.ThrottleLimit;
            AsJob = _settings.AsJob;

            this.Top = _settings.WindowTop;
            this.Left = _settings.WindowLeft;
            this.Height = _settings.WindowHeight;
            this.Width = _settings.WindowWidth;
            this.WindowState = _settings.WindowState;

            ValidatePosition();

            if (_settings.SelectedTabIndex >= 0 && _settings.SelectedTabIndex < MainTabControl.Items.Count)
            {
                MainTabControl.SelectedIndex = _settings.SelectedTabIndex;
            }
        }

        private void SaveSettings()
        {
            try
            {
                _settings.ConfirmOnExit = this.ConfirmOnExit;
                _settings.MinimizeToTrayOnClose = this.MinimizeToTrayOnClose;
                _settings.SelectedUpdateSource = this.SelectedUpdateSource;
                _settings.ThrottleLimit = this.ThrottleLimit;
                _settings.AsJob = this.AsJob;
                _settings.SelectedTabIndex = MainTabControl.SelectedIndex;

                _settings.WindowState = this.WindowState == WindowState.Minimized ? WindowState.Normal : this.WindowState;
                _settings.WindowTop = this.RestoreBounds.Top;
                _settings.WindowLeft = this.RestoreBounds.Left;
                _settings.WindowHeight = this.RestoreBounds.Height;
                _settings.WindowWidth = this.RestoreBounds.Width;

                string json = JsonConvert.SerializeObject(_settings, Formatting.Indented);
                File.WriteAllText(_settingsFilePath, json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving settings: {ex.Message}");
            }
        }

        private void ValidatePosition()
        {
            bool isVisible = false;
            foreach (var screen in System.Windows.Forms.Screen.AllScreens)
            {
                var windowRect = new System.Drawing.Rectangle((int)this.Left, (int)this.Top, (int)this.Width, (int)this.Height);
                if (screen.WorkingArea.IntersectsWith(windowRect))
                {
                    isVisible = true;
                    break;
                }
            }

            if (!isVisible)
            {
                this.Left = 100;
                this.Top = 100;
            }
        }
        #endregion

        #region PowerShell Logic
        private async void RunUpdate_Click(object sender, RoutedEventArgs e)
        {
            SetStatus("Updating signatures...", clearAfter: false);
            UpdateOutputText = "Starting update...";
            var command = new StringBuilder("Update-MpSignature");

            if (!string.IsNullOrEmpty(SelectedUpdateSource)) { command.Append($" -UpdateSource {SelectedUpdateSource}"); }
            if (ThrottleLimit > 0 && ThrottleLimit <= 100) { command.Append($" -ThrottleLimit {ThrottleLimit}"); }
            if (AsJob) { command.Append(" -AsJob"); }

            UpdateOutputText = $"Running command: {command.ToString()}\n\n";

            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddScript(command.ToString());
                    var psOutput = await Task.Factory.FromAsync(ps.BeginInvoke(), ps.EndInvoke);
                    StringBuilder sb = new StringBuilder();
                    if (ps.Streams.Error.Count > 0)
                    {
                        foreach (var error in ps.Streams.Error) { sb.AppendLine($"ERROR: {error.ToString()}"); }
                        SetStatus("Update failed.");
                    }
                    else
                    {
                        if (psOutput != null) { foreach (var item in psOutput) { if (item != null) { sb.AppendLine(item.ToString()); } } }
                        if (sb.Length == 0) { sb.AppendLine("Command completed with no output."); }
                        SetStatus("Update complete.");
                    }

                    await Dispatcher.InvokeAsync(() => { UpdateOutputText += sb.ToString(); });
                }
            }
            catch (Exception ex)
            {
                SetStatus("Update failed with an exception.");
                await Dispatcher.InvokeAsync(() => { UpdateOutputText += $"\nEXCEPTION: {ex.Message}"; });
            }

            await InitializeDataAsync();
        }

        private async Task InitializeDataAsync()
        {
            SetStatus("Refreshing status...", clearAfter: false);
            try
            {
                ComputerStatusGrid.RowDefinitions.Clear();
                ComputerStatusGrid.Children.Clear();
                await GetComputerStatus();
                await GetAntiVirusStatus();
                GetServicesStatus();
                SetStatus("Status refreshed.");
            }
            catch (Exception ex)
            {
                SetStatus("Failed to refresh status.");
                Debug.WriteLine($"Error during InitializeDataAsync: {ex.Message}");
            }
        }

        private async Task GetComputerStatus()
        {
            Debug.WriteLine("GetComputerStatus() started.");
            try
            {
                using (PowerShell PowerShellInstance = PowerShell.Create())
                {
                    PowerShellInstance.AddScript("Set-ExecutionPolicy RemoteSigned; Import-Module ConfigDefender; Get-MpComputerStatus | Select-Object AntivirusSignatureLastUpdated, AntivirusSignatureVersion, AMEngineVersion, AMProductVersion, AMRunningMode, AMServiceEnabled, AMServiceVersion, AntispywareSignatureVersion, AntispywareEnabled, AntispywareSignatureLastUpdated, RealTimeProtectionEnabled, RebootRequired");
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

        private async Task GetAntiVirusStatus()
        {
            try
            {
                using (PowerShell PowerShellInstance = PowerShell.Create())
                {
                    PowerShellInstance.AddScript("Set-ExecutionPolicy RemoteSigned; Import-Module ConfigDefender; Get-MpComputerStatus | Select-Object AntivirusSignatureLastUpdated, AntivirusSignatureVersion");
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
                        ComputerStatusGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                        var headerLabel = new TextBlock { Text = "Antivirus status:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 2) };
                        Grid.SetRow(headerLabel, ComputerStatusGrid.RowDefinitions.Count - 1);
                        Grid.SetColumn(headerLabel, 0);
                        Grid.SetColumnSpan(headerLabel, 2);
                        ComputerStatusGrid.Children.Add(headerLabel);
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

        private void GetServicesStatus()
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                string[] services = { "MDCoreSvc", "mpssvc", "Sense", "WdNisSvc", "WinDefend" };
                foreach (string serviceName in services)
                {
                    string displayName;
                    string status;
                    (displayName, status) = GetServiceStatus(serviceName);
                    sb.AppendLine($"Service ‘{displayName}’ status: {status}");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine(ex.Message);
            }

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
                        string serviceName = parts[0].Trim();
                        string statusText = parts[1].Trim();

                        ComputerStatusGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                        var label = new TextBlock { Text = serviceName + ":", FontWeight = FontWeights.Normal, Margin = new Thickness(10, 2, 5, 2) };
                        Grid.SetRow(label, ComputerStatusGrid.RowDefinitions.Count - 1);
                        Grid.SetColumn(label, 0);
                        ComputerStatusGrid.Children.Add(label);

                        var value = new TextBlock { Text = statusText, Margin = new Thickness(0, 2, 0, 2) };

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
        #endregion

        #region Window Events
        protected override void OnClosing(CancelEventArgs e)
        {
            // If shutdown is already in progress, do nothing more.
            if (_isExiting) return;

            if (MinimizeToTrayOnClose)
            {
                SaveSettings();
                e.Cancel = true;
                Hide();
            }
            else
            {
                // Cancel this event and let RequestExit handle the shutdown process,
                // including the confirmation prompt.
                e.Cancel = true;
                RequestExit();
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            _notifyIcon.Dispose();
            base.OnClosed(e);
        }
        #endregion

        #region Helpers
        static protected (string, string) GetServiceStatus(String serviceName)
        {
            try
            {
                ServiceController sc = new ServiceController(serviceName);
                return (sc.DisplayName.ToString(), sc.Status.ToString());
            }
            catch (Exception)
            {
                return ("Not found", "Unknown");
            }
        }
        #endregion
    }
}
