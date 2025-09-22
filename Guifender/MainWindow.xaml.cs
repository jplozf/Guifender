using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Management.Automation;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Hardcodet.Wpf.TaskbarNotification;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Guifender
{
    public class AppSettings
    {
        public bool ConfirmOnExit { get; set; } = true;
        public bool MinimizeToTrayOnClose { get; set; } = true;
        public string SelectedUpdateSource { get; set; } = "MicrosoftUpdateServer";
        public int ThrottleLimit { get; set; } = 0;
        public bool AsJob { get; set; } = false;
        public bool IsScheduledUpdateEnabled { get; set; } = false;
        public int ScheduledUpdateIntervalHours { get; set; } = 24;
        public DateTime? LastUpdateTime { get; set; } = null;
        public int StatusClearDelaySeconds { get; set; } = 5;

        public double WindowTop { get; set; } = 100;
        public double WindowLeft { get; set; } = 100;
        public double WindowHeight { get; set; } = 600;
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
        private DateTime? _lastUpdateTime;
        private System.Threading.Timer _updateSchedulerTimer;

        #region Properties
        private string _statusText; public string StatusText { get => _statusText; set { _statusText = value; OnPropertyChanged(); } }
        private bool _confirmOnExit; public bool ConfirmOnExit { get => _confirmOnExit; set { _confirmOnExit = value; OnPropertyChanged(); } }
        private bool _minimizeToTrayOnClose; public bool MinimizeToTrayOnClose { get => _minimizeToTrayOnClose; set { _minimizeToTrayOnClose = value; OnPropertyChanged(); } }
        private List<string> _updateSources; public List<string> UpdateSources { get => _updateSources; set { _updateSources = value; OnPropertyChanged(); } }
        private string _selectedUpdateSource; public string SelectedUpdateSource { get => _selectedUpdateSource; set { _selectedUpdateSource = value; OnPropertyChanged(); } }
        private int _throttleLimit; public int ThrottleLimit { get => _throttleLimit; set { _throttleLimit = value; OnPropertyChanged(); } }
        private bool _asJob; public bool AsJob { get => _asJob; set { _asJob = value; OnPropertyChanged(); } }
        private string _updateOutputText; public string UpdateOutputText { get => _updateOutputText; set { _updateOutputText = value; OnPropertyChanged(); } }
        private string _antivirusSignatureVersion; public string AntivirusSignatureVersion { get => _antivirusSignatureVersion; set { _antivirusSignatureVersion = value; OnPropertyChanged(); } }
        private string _antivirusSignatureLastUpdated; public string AntivirusSignatureLastUpdated { get => _antivirusSignatureLastUpdated; set { _antivirusSignatureLastUpdated = value; OnPropertyChanged(); } }
        private string _lastUpdateCheckString; public string LastUpdateCheckString { get => _lastUpdateCheckString; set { _lastUpdateCheckString = value; OnPropertyChanged(); } }

        private bool _isScheduledUpdateEnabled;
        public bool IsScheduledUpdateEnabled
        {
            get => _isScheduledUpdateEnabled;
            set { _isScheduledUpdateEnabled = value; OnPropertyChanged(); SetupUpdateScheduler(); }
        }

        private int _scheduledUpdateIntervalHours;
        public int ScheduledUpdateIntervalHours
        {
            get => _scheduledUpdateIntervalHours;
            set { _scheduledUpdateIntervalHours = value; OnPropertyChanged(); SetupUpdateScheduler(); }
        }

        private string _nextScheduledUpdateString;
        public string NextScheduledUpdateString
        {
            get => _nextScheduledUpdateString;
            set { _nextScheduledUpdateString = value; OnPropertyChanged(); }
        }

        private int _statusClearDelaySeconds; 
        public int StatusClearDelaySeconds
        {
            get => _statusClearDelaySeconds;
            set { _statusClearDelaySeconds = value; OnPropertyChanged(); }
        }
        #endregion

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        public MainWindow()
        {
            InitializeComponent();
            SetStatus("Guifender starting...", clearAfter: false);
            UpdateSources = new List<string> { "MicrosoftUpdateServer", "InternalDefinitionUpdateServer", "MMPC" };
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = Path.Combine(appDataPath, "Guifender");
            Directory.CreateDirectory(settingsDir);
            _settingsFilePath = Path.Combine(settingsDir, "settings.json");
            LoadSettings();
            this.Title = GetVersionInfo();
            SetupTaskbarIcon();
            SetStatus("Ready", clearAfter: false);
            _ = InitializeDataAsync();
            SetupUpdateScheduler();
        }

        #region Core App Logic
        private string GetVersionInfo()
        {
            try
            {
                string commitCount = RunGitCommand("rev-list --count HEAD").Trim();
                string commitHash = RunGitCommand("rev-parse --short HEAD").Trim();

                if (!string.IsNullOrEmpty(commitCount) && !string.IsNullOrEmpty(commitHash))
                {
                    return $"Guifender {MajorVersion}.{commitCount}-{commitHash}";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Could not get version info from Git: {ex.Message}");
            }
            return "Guifender";
        }

        private string RunGitCommand(string arguments)
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "git.exe",
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                }
            };
            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            return output;
        }

        private void SetupTaskbarIcon()
        {
            var resourceInfo = System.Windows.Application.GetResourceStream(new Uri("pack://application:,,,/app_icon.png"));
            if (resourceInfo == null || resourceInfo.Stream == null) { System.Windows.MessageBox.Show("Resource 'app_icon.png' not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error); return; }
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
            hideShowMenuItem.Click += (s, e) => { if (IsVisible) Hide(); else { Show(); WindowState = WindowState.Normal; } };
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
                if (result == MessageBoxResult.No) { if (wasHidden) Hide(); return; }
            }
            _isExiting = true;
            SetStatus("Exiting...", clearAfter: false);
            Logger.Write("Guifender exiting.");
            SaveSettings();
            System.Windows.Application.Current.Shutdown();
        }

        private async void SetStatus(string message, bool clearAfter = true)
        {
            StatusText = message;
            Logger.Write(message);

            if (clearAfter && StatusClearDelaySeconds > 0)
            {
                _statusClearTokenSource.Cancel();
                _statusClearTokenSource = new CancellationTokenSource();
                var token = _statusClearTokenSource.Token;

                await Task.Delay(StatusClearDelaySeconds * 1000);

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
                if (File.Exists(_settingsFilePath)) { _settings = JsonConvert.DeserializeObject<AppSettings>(File.ReadAllText(_settingsFilePath)) ?? new AppSettings(); }
                else { _settings = new AppSettings(); }
            }
            catch (Exception ex) { Debug.WriteLine($"Error loading settings: {ex.Message}"); _settings = new AppSettings(); }

            ConfirmOnExit = _settings.ConfirmOnExit;
            MinimizeToTrayOnClose = _settings.MinimizeToTrayOnClose;
            SelectedUpdateSource = _settings.SelectedUpdateSource;
            ThrottleLimit = _settings.ThrottleLimit;
            AsJob = _settings.AsJob;
            IsScheduledUpdateEnabled = _settings.IsScheduledUpdateEnabled;
            ScheduledUpdateIntervalHours = _settings.ScheduledUpdateIntervalHours;
            _lastUpdateTime = _settings.LastUpdateTime;
            LastUpdateCheckString = _lastUpdateTime.HasValue ? _lastUpdateTime.Value.ToString("g") : "Never";
            StatusClearDelaySeconds = _settings.StatusClearDelaySeconds;

            this.Top = _settings.WindowTop; this.Left = _settings.WindowLeft; this.Height = _settings.WindowHeight; this.Width = _settings.WindowWidth; this.WindowState = _settings.WindowState;
            ValidatePosition();
            if (_settings.SelectedTabIndex >= 0 && _settings.SelectedTabIndex < MainTabControl.Items.Count) { MainTabControl.SelectedIndex = _settings.SelectedTabIndex; }
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
                _settings.IsScheduledUpdateEnabled = this.IsScheduledUpdateEnabled;
                _settings.ScheduledUpdateIntervalHours = this.ScheduledUpdateIntervalHours;
                _settings.LastUpdateTime = this._lastUpdateTime;
                _settings.StatusClearDelaySeconds = this.StatusClearDelaySeconds;
                _settings.SelectedTabIndex = MainTabControl.SelectedIndex;
                _settings.WindowState = this.WindowState == WindowState.Minimized ? WindowState.Normal : this.WindowState;
                _settings.WindowTop = this.RestoreBounds.Top; _settings.WindowLeft = this.RestoreBounds.Left; _settings.WindowHeight = this.RestoreBounds.Height; _settings.WindowWidth = this.RestoreBounds.Width;
                File.WriteAllText(_settingsFilePath, JsonConvert.SerializeObject(_settings, Formatting.Indented));
            }
            catch (Exception ex) { Debug.WriteLine($"Error saving settings: {ex.Message}"); }
        }

        private void ValidatePosition()
        {
            bool isVisible = false;
            foreach (var screen in System.Windows.Forms.Screen.AllScreens)
            {
                var windowRect = new System.Drawing.Rectangle((int)this.Left, (int)this.Top, (int)this.Width, (int)this.Height);
                if (screen.WorkingArea.IntersectsWith(windowRect)) { isVisible = true; break; }
            }
            if (!isVisible) { this.Left = 100; this.Top = 100; }
        }
        #endregion

        #region Update and Scheduling Logic
        private void SetupUpdateScheduler()
        {
            _updateSchedulerTimer?.Dispose();
            if (!IsScheduledUpdateEnabled || ScheduledUpdateIntervalHours <= 0) { NextScheduledUpdateString = "Disabled"; return; }
            var interval = TimeSpan.FromHours(ScheduledUpdateIntervalHours);
            DateTime lastRun = _lastUpdateTime ?? DateTime.MinValue;
            DateTime nextRun = lastRun.Add(interval);
            TimeSpan dueTime = (nextRun < DateTime.Now) ? TimeSpan.Zero : nextRun - DateTime.Now;
            NextScheduledUpdateString = DateTime.Now.Add(dueTime).ToString("g");
            _updateSchedulerTimer = new System.Threading.Timer(ScheduledUpdateTimer_Callback, null, dueTime, interval);
        }

        private async void ScheduledUpdateTimer_Callback(object state)
        {
            await Dispatcher.InvokeAsync(() => SetStatus("Running scheduled update...", clearAfter: false));
            bool success = await PerformUpdateAsync(isScheduled: true);
            await Dispatcher.InvokeAsync(() =>
            {
                SetStatus(success ? "Scheduled update complete." : "Scheduled update failed.");
                SetupUpdateScheduler();
            });
        }

        private async void RunUpdate_Click(object sender, RoutedEventArgs e)
        {
            await PerformUpdateAsync(isScheduled: false);
            await InitializeDataAsync();
        }

        private async Task<bool> PerformUpdateAsync(bool isScheduled)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                if (!isScheduled) { UpdateOutputText = ""; }
                SetStatus("Updating signatures...", clearAfter: false);
            });

            var command = new StringBuilder("Update-MpSignature");
            if (!isScheduled)
            {
                if (!string.IsNullOrEmpty(SelectedUpdateSource)) { command.Append($" -UpdateSource {SelectedUpdateSource}"); }
                if (ThrottleLimit > 0 && ThrottleLimit <= 100) { command.Append($" -ThrottleLimit {ThrottleLimit}"); }
                if (AsJob) { command.Append(" -AsJob"); }
            }

            string output = $"Update started at {DateTime.Now:g}...\nRunning command: {command.ToString()}\n\n";
            bool success = false;

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
                    }
                    else
                    {
                        if (psOutput != null) { foreach (var item in psOutput) { if (item != null) { sb.AppendLine(item.ToString()); } } }
                        if (sb.Length == 0) { sb.AppendLine("Command completed with no output."); }
                        _lastUpdateTime = DateTime.Now;
                        success = true;
                    }
                    output += sb.ToString();
                }
            }
            catch (Exception ex) { output += $"\nEXCEPTION: {ex.Message}"; }

            await Dispatcher.InvokeAsync(() =>
            {
                UpdateOutputText += output;
                LastUpdateCheckString = _lastUpdateTime.HasValue ? _lastUpdateTime.Value.ToString("g") : "Never";
                if (success && !isScheduled) { SetStatus("Update complete."); }
                else if (!success) { SetStatus("Update failed."); }
            });

            return success;
        }
        #endregion

        #region Status Panel Logic
        private async Task InitializeDataAsync()
        {
            SetStatus("Refreshing status...", clearAfter: false);
            try
            {
                await Dispatcher.InvokeAsync(() => StatusPanel.Children.Clear());
                await GetComputerStatus();
                GetServicesStatus();
                await CheckForNewVersionAsync();
                SetStatus("Status refreshed.");
            }
            catch (Exception ex) { SetStatus("Failed to refresh status."); Debug.WriteLine($"Error during InitializeDataAsync: {ex.Message}"); }
        }

        private async Task CheckForNewVersionAsync()
        {
            SetStatus("Checking for new version...", clearAfter: false);
            try
            {
                string localCommitHash = RunGitCommand("rev-parse HEAD").Trim();

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "Guifender-App");
                    string url = "https://api.github.com/repos/jplozf/Guifender/commits/main";
                    string json = await client.GetStringAsync(url);
                    JObject latestCommit = JObject.Parse(json);
                    string remoteCommitHash = latestCommit["sha"].ToString();

                    if (!string.IsNullOrEmpty(localCommitHash) && !localCommitHash.Equals(remoteCommitHash, StringComparison.OrdinalIgnoreCase))
                    {
                        var hyperlink = new Hyperlink(new Run("A new version is available! Click here to download."));
                        hyperlink.NavigateUri = new Uri("https://github.com/jplozf/Guifender/releases");
                        hyperlink.RequestNavigate += Hyperlink_RequestNavigate;
                        VersionCheckTextBlock.Inlines.Clear();
                        VersionCheckTextBlock.Inlines.Add(hyperlink);
                        SetStatus("New version available!");
                    }
                    else
                    {
                        VersionCheckTextBlock.Text = "You are using the latest version.";
                        SetStatus("Ready");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Version check failed: {ex.Message}");
                VersionCheckTextBlock.Text = "Version check failed.";
                SetStatus("Version check failed.");
            }
        }

        private async Task GetComputerStatus()
        {
            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddScript("Get-MpComputerStatus");
                    var psOutput = await Task.Factory.FromAsync(ps.BeginInvoke(), ps.EndInvoke);
                    if (ps.Streams.Error.Count > 0) { /* Handle errors */ return; }

                    if (psOutput != null && psOutput.Count > 0)
                    {
                        var status = psOutput[0];
                        await Dispatcher.InvokeAsync(() =>
                        {
                            AntivirusSignatureLastUpdated = Convert.ToDateTime(status.Properties["AntivirusSignatureLastUpdated"].Value).ToString("g");
                            AntivirusSignatureVersion = status.Properties["AntivirusSignatureVersion"].Value.ToString();

                            var generalStatus = new Dictionary<string, object> { { "RealTimeProtectionEnabled", status.Properties["RealTimeProtectionEnabled"].Value }, { "RebootRequired", status.Properties["RebootRequired"].Value }, { "AMServiceEnabled", status.Properties["AMServiceEnabled"].Value }, { "AMRunningMode", status.Properties["AMRunningMode"].Value } };
                            StatusPanel.Children.Add(CreateStatusGroupBox("General Status", generalStatus));

                            var antivirusStatus = new Dictionary<string, object> { { "AntivirusEnabled", status.Properties["AntivirusEnabled"].Value }, { "AntivirusSignatureVersion", status.Properties["AntivirusSignatureVersion"].Value }, { "AntivirusSignatureLastUpdated", status.Properties["AntivirusSignatureLastUpdated"].Value } };
                            StatusPanel.Children.Add(CreateStatusGroupBox("Antivirus Status", antivirusStatus));

                            var antispywareStatus = new Dictionary<string, object> { { "AntispywareEnabled", status.Properties["AntispywareEnabled"].Value }, { "AntispywareSignatureVersion", status.Properties["AntispywareSignatureVersion"].Value }, { "AntispywareSignatureLastUpdated", status.Properties["AntispywareSignatureLastUpdated"].Value } };
                            StatusPanel.Children.Add(CreateStatusGroupBox("Antispyware Status", antispywareStatus));

                            var productInfo = new Dictionary<string, object> { { "AMEngineVersion", status.Properties["AMEngineVersion"].Value }, { "AMProductVersion", status.Properties["AMProductVersion"].Value }, { "AMServiceVersion", status.Properties["AMServiceVersion"].Value } };
                            StatusPanel.Children.Add(CreateStatusGroupBox("Product Information", productInfo));
                        });
                    }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"Exception in GetComputerStatus(): {ex.Message}"); }
        }

        private void GetServicesStatus()
        {
            var services = new Dictionary<string, object>();
            string[] serviceNames = { "MDCoreSvc", "mpssvc", "Sense", "WdNisSvc", "WinDefend" };
            foreach (string serviceName in serviceNames)
            {
                var (displayName, status) = GetServiceStatus(serviceName);
                services.Add(displayName, status);
            }
            Dispatcher.Invoke(() => StatusPanel.Children.Add(CreateStatusGroupBox("Defender Services", services)));
        }

        private System.Windows.Controls.GroupBox CreateStatusGroupBox(string header, Dictionary<string, object> properties)
        {
            var groupBox = new System.Windows.Controls.GroupBox { Header = header, Margin = new Thickness(0, 0, 0, 5) };
            var grid = new Grid { Margin = new Thickness(5) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            int row = 0;
            foreach (var prop in properties)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var label = new TextBlock { Text = prop.Key + ":", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 2, 5, 2) };
                Grid.SetRow(label, row); Grid.SetColumn(label, 0); grid.Children.Add(label);

                var value = new TextBlock { Text = prop.Value?.ToString() ?? "N/A", Margin = new Thickness(0, 2, 0, 2) };
                if (prop.Value is bool) { value.Foreground = (bool)prop.Value ? System.Windows.Media.Brushes.Green : System.Windows.Media.Brushes.Red; }
                else if (prop.Value is string && ((string)prop.Value).Equals("Running", StringComparison.OrdinalIgnoreCase)) { value.Foreground = System.Windows.Media.Brushes.Green; }
                else if (prop.Value is string && ((string)prop.Value).Equals("Stopped", StringComparison.OrdinalIgnoreCase)) { value.Foreground = System.Windows.Media.Brushes.Red; }

                Grid.SetRow(value, row); Grid.SetColumn(value, 1); grid.Children.Add(value);
                row++;
            }
            groupBox.Content = grid;
            return groupBox;
        }
        #endregion

        #region Window Events
        protected override void OnClosing(CancelEventArgs e)
        {
            if (_isExiting) return;
            if (MinimizeToTrayOnClose) { SaveSettings(); e.Cancel = true; Hide(); }
            else { e.Cancel = true; RequestExit(); }
        }

        protected override void OnClosed(EventArgs e) { _notifyIcon.Dispose(); base.OnClosed(e); }
        #endregion

        #region Helpers
        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }

        static protected (string, string) GetServiceStatus(String serviceName)
        {
            try { ServiceController sc = new ServiceController(serviceName); return (sc.DisplayName.ToString(), sc.Status.ToString()); } 
            catch (Exception) { return ("Not found", "Unknown"); }
        }
        #endregion
    }
}