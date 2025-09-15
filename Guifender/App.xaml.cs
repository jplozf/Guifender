using System.Configuration;
using System.Data;
using System.Windows;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Linq;
using System.Threading;

namespace Guifender;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : System.Windows.Application
{
    private const string AppName = "GuifenderSingleton";
    private static Mutex _mutex = null;

    protected override async void OnStartup(StartupEventArgs e)
    {
        _mutex = new Mutex(true, AppName, out bool createdNew);

        if (!createdNew)
        {
            // App is already running.
            System.Windows.Application.Current.Shutdown();
            return;
        }

        base.OnStartup(e);
        await CheckForNewVersionAsync();
    }

    private async Task CheckForNewVersionAsync()
    {
        try
        {
            var versionAttribute = Assembly.GetExecutingAssembly()
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>();
            if (versionAttribute == null) return;

            var versionInfo = versionAttribute.InformationalVersion;
            var localCommit = versionInfo.Split('+').Last();

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("Guifender", "1.0"));
                var response = await client.GetStringAsync("https://api.github.com/repos/jplozf/Guifender/branches/main");

                var githubBranch = JsonConvert.DeserializeObject<GithubBranch>(response);
                var remoteCommit = githubBranch?.Commit?.Sha;

                if (remoteCommit != null && !remoteCommit.StartsWith(localCommit))
                {
                    System.Windows.MessageBox.Show("A new version of Guifender is available on GitHub.",
                                    "Update Available",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Information);
                }
            }
        }
        catch
        {
            // Silently fail if the check doesn't succeed.
        }
    }

    public class GithubBranch
    {
        [JsonProperty("commit")]
        public GithubCommit Commit { get; set; }
    }

    public class GithubCommit
    {
        [JsonProperty("sha")]
        public string Sha { get; set; }
    }
}

