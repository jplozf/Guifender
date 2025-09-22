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

        protected override void OnStartup(StartupEventArgs e)
        {
            _mutex = new Mutex(true, AppName, out bool createdNew);
    
            if (!createdNew)
            {
                // App is already running.
                System.Windows.Application.Current.Shutdown();
                return;
            }
    
            base.OnStartup(e);
        }}

