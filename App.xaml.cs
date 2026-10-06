using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace OutlookEventForwarder;

public partial class App : Application
{
    public static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "EventForwarder_error.log");

    public static void LogError(string context, Exception ex)
    {
        var entry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}\n{ex.GetType().FullName}: {ex.Message}\n{ex.StackTrace}\nInner: {ex.InnerException}\n\n";
        File.AppendAllText(LogPath, entry);
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        LogError("DispatcherUnhandledException", e.Exception);
        MessageBox.Show($"Error (logged to Desktop):\n\n{e.Exception.Message}",
            "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true;
    }

    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            LogError("UnhandledException", ex);
            MessageBox.Show($"Fatal error (logged to Desktop):\n\n{ex.Message}",
                "Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
