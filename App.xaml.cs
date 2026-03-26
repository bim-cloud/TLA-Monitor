using System.Windows;
using AutodeskIDMonitor.Services;
using AutodeskIDMonitor.Views;

namespace AutodeskIDMonitor;

public partial class App : Application
{
    /// <summary>
    /// Indicates if the app should start minimized to system tray
    /// </summary>
    public static bool StartMinimized { get; private set; } = false;
    
    /// <summary>
    /// Current user's display name (from first-run setup or Windows)
    /// </summary>
    public static string CurrentUserDisplayName { get; private set; } = "";
    
    /// <summary>
    /// Current user's email
    /// </summary>
    public static string CurrentUserEmail { get; private set; } = "";
    
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        
        // Check for --minimized argument (used when starting with Windows)
        if (e.Args.Length > 0)
        {
            foreach (var arg in e.Args)
            {
                if (arg.Equals("--minimized", StringComparison.OrdinalIgnoreCase) ||
                    arg.Equals("-m", StringComparison.OrdinalIgnoreCase) ||
                    arg.Equals("/minimized", StringComparison.OrdinalIgnoreCase))
                {
                    StartMinimized = true;
                    break;
                }
            }
        }
        
        // Initialize services
        AdminService.Instance.Initialize();
        ConfigService.Instance.Load();
        
        // Check if first-time setup is needed
        var localStorage = new LocalStorageService();
        
        if (localStorage.NeedsFirstTimeSetup())
        {
            // Get default suggestions
            var windowsUser = Environment.UserName;
            if (windowsUser.Equals("User", StringComparison.OrdinalIgnoreCase))
            {
                windowsUser = ""; // Don't suggest generic "User"
            }
            
            // Show first-run sign-in window
            var signInWindow = new FirstTimeSetupWindow(windowsUser, "");
            var result = signInWindow.ShowDialog();
            
            if (result == true && signInWindow.SetupCompleted)
            {
                // Save user settings
                localStorage.SaveUserProfileSettings(signInWindow.EnteredName, signInWindow.EnteredEmail);
                CurrentUserDisplayName = signInWindow.EnteredName;
                CurrentUserEmail = signInWindow.EnteredEmail;
            }
            else
            {
                // User cancelled - use Windows username as fallback
                if (string.IsNullOrEmpty(windowsUser) || windowsUser.Equals("User", StringComparison.OrdinalIgnoreCase))
                {
                    windowsUser = "Team Member"; // Better default
                }
                CurrentUserDisplayName = windowsUser;
                // Still mark as needing setup later
            }
        }
        else
        {
            // Load existing user settings
            var settings = localStorage.GetUserProfileSettings();
            CurrentUserDisplayName = settings.DisplayName;
            CurrentUserEmail = settings.Email;
        }
    }
}
