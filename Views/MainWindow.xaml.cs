using System.Windows;
using System.Windows.Forms;
using System.Drawing;
using System.Windows.Media.Animation;
using System.ComponentModel;
using AutodeskIDMonitor.ViewModels;
using AutodeskIDMonitor.Models;
using AutodeskIDMonitor.Controls;

namespace AutodeskIDMonitor.Views;

public partial class MainWindow : Window
{
    private NotifyIcon? _notifyIcon;
    private bool _isActuallyClosing = false;
    private bool _hasShownTrayBalloon = false;
    private string _lastKnownProject = "";
    
    public MainWindow()
    {
        InitializeComponent();
        
        // Initialize system tray icon
        InitializeSystemTray();
        
        // Wire up chart update events
        Loaded += MainWindow_Loaded;
        
        // Handle window closing to minimize to tray instead
        Closing += MainWindow_Closing;
        
        // Handle state changes
        StateChanged += MainWindow_StateChanged;
    }
    
    private void InitializeSystemTray()
    {
        try
        {
            // Create context menu for tray icon
            var contextMenu = new ContextMenuStrip();
            
            var openItem = new ToolStripMenuItem("Open ID Monitor");
            openItem.Click += (s, e) => ShowMainWindow();
            openItem.Font = new Font(openItem.Font, System.Drawing.FontStyle.Bold);
            contextMenu.Items.Add(openItem);
            
            contextMenu.Items.Add(new ToolStripSeparator());
            
            var dashboardItem = new ToolStripMenuItem("Open Dashboard");
            dashboardItem.Click += (s, e) => 
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "http://141.145.153.32:5000/",
                        UseShellExecute = true
                    });
                }
                catch { }
            };
            contextMenu.Items.Add(dashboardItem);
            
            contextMenu.Items.Add(new ToolStripSeparator());
            
            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (s, e) => ExitApplication();
            contextMenu.Items.Add(exitItem);
            
            // Create tray icon
            _notifyIcon = new NotifyIcon
            {
                Text = "Tangent ID Monitor - Running",
                Visible = true,
                ContextMenuStrip = contextMenu
            };
            
            // Try to load icon from file, fallback to default
            try
            {
                var iconPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tray.ico");
                if (System.IO.File.Exists(iconPath))
                {
                    _notifyIcon.Icon = new Icon(iconPath);
                }
                else
                {
                    // Try app.ico
                    iconPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.ico");
                    if (System.IO.File.Exists(iconPath))
                    {
                        _notifyIcon.Icon = new Icon(iconPath);
                    }
                    else
                    {
                        // Use default application icon
                        _notifyIcon.Icon = SystemIcons.Application;
                    }
                }
            }
            catch
            {
                _notifyIcon.Icon = SystemIcons.Application;
            }
            
            // Double-click tray icon to show window
            _notifyIcon.DoubleClick += (s, e) => ShowMainWindow();
            
            // Show balloon tip when minimized to tray (first time only)
            _notifyIcon.BalloonTipTitle = "Tangent ID Monitor";
            _notifyIcon.BalloonTipText = "Application is still running in the background. Double-click to open.";
            _notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to initialize system tray: {ex.Message}");
        }
    }
    
    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        // If not actually closing (user clicked X), minimize to tray instead
        if (!_isActuallyClosing)
        {
            e.Cancel = true;
            MinimizeToTray();
        }
        else
        {
            // Clean up tray icon when actually closing
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                _notifyIcon = null;
            }
        }
    }
    
    private void MainWindow_StateChanged(object? sender, EventArgs e)
    {
        // When minimized, hide window and show in tray only
        if (WindowState == WindowState.Minimized)
        {
            MinimizeToTray();
        }
    }
    
    private void MinimizeToTray()
    {
        Hide();
        WindowState = WindowState.Minimized;
        
        // Update tray icon text
        if (_notifyIcon != null)
        {
            _notifyIcon.Text = "Tangent ID Monitor - Running (Click to open)";
            
            // Show balloon notification (only show once per session)
            if (!_hasShownTrayBalloon)
            {
                _notifyIcon.ShowBalloonTip(3000);
                _hasShownTrayBalloon = true;
            }
        }
    }
    
    private void ShowMainWindow()
    {
        Show();
        WindowState = WindowState.Normal;
        Activate();
        Focus();
        
        // Update tray icon text
        if (_notifyIcon != null)
        {
            _notifyIcon.Text = "Tangent ID Monitor - Running";
        }
    }
    
    private void ExitApplication()
    {
        var result = System.Windows.MessageBox.Show(
            "Are you sure you want to exit?\n\nThis will stop activity monitoring.",
            "Exit Tangent ID Monitor",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        
        if (result == MessageBoxResult.Yes)
        {
            _isActuallyClosing = true;
            
            // Clean up tray icon
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                _notifyIcon = null;
            }
            
            System.Windows.Application.Current.Shutdown();
        }
    }
    
    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel vm)
        {
            // Subscribe to property changes for project animation
            vm.PropertyChanged += OnViewModelPropertyChanged;
            
            // Subscribe to activity breakdown updates (for current user)
            vm.ActivityBreakdownUpdated += (s, breakdown) =>
            {
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        // Only update chart if no user is selected (show current user's data)
                        if (vm.SelectedUserActivity == null)
                        {
                            ActivityChart?.UpdateChart(breakdown);
                            
                            // Update pie chart with today's totals
                            if (breakdown != null)
                            {
                                ActivityPie?.UpdateChart(
                                    breakdown.RevitHours,
                                    breakdown.MeetingHours,
                                    breakdown.IdleHours
                                );
                            }
                        }
                    }
                    catch { }
                });
            };
            
            // Subscribe to user selection in activity table
            vm.UserActivitySelected += (s, userActivity) =>
            {
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        if (userActivity != null)
                        {
                            // Create a breakdown from the selected user's data
                            var userBreakdown = new DailyActivityBreakdown
                            {
                                Date = DateTime.Today,
                                RevitMinutes = userActivity.RevitHours * 60,
                                MeetingMinutes = userActivity.MeetingHours * 60,
                                TotalIdleMinutes = userActivity.IdleHours * 60,
                                HourlyBreakdown = new List<HourlyActivity>()
                            };
                            
                            // Build simple hourly breakdown based on total hours
                            // Distribute hours across work hours (8 AM - 6 PM)
                            double remainingRevit = userActivity.RevitHours * 60;
                            double remainingMeeting = userActivity.MeetingHours * 60;
                            double remainingIdle = userActivity.IdleHours * 60;
                            
                            for (int hour = 0; hour < 24; hour++)
                            {
                                var hourly = new HourlyActivity
                                {
                                    Hour = hour,
                                    HourLabel = $"{(hour == 0 ? 12 : hour > 12 ? hour - 12 : hour)} {(hour < 12 ? "AM" : "PM")}",
                                    IsWorkHour = hour >= 8 && hour < 18
                                };
                                
                                // Distribute time during work hours
                                if (hour >= 8 && hour < 18)
                                {
                                    double minutesThisHour = Math.Min(60, remainingRevit);
                                    hourly.RevitMinutes = minutesThisHour;
                                    remainingRevit -= minutesThisHour;
                                    
                                    if (remainingRevit <= 0)
                                    {
                                        minutesThisHour = Math.Min(60 - hourly.RevitMinutes, remainingMeeting);
                                        hourly.MeetingMinutes = minutesThisHour;
                                        remainingMeeting -= minutesThisHour;
                                    }
                                    
                                    double accountedMinutes = hourly.RevitMinutes + hourly.MeetingMinutes;
                                    if (accountedMinutes < 60 && remainingIdle > 0)
                                    {
                                        minutesThisHour = Math.Min(60 - accountedMinutes, remainingIdle);
                                        hourly.IdleMinutes = minutesThisHour;
                                        remainingIdle -= minutesThisHour;
                                    }
                                }
                                
                                // Check for overtime (after 6 PM)
                                if (hour >= 18 && userActivity.OvertimeHours > 0)
                                {
                                    hourly.IsOvertime = true;
                                    if (remainingRevit > 0)
                                    {
                                        double minutesThisHour = Math.Min(60, remainingRevit);
                                        hourly.RevitMinutes = minutesThisHour;
                                        remainingRevit -= minutesThisHour;
                                    }
                                }
                                
                                userBreakdown.HourlyBreakdown.Add(hourly);
                            }
                            
                            ActivityChart?.SetUserName(userActivity.WindowsDisplayName);
                            ActivityChart?.UpdateChart(userBreakdown);
                            ActivityPie?.UpdateChart(
                                userActivity.RevitHours,
                                userActivity.MeetingHours,
                                userActivity.IdleHours
                            );
                        }
                    }
                    catch { }
                });
            };
            
            // Subscribe to chart bar click events
            if (ActivityChart != null)
            {
                ActivityChart.HourClicked += (s, args) =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        try
                        {
                            var hourData = args.HourlyData;
                            if (hourData == null) return;
                            
                            var message = $"📊 Activity Details for {args.UserName}\n" +
                                $"━━━━━━━━━━━━━━━━━━━━━━━\n" +
                                $"🕐 Time: {args.HourLabel}\n\n" +
                                $"🔵 Revit: {hourData.RevitMinutes:F0} minutes\n" +
                                $"🟣 Meeting: {hourData.MeetingMinutes:F0} minutes\n" +
                                $"⚪ Idle: {hourData.IdleMinutes:F0} minutes\n" +
                                $"🟠 Other: {hourData.OtherMinutes:F0} minutes\n\n" +
                                $"📈 Total: {(hourData.RevitMinutes + hourData.MeetingMinutes + hourData.IdleMinutes + hourData.OtherMinutes):F0} minutes";
                            
                            if (hourData.IsOvertime)
                            {
                                message += "\n\n⚠️ OVERTIME HOUR";
                            }
                            
                            System.Windows.MessageBox.Show(message, $"Hourly Activity - {args.HourLabel}", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        catch { }
                    });
                };
            }
            
            // Subscribe to meeting status updates
            vm.MeetingStatusUpdated += (s, meetings) =>
            {
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        MeetingStatus?.UpdateStatus(meetings);
                    }
                    catch { }
                });
            };
            
            // Initial chart update - Use ACTUAL data from ViewModel
            var initialBreakdown = vm.TodayActivityBreakdown;
            if (initialBreakdown == null || initialBreakdown.HourlyBreakdown == null || initialBreakdown.HourlyBreakdown.Count == 0)
            {
                // Create default breakdown if none exists
                initialBreakdown = new DailyActivityBreakdown
                {
                    Date = DateTime.Today,
                    HourlyBreakdown = new List<HourlyActivity>()
                };
                
                // Build initial hourly breakdown
                for (int hour = 0; hour < 24; hour++)
                {
                    initialBreakdown.HourlyBreakdown.Add(new HourlyActivity
                    {
                        Hour = hour,
                        HourLabel = $"{(hour == 0 ? 12 : hour > 12 ? hour - 12 : hour)} {(hour < 12 ? "AM" : "PM")}",
                        IsWorkHour = hour >= 8 && hour < 18
                    });
                }
            }
            
            ActivityChart?.SetUserName(vm.LocalUserDisplayName ?? "Current User");
            ActivityChart?.UpdateChart(initialBreakdown);
            ActivityPie?.UpdateChart(
                initialBreakdown.RevitHours,
                initialBreakdown.MeetingHours,
                initialBreakdown.IdleHours
            );
            MeetingStatus?.UpdateStatus(null);
            
            // Force immediate activity update
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    // Trigger an immediate breakdown update after a short delay
                    var latestBreakdown = vm.TodayActivityBreakdown;
                    if (latestBreakdown != null)
                    {
                        ActivityChart?.UpdateChart(latestBreakdown);
                        ActivityPie?.UpdateChart(
                            latestBreakdown.RevitHours,
                            latestBreakdown.MeetingHours,
                            latestBreakdown.IdleHours
                        );
                    }
                }
                catch { }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }
        
        // Check if app should start minimized to tray
        if (App.StartMinimized)
        {
            // Delay minimizing to tray to ensure window is fully loaded
            Dispatcher.BeginInvoke(new Action(() =>
            {
                MinimizeToTray();
            }), System.Windows.Threading.DispatcherPriority.Loaded);
        }
    }
    
    private void ProjectSummaryGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is System.Windows.Controls.DataGrid grid && 
            grid.SelectedItem is ProjectWorkSummary selectedProject)
        {
            // Open project details in a new window
            var detailWindow = new ProjectDetailWindow(selectedProject);
            detailWindow.Owner = this;
            detailWindow.ShowDialog();
        }
    }
    
    /// <summary>
    /// Handle property changes for smooth project switching animation
    /// </summary>
    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == "LocalCurrentProject" && sender is MainViewModel vm)
        {
            var newProject = vm.LocalCurrentProject;
            
            // Only animate if project actually changed
            if (!string.IsNullOrEmpty(newProject) && newProject != _lastKnownProject && newProject != "-")
            {
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        // Find the project name border and text block by name
                        if (FindName("ProjectNameBorder") is System.Windows.Controls.Border border)
                        {
                            // Create fade out -> fade in animation
                            var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(150));
                            var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(250));
                            
                            // Add slide animation
                            var slideOut = new ThicknessAnimation(
                                new Thickness(0), 
                                new Thickness(-20, 0, 20, 0), 
                                TimeSpan.FromMilliseconds(150));
                            var slideIn = new ThicknessAnimation(
                                new Thickness(20, 0, -20, 0), 
                                new Thickness(0), 
                                TimeSpan.FromMilliseconds(250));
                            
                            fadeOut.Completed += (s, args) =>
                            {
                                // Start fade in after fade out completes
                                border.BeginAnimation(OpacityProperty, fadeIn);
                                border.BeginAnimation(MarginProperty, slideIn);
                            };
                            
                            // Start the animation
                            border.BeginAnimation(OpacityProperty, fadeOut);
                            border.BeginAnimation(MarginProperty, slideOut);
                        }
                    }
                    catch { }
                });
                
                _lastKnownProject = newProject;
            }
        }
    }
}
