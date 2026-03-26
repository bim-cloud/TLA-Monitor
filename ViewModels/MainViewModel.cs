using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AutodeskIDMonitor.Models;
using AutodeskIDMonitor.Services;
using AutodeskIDMonitor.Views;

namespace AutodeskIDMonitor.ViewModels;

public partial class MainViewModel : ObservableObject
{
    public const string APP_VERSION = "2.0.0";

    private readonly CloudService _cloudService;
    private readonly MonitorService _monitorService;
    private readonly LocalStorageService _localStorage;
    private readonly TimeTrackingService _timeTracking;
    private readonly ExcelExportService _excelExport;
    private readonly DispatcherTimer _refreshTimer;
    private readonly DispatcherTimer _syncTimer;

    private string _currentWindowsUser = "";
    private string _currentWindowsDisplayName = "";
    private string _currentMachineName = "";
    private string _currentAutodeskEmail = "";
    private bool _isAutodeskLoggedIn = false;

    [ObservableProperty] private bool _isAdminLoggedIn;
    [ObservableProperty] private string _adminButtonText = "🔐";
    [ObservableProperty] private string _adminStatusText = "🔒 Locked";
    [ObservableProperty] private string _statusBarText = "Ready";
    [ObservableProperty] private string _connectionStatusText = "LOGGED IN";
    [ObservableProperty] private string _cloudStatusText = "Cloud: ✓ (UAE)";
    [ObservableProperty] private bool _isConnected = true;
    [ObservableProperty] private bool _isAutodeskLoggedInDisplay = false;
    [ObservableProperty] private string _loginStatusText = "Not Logged In";
    [ObservableProperty] private string _lastRefreshText = "";
    [ObservableProperty] private int _selectedTabIndex;
    [ObservableProperty] private string _debugLog = "";
    [ObservableProperty] private string _country = "";
    [ObservableProperty] private string _office = "";
    [ObservableProperty] private bool _cloudLoggingEnabled;
    [ObservableProperty] private string _cloudApiUrl = "";
    [ObservableProperty] private string _cloudApiKey = "";
    [ObservableProperty] private bool _networkLoggingEnabled;
    [ObservableProperty] private string _networkLogPath = "";
    [ObservableProperty] private int _checkIntervalSeconds = 10;
    [ObservableProperty] private bool _startMinimized;
    [ObservableProperty] private bool _autoStartEnabled;
    [ObservableProperty] private string _cloudServerInfo = "";
    [ObservableProperty] private UserProfile? _selectedProfile;
    [ObservableProperty] private string _editName = "";
    [ObservableProperty] private string _editEmail = "";
    [ObservableProperty] private string _editPassword = "";
    [ObservableProperty] private int _totalRevitSessions;
    [ObservableProperty] private string _totalRevitSessionsText = "Total Revit Sessions: 0";
    [ObservableProperty] private UserSession? _selectedSession;
    [ObservableProperty] private string _editSessionName = "";
    
    // Version display
    public string AppVersionDisplay => $"v{APP_VERSION}";
    
    // Time Tracking Properties
    [ObservableProperty] private DateTime _selectedDate = DateTime.Today;
    [ObservableProperty] private DateTime _calendarMonth = DateTime.Today;
    [ObservableProperty] private string _workDayStatus = "";
    [ObservableProperty] private string _selectedUserKey = "";
    [ObservableProperty] private UserDetailedProfile? _selectedUserProfile;
    [ObservableProperty] private bool _isServerConnected = true;
    [ObservableProperty] private string _lastSyncTime = "";
    
    // Real-Time Dashboard Properties
    [ObservableProperty] private string _currentActivityState = "Offline";
    [ObservableProperty] private string _currentActivityIcon = "⚫";
    [ObservableProperty] private int _currentIdleSeconds;
    [ObservableProperty] private string _currentIdleText = "";
    [ObservableProperty] private int _localRevitCount;
    [ObservableProperty] private string _localCurrentProject = "";
    [ObservableProperty] private string _localSessionDuration = "00:00:00";
    [ObservableProperty] private DashboardStats _dashboardStats = new();
    [ObservableProperty] private LiveProjectSummary? _selectedLiveProject;
    
    // Idle and Meeting Tracking
    [ObservableProperty] private string _idleTimeDisplay = "0:00";
    [ObservableProperty] private bool _isInMeeting;
    [ObservableProperty] private string _meetingAppName = "";
    [ObservableProperty] private TimeSpan _meetingDuration;
    [ObservableProperty] private DailyActivityBreakdown? _todayActivityBreakdown;
    [ObservableProperty] private UserActivitySummary? _selectedUserActivity;
    
    // Update availability
    [ObservableProperty] private bool _isUpdateAvailable;
    [ObservableProperty] private string _latestVersion = "";
    
    // Public property for current user display name
    public string LocalUserDisplayName => _currentWindowsDisplayName;
    
    // Event for chart updates
    public event EventHandler<DailyActivityBreakdown>? ActivityBreakdownUpdated;
    public event EventHandler<List<MeetingAppInfo>>? MeetingStatusUpdated;
    public event EventHandler<UserActivitySummary>? UserActivitySelected;
    
    private readonly DispatcherTimer _realtimeTimer;

    public ObservableCollection<UserSession> CurrentSessions { get; } = new();
    public ObservableCollection<HistoryEntry> ActivityHistory { get; } = new();
    public ObservableCollection<UserProfile> UserProfiles { get; } = new();
    public ObservableCollection<ProjectInfo> Projects { get; } = new();
    
    // Time Tracking Collections
    public ObservableCollection<DailyWorkRecord> DailyRecords { get; } = new();
    public ObservableCollection<ProjectWorkSummary> ProjectSummaries { get; } = new();
    public ObservableCollection<CalendarDayEntry> CalendarDays { get; } = new();
    public ObservableCollection<UserDetailedProfile> CachedProfiles { get; } = new();
    
    // Selected daily record (for user-specific export)
    [ObservableProperty] private DailyWorkRecord? _selectedDailyRecord;
    
    // Real-Time Dashboard Collections
    public ObservableCollection<LiveUserActivity> LiveUsers { get; } = new();
    public ObservableCollection<LiveProjectSummary> LiveProjects { get; } = new();
    public ObservableCollection<RevitProcessInfo> LocalRevitProcesses { get; } = new();
    
    // All Users Activity Summary (Admin View)
    public ObservableCollection<UserActivitySummary> AllUsersActivity { get; } = new();

    private Dictionary<string, string> _userEmailMapping = new();
    
    // Store name overrides - key is "MachineName|WindowsUser", value is custom name
    private Dictionary<string, string> _nameOverrides = new();
    
    // Store original names from cloud - key is "MachineName|WindowsUser", value is original name
    private Dictionary<string, string> _originalNames = new();

    private static readonly List<(string Email, string Password)> DefaultUsers = new()
    {
        ("abhilash.rajesh@tangentlandscape.com", "Abhilash@123"),
        ("adhithyan.biju@tangentlandscape.com", "TLA@2025"),
        ("afsal.badharu@tangentlandscape.com", "TLA2024@209"),
        ("akshaya.jayakrishnan@tangentlandscape.com", "TLA2023_209"),
        ("amna.salim@tangentlandscape.com", "TLA2023_209"),
        ("ananthu.unnikrishnan@tangentlandscape.com", "TLA20252101"),
        ("anshu.jalaludeen@tangentlandscape.com", "TLA2024@213"),
        ("aparna.lakshmi@tangentlandscape.com", "TLA2024@213"),
        ("athira.sivdas@tangentlandscape.com", "TLA2025_209"),
        ("elbin.paulose@tangentlandscape.com", "Tangent#2026"),
        ("jesto.joy@tangentlandscape.com", "TLA2023@209"),
        ("jibin.issac@tangentlandscape.com", "TLA2024_101"),
        ("jithin.pavithran@tangentlandscape.com", "Tla2025@208"),
        ("jovanie.apa@tangentlandscape.com", "Tangent@2024"),
        ("laura.cruz@tangentlandscape.com", "TLA@2025"),
        ("lincy.kirubaharan@tangentlandscape.com", "Tangent@123"),
        ("maznaz.firoz@tangentlandscape.com", "TLA2025@101"),
        ("min.zaw@tangentlandscape.com", "Tla.MIN@12345"),
        ("mohamed.asif@tangentlandscape.com", "TLA2024@216"),
        ("narsha.abdura@tangentlandscape.com", "Tangent@2026"),
        ("noufal.palliparambil@tangentlandscape.com", "TLA2024@209"),
        ("rahul.jain@tangentlandscape.com", "Tangent@2025"),
        ("rashid.abdullah@tangentlandscape.com", "Tangent@2025"),
        ("rivin.wilson@tangentlandscape.com", "Tla2025@208"),
        ("sabarish.malayath@tangentlandscape.com", "Tla123@2025"),
        ("safas.umar@tangentlandscape.com", "TLA2024@215"),
        ("shasti.dharan@tangentlandscape.com", "TLA2021@209"),
        ("syed.tahir@tangentlandscape.com", "Tangent@2025"),
        ("toby.jose@tangentlandscape.com", "Tangent@123")
    };

    public MainViewModel()
    {
        _cloudService = new CloudService();
        _monitorService = new MonitorService();
        _localStorage = new LocalStorageService();
        _timeTracking = new TimeTrackingService(_localStorage);
        _excelExport = new ExcelExportService();
        
        // Get the actual logged-in user (not admin account when elevated)
        _currentWindowsUser = GetActualLoggedInUser();
        _currentMachineName = Environment.MachineName;
        
        // Use display name from first-run setup if available
        if (!string.IsNullOrEmpty(App.CurrentUserDisplayName))
        {
            _currentWindowsDisplayName = App.CurrentUserDisplayName;
        }
        else
        {
            _currentWindowsDisplayName = GetWindowsDisplayName();
        }
        
        // Also get email from setup
        if (!string.IsNullOrEmpty(App.CurrentUserEmail))
        {
            _currentAutodeskEmail = App.CurrentUserEmail;
        }

        AdminService.Instance.AdminStatusChanged += OnAdminStatusChanged;
        LoadSettings();
        InitializeDefaultProfiles();
        BuildUserEmailMapping();
        LoadCachedProfiles();
        
        // Load saved Activity History
        LoadSavedActivityHistory();
        
        // Enforce date boundaries - end sessions from previous days
        _localStorage.EnforceDateBoundaries();
        
        // Clean up stale sessions on startup
        var cleaned = _localStorage.CleanupStaleSessions(24);
        if (cleaned > 0)
            AddDebug($"Cleaned up {cleaned} stale sessions");

        var initialStatus = _monitorService.GetCurrentStatus();
        _currentAutodeskEmail = initialStatus.Email ?? "";
        _isAutodeskLoggedIn = initialStatus.IsLoggedIn;
        UpdateLoginStatusDisplay();
        UpdateWorkDayStatus();
        
        AddDebug($"Initial: {(_isAutodeskLoggedIn ? $"Logged in as {_currentAutodeskEmail}" : "Not logged in")}");

        _refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(CheckIntervalSeconds) };
        _refreshTimer.Tick += async (s, e) => await RefreshAllAsync();
        _refreshTimer.Start();

        _syncTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(CheckIntervalSeconds) };
        _syncTimer.Tick += async (s, e) => await SyncToCloudAsync();
        _syncTimer.Start();
        
        // Real-time update timer (every 1 second for live dashboard)
        _realtimeTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _realtimeTimer.Tick += (s, e) => UpdateRealTimeDisplay();
        _realtimeTimer.Start();

        // Subscribe to monitor service events
        _monitorService.StatusChanged += OnAutodeskStatusChanged;
        _monitorService.RevitActivityChanged += OnRevitActivityChanged;
        _monitorService.ActivityStateChanged += OnActivityStateChanged;
        _monitorService.StartMonitoring();

        _ = RefreshAllAsync();
        _ = SyncToCloudAsync();
        _ = LoadDailyRecordsAsync();
        
        // Initialize dashboard stats
        DashboardStats = new DashboardStats
        {
            CurrentTime = DateTime.Now.ToString("HH:mm:ss"),
            WorkDayStatus = _timeTracking.GetWorkDayStatus()
        };
        
        // Initialize auto-update service and lock detection
        InitializeUpdateService();
        
        // Schedule first-time setup check (runs after UI is loaded)
        Application.Current.Dispatcher.BeginInvoke(new Action(() =>
        {
            CheckFirstTimeSetup();
        }), System.Windows.Threading.DispatcherPriority.Loaded);
    }
    
    /// <summary>
    /// Initialize auto-update checking and system lock detection
    /// </summary>
    private void InitializeUpdateService()
    {
        try
        {
            // Subscribe to system lock/unlock events to pause monitoring
            UpdateService.Instance.SystemLockStateChanged += OnSystemLockStateChanged;
            
            // Subscribe to update available events
            UpdateService.Instance.UpdateAvailable += OnUpdateAvailable;
            
            // Start periodic update checks (every 4 hours)
            UpdateService.Instance.StartPeriodicUpdateCheck(APP_VERSION, TimeSpan.FromHours(4));
            
            // Clean up old update files
            UpdateService.Instance.CleanupOldUpdates();
            
            AddDebug($"Update service initialized - current version: {APP_VERSION}");
        }
        catch (Exception ex)
        {
            AddDebug($"Error initializing update service: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Handle system lock/unlock to pause/resume monitoring
    /// </summary>
    private void OnSystemLockStateChanged(object? sender, bool isLocked)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            if (isLocked)
            {
                // System is locked - pause all monitoring
                AddDebug("⏸️ System locked - pausing monitoring");
                _refreshTimer.Stop();
                _syncTimer.Stop();
                _realtimeTimer.Stop();
                _monitorService.StopMonitoring();
                
                // Update status display
                CurrentActivityState = "System Locked";
                CurrentActivityIcon = "🔒";
                StatusBarText = "Monitoring paused (system locked)";
            }
            else
            {
                // System is unlocked - resume monitoring
                AddDebug("▶️ System unlocked - resuming monitoring");
                _refreshTimer.Start();
                _syncTimer.Start();
                _realtimeTimer.Start();
                _monitorService.StartMonitoring();
                
                // Update status display
                StatusBarText = "Monitoring resumed";
                
                // Refresh data immediately
                _ = RefreshAllAsync();
            }
        });
    }
    
    /// <summary>
    /// Handle update available notification
    /// </summary>
    private void OnUpdateAvailable(object? sender, UpdateService.UpdateAvailableEventArgs e)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            try
            {
                AddDebug($"🔄 Update available: v{e.CurrentVersion} → v{e.LatestVersion}");
                
                // Set the update available flag for header button
                IsUpdateAvailable = true;
                LatestVersion = e.LatestVersion;
                
                // Show update notification window
                var updateWindow = new Views.UpdateNotificationWindow(e);
                updateWindow.Owner = Application.Current.MainWindow;
                updateWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                AddDebug($"Error showing update notification: {ex.Message}");
            }
        });
    }
    
    private void LoadSavedActivityHistory()
    {
        try
        {
            var savedHistory = _localStorage.LoadActivityHistory();
            foreach (var entry in savedHistory.Take(100)) // Load last 100 entries
            {
                ActivityHistory.Add(entry);
            }
            AddDebug($"Loaded {savedHistory.Count} activity history entries");
        }
        catch (Exception ex)
        {
            AddDebug($"Error loading activity history: {ex.Message}");
        }
    }

    private void OnRevitActivityChanged(object? sender, RevitActivityEventArgs e)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            // Update local Revit processes display
            LocalRevitProcesses.Clear();
            foreach (var proc in e.AllActiveProcesses)
            {
                LocalRevitProcesses.Add(proc);
            }
            
            LocalRevitCount = e.AllActiveProcesses.Count;
            LocalCurrentProject = e.AllActiveProcesses
                .FirstOrDefault(p => !string.IsNullOrEmpty(p.ProjectName))?.ProjectName ?? "";
            
            // Log activity
            var eventText = e.EventType switch
            {
                RevitEventType.Opened => $"Revit opened: {e.Process.RevitVersion}",
                RevitEventType.Closed => $"Revit closed: {e.Process.RevitVersion}",
                RevitEventType.ProjectChanged => $"Project changed: {e.Process.ProjectName}",
                _ => "Unknown"
            };
            
            AddDebug(eventText);
            
            // Track time
            if (e.EventType == RevitEventType.Opened && !string.IsNullOrEmpty(e.Process.ProjectName))
            {
                _localStorage.StartWorkSession(
                    _currentMachineName, _currentWindowsUser, _currentWindowsDisplayName,
                    e.Process.ProjectName, "Revit", e.Process.RevitVersion);
            }
            else if (e.EventType == RevitEventType.Closed && !string.IsNullOrEmpty(e.Process.ProjectName))
            {
                _localStorage.EndWorkSession(_currentMachineName, _currentWindowsUser, e.Process.ProjectName);
            }
            
            // Sync to cloud immediately
            _ = SyncToCloudAsync();
        });
    }

    private void OnActivityStateChanged(object? sender, ActivityStateEventArgs e)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            CurrentActivityState = e.CurrentState.ToString();
            CurrentActivityIcon = e.CurrentState switch
            {
                ActivityState.Active => "🟢",
                ActivityState.Idle => "🟡",
                ActivityState.Offline => "⚫",
                _ => "⚫"
            };
            CurrentIdleSeconds = e.IdleSeconds;
            CurrentIdleText = e.IdleSeconds > 60 
                ? $"{e.IdleSeconds / 60}m {e.IdleSeconds % 60}s idle"
                : e.IdleSeconds > 0 ? $"{e.IdleSeconds}s idle" : "";
            
            if (e.PreviousState != e.CurrentState)
            {
                AddDebug($"Activity state: {e.PreviousState} → {e.CurrentState}");
            }
        });
    }

    private void UpdateRealTimeDisplay()
    {
        try
        {
            // ========== STEP 1: TRACK ACTIVITY FIRST (before getting any data) ==========
            _monitorService.TrackAppActivity();
            
            // ========== STEP 2: UPDATE REVIT PROCESS LIST ==========
            var processes = _monitorService.GetActiveRevitProcesses();
            
            // Update LocalRevitProcesses collection for UI binding
            var currentProjectKeys = LocalRevitProcesses.Select(p => p.ProcessId).ToHashSet();
            var newProcessKeys = processes.Select(p => p.ProcessId).ToHashSet();
            
            // Remove closed processes
            var toRemove = LocalRevitProcesses.Where(p => !newProcessKeys.Contains(p.ProcessId)).ToList();
            foreach (var p in toRemove)
            {
                LocalRevitProcesses.Remove(p);
            }
            
            // Add or update processes
            foreach (var proc in processes)
            {
                var existing = LocalRevitProcesses.FirstOrDefault(p => p.ProcessId == proc.ProcessId);
                if (existing == null)
                {
                    LocalRevitProcesses.Add(proc);
                }
                else
                {
                    // Update duration display
                    existing.LastActivity = proc.LastActivity;
                }
            }
            
            // ========== STEP 3: UPDATE CURRENT PROJECT (use FOREGROUND project) ==========
            LocalRevitCount = processes.Count;
            
            // Get the ACTUAL foreground project (what user is actively working on)
            var foregroundProject = _monitorService.GetCurrentForegroundProject();
            if (!string.IsNullOrEmpty(foregroundProject))
            {
                LocalCurrentProject = foregroundProject;
            }
            else
            {
                // Fallback to first open project
                LocalCurrentProject = processes
                    .FirstOrDefault(p => !string.IsNullOrEmpty(p.ProjectName))?.ProjectName ?? "-";
            }
            
            // ========== STEP 4: UPDATE SESSION DURATION ==========
            // Calculate session duration - use ACTIVE Revit time from hourly buckets (TODAY only)
            if (processes.Any())
            {
                // Get TODAY's active Revit time from MonitorService (accurate, hourly-bucket based)
                var activeRevitTime = _monitorService.GetTodayActiveRevitTime();
                
                // Sanity check - can't exceed hours since work start (8 AM)
                var hoursSince8AM = Math.Max(0, DateTime.Now.Hour - 8 + DateTime.Now.Minute / 60.0);
                var maxTime = TimeSpan.FromHours(Math.Min(hoursSince8AM, 12));
                
                if (activeRevitTime > maxTime)
                {
                    activeRevitTime = maxTime;
                }
                
                // Display format
                LocalSessionDuration = $"{(int)activeRevitTime.TotalHours:D2}:{activeRevitTime.Minutes:D2}:{activeRevitTime.Seconds:D2}";
            }
            else
            {
                LocalSessionDuration = "00:00:00";
            }
            
            // Update dashboard stats
            DashboardStats.CurrentTime = DateTime.Now.ToString("HH:mm:ss");
            DashboardStats.WorkDayStatus = _timeTracking.GetWorkDayStatus();
            
            // Update idle info
            CurrentIdleSeconds = _monitorService.GetIdleSeconds();
            CurrentIdleText = CurrentIdleSeconds > 60 
                ? $"{CurrentIdleSeconds / 60}m {CurrentIdleSeconds % 60}s idle"
                : CurrentIdleSeconds > 0 ? $"{CurrentIdleSeconds}s idle" : "";
            
            // Update idle time display (formatted) - TODAY ONLY
            var idleTime = _monitorService.GetTodayIdleTime();
            
            // Additional sanity check - idle can't exceed hours since midnight
            var hoursSinceMidnight = (DateTime.Now - DateTime.Today).TotalHours;
            if (idleTime.TotalHours > hoursSinceMidnight)
            {
                idleTime = TimeSpan.FromHours(hoursSinceMidnight);
            }
            
            if (idleTime.TotalMinutes >= 1 && idleTime.TotalHours < 24)
            {
                IdleTimeDisplay = idleTime.TotalHours >= 1 
                    ? $"{(int)idleTime.TotalHours}h {idleTime.Minutes}m"
                    : $"{idleTime.Minutes}m";
            }
            else
            {
                IdleTimeDisplay = "0m";
            }
            
            // ========== STEP 6: UPDATE MEETING STATUS ==========
            var meetings = _monitorService.GetActiveMeetings();
            var activeMeeting = meetings.FirstOrDefault(m => m.IsInActiveMeeting);
            
            IsInMeeting = activeMeeting != null;
            if (activeMeeting != null)
            {
                MeetingAppName = activeMeeting.AppName;
                // Use the dedicated method for accurate real-time duration
                MeetingDuration = _monitorService.GetCurrentMeetingDuration();
            }
            else
            {
                MeetingDuration = TimeSpan.Zero;
            }
            
            // Raise event for meeting status update
            MeetingStatusUpdated?.Invoke(this, meetings);
            
            // Update activity breakdown EVERY SECOND for current user (local-first, instant updates)
            TodayActivityBreakdown = _monitorService.GetDailyActivityBreakdown();
            ActivityBreakdownUpdated?.Invoke(this, TodayActivityBreakdown);
            
            // Update hourly breakdown property for chart binding
            OnPropertyChanged(nameof(TodayActivityBreakdown));
            
            // Auto-save activity data every 30 seconds (only for current user, today)
            if (DateTime.Now.Second % 30 == 0 && SelectedDate.Date == DateTime.Today)
            {
                _localStorage.SaveUserDailyActivity(_currentWindowsUser, DateTime.Today, TodayActivityBreakdown);
            }
            
            // Force property changed for LocalRevitProcesses items (for duration updates)
            OnPropertyChanged(nameof(LocalRevitProcesses));
        }
        catch { }
    }

    private void LoadCachedProfiles()
    {
        try
        {
            var cached = _localStorage.GetCachedProfiles();
            CachedProfiles.Clear();
            foreach (var profile in cached)
            {
                CachedProfiles.Add(profile);
            }
            LastSyncTime = $"Last sync: {_localStorage.GetLastSyncTime():HH:mm:ss}";
        }
        catch (Exception ex)
        {
            AddDebug($"Error loading cached profiles: {ex.Message}");
        }
    }

    private void UpdateWorkDayStatus()
    {
        WorkDayStatus = _timeTracking.GetWorkDayStatus();
    }

    partial void OnSelectedDateChanged(DateTime value)
    {
        // Save current day's activity before switching dates
        if (TodayActivityBreakdown != null)
        {
            _localStorage.SaveUserDailyActivity(_currentWindowsUser, DateTime.Today, TodayActivityBreakdown);
        }
        
        // Load activity data for selected date
        if (value.Date == DateTime.Today)
        {
            // Live data for today
            TodayActivityBreakdown = _monitorService.GetDailyActivityBreakdown();
        }
        else
        {
            // Historical data for past dates
            var historicalData = _localStorage.GetUserDailyActivity(_currentWindowsUser, value.Date);
            if (historicalData != null)
            {
                TodayActivityBreakdown = new DailyActivityBreakdown
                {
                    Date = historicalData.Date,
                    RevitMinutes = historicalData.RevitMinutes,
                    MeetingMinutes = historicalData.MeetingMinutes,
                    TotalIdleMinutes = historicalData.IdleMinutes,
                    OtherMinutes = historicalData.OtherMinutes,
                    HourlyBreakdown = historicalData.HourlyBreakdown.Select(h => new HourlyActivity
                    {
                        Hour = h.Hour,
                        RevitMinutes = h.RevitMinutes,
                        MeetingMinutes = h.MeetingMinutes,
                        IdleMinutes = h.IdleMinutes,
                        OtherMinutes = h.OtherMinutes,
                        IsWorkHour = h.Hour >= 8 && h.Hour < 18
                    }).ToList()
                };
            }
            else
            {
                // No historical data - show empty
                TodayActivityBreakdown = new DailyActivityBreakdown { Date = value.Date };
            }
        }
        
        ActivityBreakdownUpdated?.Invoke(this, TodayActivityBreakdown);
        _ = LoadDailyRecordsAsync();
    }

    partial void OnCalendarMonthChanged(DateTime value)
    {
        LoadCalendarDays();
    }

    private async Task LoadDailyRecordsAsync()
    {
        try
        {
            var records = _localStorage.GetDailyRecords(SelectedDate);
            var summaries = _timeTracking.GetProjectStats(SelectedDate);
            
            // Get LOCAL real-time project data for current user
            var localRevitProcesses = _monitorService.GetActiveRevitProcesses();
            var localOpenProjects = localRevitProcesses
                .Where(p => !string.IsNullOrEmpty(p.ProjectName))
                .Select(p => p.ProjectName)
                .Distinct()
                .ToList();
            var localProjectsStr = localOpenProjects.Count > 0 
                ? string.Join(", ", localOpenProjects) 
                : "";
            
            // If no local records for today, build from live cloud sessions
            if (records.Count == 0 && SelectedDate.Date == DateTime.Today)
            {
                var sessions = await _cloudService.GetSessionsAsync();
                
                // Build records from active sessions
                records = sessions
                    .Where(s => s.IsLoggedIn)
                    .Select(s => 
                    {
                        // Check if this is the current user - use local data
                        bool isCurrentUser = s.MachineName.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase) &&
                                            s.WindowsUser.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase);
                        
                        string projectsWorked;
                        if (isCurrentUser && !string.IsNullOrEmpty(localProjectsStr))
                        {
                            projectsWorked = localProjectsStr;
                        }
                        else if (s.OpenProjects?.Any() == true)
                        {
                            projectsWorked = string.Join(", ", s.OpenProjects.Select(p => p.ProjectName).Where(n => !string.IsNullOrEmpty(n)));
                        }
                        else if (!string.IsNullOrEmpty(s.CurrentProject))
                        {
                            projectsWorked = s.CurrentProject;
                        }
                        else
                        {
                            projectsWorked = "-";
                        }
                        
                        return new DailyWorkRecord
                        {
                            Date = DateTime.Today,
                            UserId = s.WindowsUser,
                            UserName = s.GetDisplayName(),
                            MachineId = s.MachineName,
                            // Convert from server UTC time to UAE local time (UTC+4)
                            FirstActivity = s.LastSeen.AddHours(4).AddHours(-1), // Estimate (UAE time - 1 hour)
                            LastActivity = s.LastSeen.AddHours(4), // UAE time
                            // Use conservative estimate - will be updated with actual active time for current user
                            TotalHours = s.RevitSessionCount > 0 ? 0.5 : 0, // Conservative estimate
                            RegularHours = s.RevitSessionCount > 0 ? 0.5 : 0,
                            OvertimeHours = 0,
                            ProjectsWorked = projectsWorked
                        };
                    })
                    .ToList();
                
                // Update current user's record with STRICT active time from MonitorService
                var currentUserRecord = records.FirstOrDefault(r => 
                    r.MachineId.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase) &&
                    r.UserId.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase));
                if (currentUserRecord != null)
                {
                    var activeRevitTime = _monitorService.GetTodayActiveRevitTime();
                    if (activeRevitTime.TotalHours > 0)
                    {
                        currentUserRecord.TotalHours = Math.Round(activeRevitTime.TotalHours, 2);
                        currentUserRecord.RegularHours = Math.Min(currentUserRecord.TotalHours, 8.0);
                        currentUserRecord.OvertimeHours = Math.Max(0, currentUserRecord.TotalHours - 8.0);
                    }
                }
                
                // Build project summaries from cloud sessions + local data
                var projectUserPairs = new List<(string Project, CloudSessionInfo User)>();
                
                foreach (var s in sessions.Where(s => s.IsLoggedIn))
                {
                    // Check if this is the current user - use local data
                    bool isCurrentUser = s.MachineName.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase) &&
                                        s.WindowsUser.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase);
                    
                    if (isCurrentUser && localOpenProjects.Count > 0)
                    {
                        foreach (var proj in localOpenProjects)
                        {
                            projectUserPairs.Add((proj, s));
                        }
                    }
                    else if (s.OpenProjects?.Any() == true)
                    {
                        foreach (var p in s.OpenProjects)
                        {
                            if (!string.IsNullOrEmpty(p.ProjectName))
                                projectUserPairs.Add((p.ProjectName, s));
                        }
                    }
                    else if (!string.IsNullOrEmpty(s.CurrentProject))
                    {
                        projectUserPairs.Add((s.CurrentProject, s));
                    }
                }
                
                summaries = projectUserPairs
                    .GroupBy(x => x.Project)
                    .Select(g => new ProjectWorkSummary
                    {
                        ProjectName = g.Key,
                        // Use 0.5 hour estimate per user working on project (conservative estimate)
                        TotalHours = g.Select(x => x.User.WindowsUser).Distinct().Count() * 0.5,
                        UserCount = g.Select(x => x.User.WindowsUser).Distinct().Count(),
                        UsersWorking = string.Join(", ", g.Select(x => x.User.GetDisplayName()).Distinct()),
                        UserTimes = g.Select(x => new UserProjectTime
                        {
                            UserName = x.User.GetDisplayName(),
                            Hours = 0.5  // Conservative estimate
                        }).ToList()
                    }).ToList();
                
                // NOW apply STRICT active time from MonitorService for current user's projects
                var localActiveHours = _monitorService.GetAllProjectActiveHours();
                foreach (var summary in summaries)
                {
                    // Check if we have actual active time tracked for this project
                    if (localActiveHours.TryGetValue(summary.ProjectName, out var activeHours) && activeHours > 0)
                    {
                        // Update with actual tracked time for current user
                        var currentUserTime = summary.UserTimes?.FirstOrDefault(ut => 
                            ut.UserName.Equals(_currentWindowsDisplayName, StringComparison.OrdinalIgnoreCase));
                        if (currentUserTime != null)
                        {
                            // Replace estimate with actual active time
                            var oldHours = currentUserTime.Hours;
                            currentUserTime.Hours = activeHours;
                            summary.TotalHours = summary.TotalHours - oldHours + activeHours;
                        }
                        else if (summary.UsersWorking.Contains(_currentWindowsDisplayName))
                        {
                            // Add active time for current user
                            summary.TotalHours += activeHours - 0.5; // Subtract estimate, add actual
                        }
                    }
                }
            }
            else
            {
                // Update current user's record with LOCAL project data (most accurate)
                foreach (var record in records)
                {
                    if (record.MachineId.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase) &&
                        record.UserId.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase))
                    {
                        // Replace "Revit Session" with actual project names if available
                        if (!string.IsNullOrEmpty(localProjectsStr) && 
                            (record.ProjectsWorked == "Revit Session" || string.IsNullOrEmpty(record.ProjectsWorked) || record.ProjectsWorked == "-"))
                        {
                            record.ProjectsWorked = localProjectsStr;
                        }
                        else if (!string.IsNullOrEmpty(localProjectsStr) && record.ProjectsWorked == "Revit Session")
                        {
                            record.ProjectsWorked = localProjectsStr;
                        }
                    }
                }
                
                // Update summaries - replace "Revit Session" with actual projects for current user
                if (!string.IsNullOrEmpty(localProjectsStr))
                {
                    // Remove any "Revit Session" summary and replace with real project summaries
                    var revitSessionSummary = summaries.FirstOrDefault(s => s.ProjectName == "Revit Session");
                    if (revitSessionSummary != null)
                    {
                        summaries.Remove(revitSessionSummary);
                        
                        // Add real projects for current user
                        foreach (var proj in localOpenProjects)
                        {
                            var existingSummary = summaries.FirstOrDefault(s => s.ProjectName.Equals(proj, StringComparison.OrdinalIgnoreCase));
                            if (existingSummary != null)
                            {
                                // Add current user to existing summary if not already there
                                if (!existingSummary.UsersWorking.Contains(_currentWindowsDisplayName))
                                {
                                    existingSummary.UsersWorking += ", " + _currentWindowsDisplayName;
                                    existingSummary.UserCount++;
                                    existingSummary.TotalHours += revitSessionSummary.TotalHours / localOpenProjects.Count;
                                }
                            }
                            else
                            {
                                // Create new summary for this project
                                summaries.Add(new ProjectWorkSummary
                                {
                                    ProjectName = proj,
                                    TotalHours = revitSessionSummary.TotalHours / localOpenProjects.Count,
                                    UserCount = 1,
                                    UsersWorking = _currentWindowsDisplayName,
                                    UserTimes = new List<UserProjectTime>
                                    {
                                        new UserProjectTime { UserName = _currentWindowsDisplayName, Hours = revitSessionSummary.TotalHours / localOpenProjects.Count }
                                    }
                                });
                            }
                        }
                    }
                }
            }

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                DailyRecords.Clear();
                foreach (var record in records.OrderBy(r => r.UserName))
                {
                    DailyRecords.Add(record);
                }

                // Filter out invalid project names from summaries
                var invalidProjectNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "Revit Session", "Home", "Start", "Revit", "Recent", "New", "Open",
                    "Models", "Families", "Templates", "Samples", "Autodesk", 
                    "Architecture", "Structure", "MEP", "Project Opening Time"
                };
                
                ProjectSummaries.Clear();
                foreach (var summary in summaries
                    .Where(s => !string.IsNullOrEmpty(s.ProjectName) && !invalidProjectNames.Contains(s.ProjectName))
                    .OrderByDescending(s => s.TotalHours))
                {
                    ProjectSummaries.Add(summary);
                }
            });
            
            // Load all users' activity from server (for admin view)
            // Always load activity data to show team activity
            await LoadAllUsersActivityAsync();
        }
        catch (Exception ex)
        {
            AddDebug($"Error loading daily records: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Load all users' activity breakdown from server (Admin view)
    /// </summary>
    private async Task LoadAllUsersActivityAsync()
    {
        try
        {
            var queryDate = SelectedDate.Date;
            var today = DateTime.Today;
            List<UserActivitySummary> activities;
            bool isLiveData = false;
            string message = "";
            
            // Check if querying TODAY vs HISTORICAL date
            if (queryDate == today)
            {
                // TODAY: Use live data from /api/activity/all + local data
                activities = await _cloudService.GetAllUsersActivityAsync();
                isLiveData = true;
                
                // If no activities from /api/activity/all, build from sessions
                if (activities.Count == 0)
                {
                    var sessions = await _cloudService.GetSessionsAsync();
                    activities = sessions.Select(s => new UserActivitySummary
                    {
                        WindowsUser = s.WindowsUser,
                        WindowsDisplayName = s.GetDisplayName(),
                        MachineName = s.MachineName,
                        Date = DateTime.Today,
                        RevitHours = s.TodayRevitHours,
                        MeetingHours = s.TodayMeetingHours,
                        IdleHours = s.TodayIdleHours,
                        OtherHours = s.TodayOtherHours,
                        OvertimeHours = s.TodayOvertimeHours,
                        TotalHours = s.TodayTotalHours,
                        IsInMeeting = s.IsInMeeting,
                        MeetingApp = s.MeetingApp,
                        IdleSeconds = s.IdleSeconds,
                        ActivityState = s.ActivityState,
                        CurrentProject = s.CurrentProject
                    }).ToList();
                }
                
                // Update current user's data with LOCAL data (most accurate)
                var localBreakdown = _monitorService.GetDailyActivityBreakdown();
                var localMeetings = _monitorService.GetActiveMeetings();
                var activeMeeting = localMeetings.FirstOrDefault(m => m.IsInActiveMeeting);
                var localRevitProcesses = _monitorService.GetActiveRevitProcesses();
                var localIdleSeconds = _monitorService.GetIdleSeconds();
                
                var currentUserActivity = activities.FirstOrDefault(a => 
                    a.MachineName.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase) &&
                    a.WindowsUser.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase));
                
                if (currentUserActivity != null)
                {
                    // Update with local data (more accurate than server data)
                    currentUserActivity.RevitHours = localBreakdown.RevitHours;
                    currentUserActivity.MeetingHours = localBreakdown.MeetingHours;
                    currentUserActivity.IdleHours = localBreakdown.IdleHours;
                    currentUserActivity.TotalHours = localBreakdown.TotalWorkHours;
                    currentUserActivity.OvertimeHours = Math.Max(0, localBreakdown.TotalWorkHours - 9.0);
                    currentUserActivity.IsInMeeting = activeMeeting != null;
                    currentUserActivity.MeetingApp = activeMeeting?.AppName ?? "";
                    currentUserActivity.IdleSeconds = localIdleSeconds;
                    currentUserActivity.ActivityState = _monitorService.GetActivityStateString();
                    currentUserActivity.CurrentProject = LocalCurrentProject; // Use foreground project
                    
                    // Add project breakdowns from local tracking
                    currentUserActivity.ProjectBreakdowns = GetLocalProjectBreakdowns();
                }
                else
                {
                    // Add current user if not in list
                    var newActivity = new UserActivitySummary
                    {
                        WindowsUser = _currentWindowsUser,
                        WindowsDisplayName = _currentWindowsDisplayName,
                        MachineName = _currentMachineName,
                        Date = DateTime.Today,
                        RevitHours = localBreakdown.RevitHours,
                        MeetingHours = localBreakdown.MeetingHours,
                        IdleHours = localBreakdown.IdleHours,
                        TotalHours = localBreakdown.TotalWorkHours,
                        OvertimeHours = Math.Max(0, localBreakdown.TotalWorkHours - 9.0),
                        IsInMeeting = activeMeeting != null,
                        MeetingApp = activeMeeting?.AppName ?? "",
                        IdleSeconds = localIdleSeconds,
                        ActivityState = _monitorService.GetActivityStateString(),
                        CurrentProject = LocalCurrentProject, // Use foreground project
                        ProjectBreakdowns = GetLocalProjectBreakdowns()
                    };
                    activities.Add(newActivity);
                }
            }
            else if (queryDate > today)
            {
                // FUTURE DATE: No data available
                activities = new List<UserActivitySummary>();
                message = "No data available for future dates";
                AddDebug($"Date {queryDate:yyyy-MM-dd}: {message}");
            }
            else
            {
                // HISTORICAL DATE: Query server for that specific date
                var result = await _cloudService.GetActivityByDateAsync(queryDate);
                activities = result.Activities;
                isLiveData = result.IsLive;
                message = result.Message;
                
                if (!string.IsNullOrEmpty(message))
                {
                    AddDebug($"Date {queryDate:yyyy-MM-dd}: {message}");
                }
            }
            
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                AllUsersActivity.Clear();
                foreach (var activity in activities.OrderBy(a => a.WindowsDisplayName))
                {
                    // Mark current user
                    activity.IsCurrentUser = activity.WindowsUser.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase) &&
                                             activity.MachineName.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase);
                    AllUsersActivity.Add(activity);
                }
                
                // Update status message
                if (queryDate != today)
                {
                    var dateStr = queryDate.ToString("MMMM dd, yyyy");
                    if (activities.Count == 0)
                    {
                        WorkDayStatus = $"📅 {dateStr} - No data recorded";
                    }
                    else
                    {
                        WorkDayStatus = $"📅 {dateStr} - Historical data ({activities.Count} users)";
                    }
                }
                else
                {
                    WorkDayStatus = _timeTracking.GetWorkDayStatus();
                }
            });
        }
        catch (Exception ex)
        {
            AddDebug($"Error loading users activity: {ex.Message}");
        }
    }

    private void LoadCalendarDays()
    {
        try
        {
            var entries = _localStorage.GetCalendarMonth(CalendarMonth.Year, CalendarMonth.Month);
            CalendarDays.Clear();
            foreach (var entry in entries)
            {
                CalendarDays.Add(entry);
            }
        }
        catch (Exception ex)
        {
            AddDebug($"Error loading calendar: {ex.Message}");
        }
    }

    public void LoadUserDetails(string machineId, string userId)
    {
        try
        {
            SelectedUserKey = $"{machineId}|{userId}";
            SelectedUserProfile = _timeTracking.GetDetailedProfile(machineId, userId);
            if (SelectedUserProfile != null)
            {
                LoadCalendarDays();
            }
        }
        catch (Exception ex)
        {
            AddDebug($"Error loading user details: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Called when a user is selected in the AllUsersActivity table
    /// Shows detailed project breakdown in a popup
    /// </summary>
    partial void OnSelectedUserActivityChanged(UserActivitySummary? value)
    {
        if (value != null)
        {
            // Raise event for chart update
            UserActivitySelected?.Invoke(this, value);
            AddDebug($"User selected: {value.WindowsDisplayName} - Revit: {value.RevitHours:F1}h, Meeting: {value.MeetingHours:F1}h");
            
            // Show project breakdown dialog
            ShowUserProjectBreakdown(value);
        }
    }
    
    /// <summary>
    /// Show detailed project breakdown for a user - Uses LOCAL DATA for current user
    /// </summary>
    private void ShowUserProjectBreakdown(UserActivitySummary user)
    {
        try
        {
            // Check if this is the CURRENT USER
            bool isCurrentUser = user.WindowsUser.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase) &&
                                 user.MachineName.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase);
            
            // For current user, use LIVE LOCAL DATA (not cloud data)
            DailyActivityBreakdown? localBreakdown = null;
            if (isCurrentUser)
            {
                localBreakdown = TodayActivityBreakdown ?? _monitorService.GetDailyActivityBreakdown();
                
                // Update user object with live local data
                if (localBreakdown != null)
                {
                    user.RevitHours = localBreakdown.RevitHours;
                    user.MeetingHours = localBreakdown.MeetingHours;
                    user.IdleHours = localBreakdown.IdleHours;
                    user.TotalHours = localBreakdown.TotalWorkHours;
                    user.ActivityState = _monitorService.GetActivityStateString();
                    user.ProjectBreakdowns = GetLocalProjectBreakdowns();
                    user.CurrentProject = LocalCurrentProject;
                }
            }
            
            // Show the new Activity Detail Window
            var detailWindow = new Views.UserActivityDetailWindow(user, localBreakdown, isCurrentUser);
            detailWindow.Owner = Application.Current.MainWindow;
            detailWindow.ShowDialog();
        }
        catch (Exception ex)
        {
            AddDebug($"Error showing project breakdown: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Get project-level activity breakdowns from local tracking
    /// </summary>
    private List<ProjectActivityBreakdown> GetLocalProjectBreakdowns()
    {
        var breakdowns = new List<ProjectActivityBreakdown>();
        
        try
        {
            // Get project active hours from MonitorService
            var projectHours = _monitorService.GetAllProjectActiveHours();
            var activeRevitProcesses = _monitorService.GetActiveRevitProcesses();
            
            foreach (var kvp in projectHours)
            {
                var projectName = kvp.Key;
                var activeHours = kvp.Value;
                
                // Check if this project is currently active
                var isActive = activeRevitProcesses.Any(p => 
                    p.ProjectName.Equals(projectName, StringComparison.OrdinalIgnoreCase));
                
                breakdowns.Add(new ProjectActivityBreakdown
                {
                    ProjectName = projectName,
                    ActiveRevitHours = Math.Round(activeHours, 2),
                    IdleHours = 0, // Could be calculated from hourly buckets
                    OtherHours = 0,
                    IsCurrentlyActive = isActive,
                    LastActivity = DateTime.Now
                });
            }
            
            // Add any active projects not in the hours list
            foreach (var process in activeRevitProcesses)
            {
                if (!string.IsNullOrEmpty(process.ProjectName) && 
                    !breakdowns.Any(b => b.ProjectName.Equals(process.ProjectName, StringComparison.OrdinalIgnoreCase)))
                {
                    breakdowns.Add(new ProjectActivityBreakdown
                    {
                        ProjectName = process.ProjectName,
                        ActiveRevitHours = process.Duration.TotalHours,
                        IsCurrentlyActive = true,
                        LastActivity = DateTime.Now
                    });
                }
            }
        }
        catch (Exception ex)
        {
            AddDebug($"Error getting project breakdowns: {ex.Message}");
        }
        
        return breakdowns.OrderByDescending(b => b.ActiveRevitHours).ToList();
    }

    private void BuildUserEmailMapping()
    {
        _userEmailMapping.Clear();
        foreach (var profile in UserProfiles)
            _userEmailMapping[profile.Name.ToLower().Trim()] = profile.Email.ToLower();
    }

    private string GetWindowsDisplayName()
    {
        try
        {
            // FIRST: Check if user has set up their profile
            var userSettings = _localStorage.GetUserProfileSettings();
            if (userSettings.SetupCompleted && !string.IsNullOrEmpty(userSettings.DisplayName))
            {
                return userSettings.DisplayName;
            }
            
            // Get the actual logged-in user (console session), not the admin account
            var actualUser = GetActualLoggedInUser();
            
            // If username is generic "User", we'll need first-time setup
            if (actualUser.ToLower() == "user" || actualUser.ToLower() == "admin")
            {
                return actualUser; // Will trigger first-time setup check
            }
            
            using var searcher = new System.Management.ManagementObjectSearcher(
                $"SELECT FullName FROM Win32_UserAccount WHERE Name='{actualUser}' AND Domain='{Environment.UserDomainName}'");
            foreach (var obj in searcher.Get())
            {
                var fullName = obj["FullName"]?.ToString();
                if (!string.IsNullOrEmpty(fullName)) return fullName;
            }
            
            // Fallback: Format username nicely (e.g., "anshu.jalaludeen" -> "Anshu Jalaludeen")
            return FormatUserName(actualUser);
        }
        catch { }
        return FormatUserName(GetActualLoggedInUser());
    }
    
    /// <summary>
    /// Check if first-time setup is needed and show dialog if necessary
    /// </summary>
    public void CheckFirstTimeSetup()
    {
        try
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                // Determine if this is first-time (no stored profile/creds)
                bool isFirstTime = _localStorage.NeedsFirstTimeSetup() ||
                    !Services.UserCredentialService.Instance.IsSetup(_currentAutodeskEmail);

                var prefillEmail = _currentAutodeskEmail;
                var prefillName  = _currentWindowsDisplayName;

                var loginWindow = new Views.UserLoginWindow(isFirstTime, prefillEmail, prefillName);
                var result = loginWindow.ShowDialog();

                if (result == true && loginWindow.LoginSuccess)
                {
                    // Update display name and email from login
                    if (!string.IsNullOrEmpty(loginWindow.LoggedInName))
                        _currentWindowsDisplayName = loginWindow.LoggedInName;
                    if (!string.IsNullOrEmpty(loginWindow.LoggedInEmail))
                        _currentAutodeskEmail = loginWindow.LoggedInEmail;

                    // Save profile settings
                    _localStorage.SaveUserProfileSettings(_currentWindowsDisplayName, _currentAutodeskEmail);
                    OnPropertyChanged(nameof(LocalUserDisplayName));
                    AddDebug($"User login OK: {_currentWindowsDisplayName} <{_currentAutodeskEmail}>");
                }
                else if (!loginWindow.LoginSuccess)
                {
                    // User closed window without logging in – exit the app
                    Application.Current.Shutdown();
                }
            });
        }
        catch (Exception ex)
        {
            AddDebug($"Login window error: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Gets the actual logged-in user from the console session, even when running as administrator.
    /// This ensures we track the real user, not "Administrator" or a different elevated account.
    /// </summary>
    private string GetActualLoggedInUser()
    {
        try
        {
            // Method 1: Get the owner of explorer.exe (most reliable for actual logged-in user)
            var explorerProcesses = Process.GetProcessesByName("explorer");
            if (explorerProcesses.Length > 0)
            {
                foreach (var explorer in explorerProcesses)
                {
                    try
                    {
                        var owner = GetProcessOwner(explorer.Id);
                        if (!string.IsNullOrEmpty(owner) && 
                            !owner.Equals("SYSTEM", StringComparison.OrdinalIgnoreCase) &&
                            !owner.Equals("Administrator", StringComparison.OrdinalIgnoreCase))
                        {
                            // Extract just the username if it's in DOMAIN\User format
                            if (owner.Contains("\\"))
                                owner = owner.Split('\\').Last();
                            return owner;
                        }
                    }
                    catch { }
                }
            }
            
            // Method 2: Query WMI for interactive logon sessions
            using var searcher = new System.Management.ManagementObjectSearcher(
                "SELECT * FROM Win32_ComputerSystem");
            foreach (var obj in searcher.Get())
            {
                var userName = obj["UserName"]?.ToString();
                if (!string.IsNullOrEmpty(userName))
                {
                    // Extract just the username if it's in DOMAIN\User format
                    if (userName.Contains("\\"))
                        userName = userName.Split('\\').Last();
                    return userName;
                }
            }
        }
        catch { }
        
        // Fallback to Environment.UserName
        return Environment.UserName;
    }
    
    /// <summary>
    /// Gets the owner of a process by ID using WMI
    /// </summary>
    private string GetProcessOwner(int processId)
    {
        try
        {
            var query = $"SELECT * FROM Win32_Process WHERE ProcessId = {processId}";
            using var searcher = new System.Management.ManagementObjectSearcher(query);
            foreach (System.Management.ManagementObject obj in searcher.Get())
            {
                var outParams = obj.InvokeMethod("GetOwner", null, null);
                if (outParams != null)
                {
                    var user = outParams["User"]?.ToString();
                    var domain = outParams["Domain"]?.ToString();
                    if (!string.IsNullOrEmpty(user))
                    {
                        return string.IsNullOrEmpty(domain) ? user : $"{domain}\\{user}";
                    }
                }
            }
        }
        catch { }
        return "";
    }
    
    /// <summary>
    /// Formats a username like "anshu.jalaludeen" into "Anshu Jalaludeen"
    /// </summary>
    private string FormatUserName(string userName)
    {
        if (string.IsNullOrEmpty(userName)) return userName;
        
        // Handle domain\user format
        if (userName.Contains("\\"))
            userName = userName.Split('\\').Last();
            
        // Replace dots and underscores with spaces, then title case
        var parts = userName.Replace(".", " ").Replace("_", " ").Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return string.Join(" ", parts.Select(p => 
            char.ToUpper(p[0]) + (p.Length > 1 ? p.Substring(1).ToLower() : "")));
    }

    private void OnAdminStatusChanged(object? sender, bool isAdmin)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            IsAdminLoggedIn = isAdmin;
            AdminButtonText = isAdmin ? "🔓" : "🔐";
            AdminStatusText = isAdmin ? "🔓 Admin" : "🔒 Locked";
            AddDebug(isAdmin ? "Admin logged in" : "Admin logged out");
        });
    }

    private void OnAutodeskStatusChanged(object? sender, AutodeskStatus status)
    {
        Application.Current.Dispatcher.Invoke(async () =>
        {
            var previousEmail = _currentAutodeskEmail;
            _currentAutodeskEmail = status.Email ?? "";
            _isAutodeskLoggedIn = status.IsLoggedIn;
            UpdateLoginStatusDisplay();

            if (!string.IsNullOrEmpty(previousEmail) && !string.IsNullOrEmpty(_currentAutodeskEmail) && previousEmail != _currentAutodeskEmail)
            {
                AddDebug($"ID CHANGE: {previousEmail} → {_currentAutodeskEmail}");
                ActivityHistory.Insert(0, new HistoryEntry
                {
                    Timestamp = DateTime.Now, EventType = "ID Change", PersonName = _currentWindowsDisplayName,
                    PreviousEmail = previousEmail, AutodeskEmail = _currentAutodeskEmail, MachineId = _currentMachineName
                });
            }
            else if (!string.IsNullOrEmpty(_currentAutodeskEmail) && string.IsNullOrEmpty(previousEmail))
            {
                ActivityHistory.Insert(0, new HistoryEntry
                {
                    Timestamp = DateTime.Now, EventType = "Login", PersonName = _currentWindowsDisplayName,
                    AutodeskEmail = _currentAutodeskEmail, MachineId = _currentMachineName
                });
            }
            else if (string.IsNullOrEmpty(_currentAutodeskEmail) && !string.IsNullOrEmpty(previousEmail))
            {
                ActivityHistory.Insert(0, new HistoryEntry
                {
                    Timestamp = DateTime.Now, EventType = "Logout", PersonName = _currentWindowsDisplayName,
                    PreviousEmail = previousEmail, MachineId = _currentMachineName
                });
            }
            await RefreshAllAsync();
            await SyncToCloudAsync();
        });
    }

    private void UpdateLoginStatusDisplay()
    {
        IsAutodeskLoggedInDisplay = _isAutodeskLoggedIn;
        LoginStatusText = _isAutodeskLoggedIn ? "Logged In" : "Not Logged In";
        ConnectionStatusText = _isAutodeskLoggedIn ? "LOGGED IN" : "LOGGED OUT";
        CloudStatusText = CloudLoggingEnabled ? $"Cloud: ✓ ({Country})" : "Cloud: ✗";
    }

    private void LoadSettings()
    {
        var config = ConfigService.Instance.Config;
        Country = config.Country;
        Office = config.Office;
        CloudLoggingEnabled = config.CloudLoggingEnabled;
        CloudApiUrl = config.CloudApiUrl;
        CloudApiKey = config.CloudApiKey;
        NetworkLoggingEnabled = config.NetworkLoggingEnabled;
        NetworkLogPath = config.NetworkLogPath;
        CheckIntervalSeconds = config.MonitoringIntervalSeconds;
        StartMinimized = config.StartMinimized;
        AutoStartEnabled = config.AutoStartEnabled;
        CloudServerInfo = $"Cloud Server: {config.CloudApiUrl}";
    }

    private void InitializeDefaultProfiles()
    {
        UserProfiles.Clear();
        foreach (var (email, password) in DefaultUsers)
        {
            var nameParts = email.Split('@')[0].Split('.');
            var name = string.Join(" ", nameParts.Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));
            UserProfiles.Add(new UserProfile { Name = name, Email = email, Password = password, IsActive = false });
        }
    }

    partial void OnSelectedProfileChanged(UserProfile? value)
    {
        if (value != null) { EditName = value.Name; EditEmail = value.Email; EditPassword = value.Password; }
    }

    partial void OnSelectedSessionChanged(UserSession? value)
    {
        if (value != null) { EditSessionName = value.PersonName; }
    }

    private void AddDebug(string message) => 
        Application.Current.Dispatcher.Invoke(() => DebugLog += $"[{DateTime.Now:HH:mm:ss}] {message}\n");
}
