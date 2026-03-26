using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using AutodeskIDMonitor.Models;

namespace AutodeskIDMonitor.Services;

public class MonitorService : IDisposable
{
    // Win32 API for idle detection
    [DllImport("user32.dll")]
    static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
    
    // Win32 API for foreground window detection
    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [StructLayout(LayoutKind.Sequential)]
    struct LASTINPUTINFO
    {
        public uint cbSize;
        public uint dwTime;
    }

    private readonly List<string> _possiblePaths = new();
    private FileSystemWatcher? _watcher;
    private readonly System.Threading.Timer _timer;
    private readonly System.Threading.Timer _revitTimer;
    private AutodeskStatus _lastStatus = new();
    private string? _activeLoginStatePath;
    
    // Real-time tracking
    private readonly Dictionary<int, RevitProcessInfo> _activeRevitProcesses = new();
    private readonly Dictionary<string, MonitorProjectTimeEntry> _projectTimes = new();
    private readonly Dictionary<string, AppActivityEntry> _appActivities = new();
    private ActivityState _currentActivityState = ActivityState.Offline;
    private DateTime _lastUserInput = DateTime.Now;
    private DateTime _idleStartTime = DateTime.MinValue;
    private TimeSpan _totalIdleTime = TimeSpan.Zero;
    private bool _isCurrentlyIdle = false;
    
    // ========== DAY TRACKING - AUTO RESET AT MIDNIGHT ==========
    private DateTime _currentTrackingDay = DateTime.Today;
    
    // Active Revit tracking - only counts when Revit is foreground AND user is ACTIVELY interacting
    private bool _isRevitForeground = false;
    private DateTime _revitActiveStartTime = DateTime.MinValue;
    private TimeSpan _totalRevitActiveTime = TimeSpan.Zero;
    private string _currentForegroundProject = "";
    
    // Per-project active time tracking (only when Revit foreground + user active)
    private readonly Dictionary<string, TimeSpan> _projectActiveTime = new();
    private DateTime _lastProjectActiveUpdate = DateTime.MinValue;
    
    // Short idle threshold for Revit (30 seconds = stop counting)
    private const int REVIT_IDLE_THRESHOLD_SECONDS = 30;
    
    // ========== HOURLY ACTIVITY BUCKETS ==========
    // Track ACTUAL activity time per hour (only when actively working)
    private readonly Dictionary<int, double> _hourlyRevitMinutes = new();
    private readonly Dictionary<int, double> _hourlyMeetingMinutes = new();
    private readonly Dictionary<int, double> _hourlyIdleMinutes = new();
    private readonly Dictionary<int, double> _hourlyOtherMinutes = new();
    private DateTime _lastActivityUpdateTime = DateTime.Now; // Initialize to NOW for immediate tracking
    private string _lastActivityType = "Other"; // "Revit", "Meeting", "Idle", "Other"
    
    // Meeting apps to detect
    private static readonly HashSet<string> MeetingApps = new(StringComparer.OrdinalIgnoreCase)
    {
        // Microsoft Teams - all variations
        "Teams", "ms-teams", "MSTeams", "Microsoft Teams", "Microsoft Teams (work or school)",
        // Zoom
        "Zoom", "zoom", "ZoomIt",
        // Webex
        "webex", "CiscoWebex", "webexmta",
        // Other
        "Slack", "slack", "Discord", "discord", "Skype", "lync",
        "Google Meet", "meet.google.com", "BlueJeans", "GoToMeeting"
    };
    
    // Communication apps
    private static readonly HashSet<string> CommunicationApps = new(StringComparer.OrdinalIgnoreCase)
    {
        "Outlook", "OUTLOOK", "Thunderbird", "Mail", "Gmail"
    };

    public event EventHandler<AutodeskStatus>? StatusChanged;
    public event EventHandler<RevitActivityEventArgs>? RevitActivityChanged;
    public event EventHandler<ActivityStateEventArgs>? ActivityStateChanged;

    public MonitorService()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var roamingData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        
        // Add all possible paths for loginstate.json (LocalAppData)
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", ".autodesk", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", "Web Services", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", "identity", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(appData, ".autodesk", "loginstate.json"));
        
        // Additional paths for newer Autodesk versions
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", "Genuine Service", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", "webdeploy", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", "Autodesk Identity Manager", "loginstate.json"));
        
        // Roaming AppData paths
        _possiblePaths.Add(Path.Combine(roamingData, "Autodesk", ".autodesk", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(roamingData, "Autodesk", "Web Services", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(roamingData, "Autodesk", "identity", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(roamingData, ".autodesk", "loginstate.json"));
        
        // Newer Autodesk Identity Manager paths (Revit 2024/2025)
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", "Autodesk Identity Manager", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", "AIM", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", "ADSK", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", "Web Services", "LoginState.json"));
        _possiblePaths.Add(Path.Combine(appData, "Autodesk", "CLM", "loginstate.json"));

        // Also check under user profile directly
        var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        _possiblePaths.Add(Path.Combine(userProfile, ".autodesk", "loginstate.json"));
        _possiblePaths.Add(Path.Combine(userProfile, "AppData", "Local", "Autodesk", ".autodesk", "loginstate.json"));

        // Try to find ANY loginstate.json in Autodesk folders (recursive scan)
        try
        {
            var autodeskDir = Path.Combine(appData, "Autodesk");
            if (Directory.Exists(autodeskDir))
            {
                foreach (var file in Directory.GetFiles(autodeskDir, "loginstate.json", SearchOption.AllDirectories))
                {
                    if (!_possiblePaths.Contains(file))
                        _possiblePaths.Add(file);
                }
                // Also search for LoginState.json (capitalised variant)
                foreach (var file in Directory.GetFiles(autodeskDir, "LoginState.json", SearchOption.AllDirectories))
                {
                    if (!_possiblePaths.Contains(file))
                        _possiblePaths.Add(file);
                }
            }
        }
        catch { }
        
        // Find the first existing path
        foreach (var path in _possiblePaths)
        {
            if (File.Exists(path))
            {
                _activeLoginStatePath = path;
                break;
            }
            // Also check if directory exists for watcher
            var dir = Path.GetDirectoryName(path);
            if (Directory.Exists(dir) && _activeLoginStatePath == null)
            {
                _activeLoginStatePath = path;
            }
        }

        // Setup file watcher if we have a path
        if (_activeLoginStatePath != null)
        {
            var watchDir = Path.GetDirectoryName(_activeLoginStatePath);
            if (Directory.Exists(watchDir))
            {
                try
                {
                    _watcher = new FileSystemWatcher(watchDir)
                    {
                        Filter = "loginstate.json",
                        NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName
                    };
                    _watcher.Changed += OnFileChanged;
                    _watcher.Created += OnFileChanged;
                }
                catch { }
            }
        }

        // Autodesk login check timer (every 10 seconds)
        _timer = new System.Threading.Timer(CheckStatus, null, TimeSpan.Zero, TimeSpan.FromSeconds(10));
        
        // Revit process monitor timer (every 3 seconds for real-time tracking)
        _revitTimer = new System.Threading.Timer(CheckRevitProcesses, null, TimeSpan.Zero, TimeSpan.FromSeconds(3));
    }

    public void StartMonitoring()
    {
        if (_watcher != null)
        {
            try { _watcher.EnableRaisingEvents = true; } catch { }
        }
        CheckStatus(null);
        CheckRevitProcesses(null);
    }

    public void StopMonitoring()
    {
        if (_watcher != null)
        {
            try { _watcher.EnableRaisingEvents = false; } catch { }
        }
        
        // End all active project time entries
        foreach (var entry in _projectTimes.Values.Where(p => p.IsActive))
        {
            entry.EndTime = DateTime.Now;
            entry.IsActive = false;
        }
    }

    private void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        Thread.Sleep(500);
        CheckStatus(null);
    }

    // ===== REVIT PROCESS MONITORING =====
    
    private void CheckRevitProcesses(object? state)
    {
        try
        {
            var currentProcesses = GetRevitProcesses();
            var currentIds = currentProcesses.Select(p => p.ProcessId).ToHashSet();
            var previousIds = _activeRevitProcesses.Keys.ToHashSet();
            
            // Detect newly opened Revit instances
            var newProcesses = currentProcesses.Where(p => !previousIds.Contains(p.ProcessId)).ToList();
            foreach (var proc in newProcesses)
            {
                _activeRevitProcesses[proc.ProcessId] = proc;
                proc.StartTime = DateTime.Now;
                
                // Start tracking project time - use "Revit Session" as fallback if no project name
                var projectName = !string.IsNullOrEmpty(proc.ProjectName) ? proc.ProjectName : "Revit Session";
                StartProjectTime(projectName, proc.RevitVersion);
                
                RevitActivityChanged?.Invoke(this, new RevitActivityEventArgs
                {
                    EventType = RevitEventType.Opened,
                    Process = proc,
                    AllActiveProcesses = _activeRevitProcesses.Values.ToList()
                });
            }
            
            // Detect closed Revit instances
            var closedIds = previousIds.Except(currentIds).ToList();
            foreach (var id in closedIds)
            {
                if (_activeRevitProcesses.TryGetValue(id, out var closedProc))
                {
                    closedProc.EndTime = DateTime.Now;
                    
                    // End tracking project time - use "Revit Session" as fallback if no project name
                    var projectName = !string.IsNullOrEmpty(closedProc.ProjectName) ? closedProc.ProjectName : "Revit Session";
                    EndProjectTime(projectName);
                    
                    RevitActivityChanged?.Invoke(this, new RevitActivityEventArgs
                    {
                        EventType = RevitEventType.Closed,
                        Process = closedProc,
                        AllActiveProcesses = _activeRevitProcesses.Values.Where(p => currentIds.Contains(p.ProcessId)).ToList()
                    });
                    
                    _activeRevitProcesses.Remove(id);
                }
            }
            
            // Check for project changes in existing processes
            foreach (var proc in currentProcesses.Where(p => previousIds.Contains(p.ProcessId)))
            {
                if (_activeRevitProcesses.TryGetValue(proc.ProcessId, out var existing))
                {
                    // Always update window title from current state
                    existing.WindowTitle = proc.WindowTitle;
                    existing.RevitVersion = proc.RevitVersion;
                    existing.LastActivity = DateTime.Now;
                    
                    // Check if project name changed (including from empty to something)
                    var oldProject = existing.ProjectName ?? "";
                    var newProject = proc.ProjectName ?? "";
                    
                    if (oldProject != newProject && !string.IsNullOrEmpty(newProject))
                    {
                        // Project changed or was just detected
                        if (!string.IsNullOrEmpty(oldProject))
                        {
                            EndProjectTime(oldProject);
                        }
                        
                        existing.ProjectName = newProject;
                        StartProjectTime(newProject, proc.RevitVersion);
                        
                        RevitActivityChanged?.Invoke(this, new RevitActivityEventArgs
                        {
                            EventType = RevitEventType.ProjectChanged,
                            Process = existing,
                            AllActiveProcesses = _activeRevitProcesses.Values.ToList()
                        });
                    }
                }
            }
            
            // Check activity state
            UpdateActivityState();
        }
        catch { }
    }

    private List<RevitProcessInfo> GetRevitProcesses()
    {
        var result = new List<RevitProcessInfo>();
        
        try
        {
            // Only look for the main Revit.exe process, not workers
            var processes = Process.GetProcesses()
                .Where(p => p.ProcessName.Equals("Revit", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var proc in processes)
            {
                try
                {
                    var windowTitle = proc.MainWindowTitle;
                    
                    // Skip if no window title (minimized or no project open)
                    if (string.IsNullOrEmpty(windowTitle) || 
                        windowTitle.Equals("Autodesk Revit", StringComparison.OrdinalIgnoreCase))
                    {
                        // Try to get window title from all windows of this process
                        windowTitle = GetRevitWindowTitle(proc.Id);
                    }
                    
                    var projectName = ExtractProjectName(windowTitle);
                    
                    var info = new RevitProcessInfo
                    {
                        ProcessId = proc.Id,
                        ProcessName = proc.ProcessName,
                        WindowTitle = windowTitle,
                        RevitVersion = ExtractRevitVersion(proc),
                        ProjectName = projectName,
                        LastActivity = DateTime.Now
                    };
                    result.Add(info);
                }
                catch { }
            }
        }
        catch { }
        
        return result;
    }
    
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
    
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern int GetWindowTextLength(IntPtr hWnd);
    
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);
    
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    
    private string GetRevitWindowTitle(int processId)
    {
        string result = "";
        
        EnumWindows((hWnd, lParam) =>
        {
            GetWindowThreadProcessId(hWnd, out uint pid);
            if (pid == processId && IsWindowVisible(hWnd))
            {
                int length = GetWindowTextLength(hWnd);
                if (length > 0)
                {
                    var sb = new System.Text.StringBuilder(length + 1);
                    GetWindowText(hWnd, sb, sb.Capacity);
                    var title = sb.ToString();
                    
                    // Look for Revit window with project name
                    if (title.Contains(".rvt") || 
                        (title.Contains("Autodesk Revit") && title.Contains(" - ")))
                    {
                        result = title;
                        return false; // Stop enumeration
                    }
                    
                    // Fallback to any Revit window
                    if (title.Contains("Revit") && string.IsNullOrEmpty(result))
                    {
                        result = title;
                    }
                }
            }
            return true;
        }, IntPtr.Zero);
        
        return result;
    }

    private string ExtractRevitVersion(Process proc)
    {
        try
        {
            var path = proc.MainModule?.FileName ?? "";
            if (path.Contains("2025")) return "Revit 2025";
            if (path.Contains("2024")) return "Revit 2024";
            if (path.Contains("2023")) return "Revit 2023";
            if (path.Contains("2022")) return "Revit 2022";
            
            var title = proc.MainWindowTitle;
            if (title.Contains("2025")) return "Revit 2025";
            if (title.Contains("2024")) return "Revit 2024";
            if (title.Contains("2023")) return "Revit 2023";
            if (title.Contains("2022")) return "Revit 2022";
        }
        catch { }
        return "Revit";
    }

    private string ExtractProjectName(string windowTitle)
    {
        if (string.IsNullOrEmpty(windowTitle)) return "";
        
        try
        {
            // Skip if it's just the Revit startup screen
            if (windowTitle.Equals("Autodesk Revit", StringComparison.OrdinalIgnoreCase) ||
                (windowTitle.StartsWith("Autodesk Revit 20", StringComparison.OrdinalIgnoreCase) && !windowTitle.Contains("-")))
                return "";
            
            // Format: "Autodesk Revit 2024.3 - TV-9150409-TLA-MOD-VTC-ZZZ-LAN-0002.rvt - Sheet: 0000 - COVER SHEET"
            // We need to find the .rvt file name which can contain dashes
            
            // Method 1: Find .rvt and extract the filename (allowing dashes)
            // Look for pattern: " - FILENAME.rvt" where FILENAME can contain dashes, letters, numbers
            var rvtMatch = System.Text.RegularExpressions.Regex.Match(windowTitle, 
                @" - ([A-Za-z0-9_\-\.]+?)\.rvt", 
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (rvtMatch.Success)
            {
                var projectName = rvtMatch.Groups[1].Value.Trim();
                projectName = CleanProjectName(projectName);
                if (!string.IsNullOrWhiteSpace(projectName) && projectName.Length > 1)
                    return projectName;
            }
            
            // Method 2: Find any .rvt in the title (for format: "ProjectName.rvt - Autodesk Revit")
            var rvtMatch2 = System.Text.RegularExpressions.Regex.Match(windowTitle, 
                @"^([A-Za-z0-9_\-\.]+?)\.rvt", 
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (rvtMatch2.Success)
            {
                var projectName = rvtMatch2.Groups[1].Value.Trim();
                projectName = CleanProjectName(projectName);
                if (!string.IsNullOrWhiteSpace(projectName) && projectName.Length > 1)
                    return projectName;
            }
            
            // Method 3: Split by " - " and find the part with .rvt
            var parts = windowTitle.Split(new[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                if (part.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                {
                    var projectName = part.Substring(0, part.Length - 4).Trim();
                    projectName = CleanProjectName(projectName);
                    if (!string.IsNullOrWhiteSpace(projectName) && projectName.Length > 1)
                        return projectName;
                }
            }
            
            // Method 4: If Autodesk Revit is at the start, project might be second part
            if (parts.Length >= 2 && parts[0].Trim().StartsWith("Autodesk Revit", StringComparison.OrdinalIgnoreCase))
            {
                var projectPart = parts[1].Trim();
                
                // Remove file extensions
                if (projectPart.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                    projectPart = projectPart.Substring(0, projectPart.Length - 4);
                if (projectPart.EndsWith(".rfa", StringComparison.OrdinalIgnoreCase))
                    projectPart = projectPart.Substring(0, projectPart.Length - 4);
                
                projectPart = CleanProjectName(projectPart);
                if (!string.IsNullOrWhiteSpace(projectPart) && projectPart.Length > 1 &&
                    !projectPart.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase))
                    return projectPart;
            }
        }
        catch { }
        
        return "";
    }
    
    private string CleanProjectName(string name)
    {
        if (string.IsNullOrEmpty(name)) return "";
        
        // Remove common suffixes
        if (name.EndsWith("_detached", StringComparison.OrdinalIgnoreCase))
            name = name.Substring(0, name.Length - 9);
        if (name.EndsWith("_central", StringComparison.OrdinalIgnoreCase))
            name = name.Substring(0, name.Length - 8);
        if (name.EndsWith("_local", StringComparison.OrdinalIgnoreCase))
            name = name.Substring(0, name.Length - 6);
        if (name.EndsWith(" (edited)", StringComparison.OrdinalIgnoreCase))
            name = name.Substring(0, name.Length - 9);
        
        // Remove leading/trailing whitespace and special chars
        name = name.Trim(' ', '-', '_', '[', ']', '(', ')');
        
        // Filter out non-project names (Revit UI elements, generic names)
        var invalidNames = new[] { 
            "Home", "Start", "Revit", "Recent", "New", "Open", 
            "Models", "Families", "Templates", "Samples",
            "Autodesk", "Architecture", "Structure", "MEP"
        };
        
        foreach (var invalid in invalidNames)
        {
            if (name.Equals(invalid, StringComparison.OrdinalIgnoreCase))
                return "";
        }
        
        return name;
    }

    private void StartProjectTime(string projectName, string revitVersion)
    {
        var key = projectName.ToLower();
        var today = DateTime.Today;
        
        if (!_projectTimes.ContainsKey(key))
        {
            _projectTimes[key] = new MonitorProjectTimeEntry
            {
                ProjectName = projectName,
                RevitVersion = revitVersion,
                StartTime = DateTime.Now,
                TodaySessionStart = DateTime.Now,
                TodayDuration = TimeSpan.Zero,
                IsActive = true
            };
        }
        else
        {
            var entry = _projectTimes[key];
            
            // Check if we need to reset for new day
            if (entry.TodaySessionStart.Date < today)
            {
                entry.ResetForNewDay();
            }
            
            if (!entry.IsActive)
            {
                // Restarting work on this project
                entry.StartTime = DateTime.Now;
                entry.TodaySessionStart = DateTime.Now;
                entry.IsActive = true;
            }
        }
    }

    private void EndProjectTime(string projectName)
    {
        var key = projectName.ToLower();
        if (_projectTimes.TryGetValue(key, out var entry) && entry.IsActive)
        {
            var now = DateTime.Now;
            entry.EndTime = now;
            
            // Calculate session duration
            var sessionDuration = now - entry.StartTime;
            entry.TotalDuration += sessionDuration;
            
            // Add to TODAY's duration (only if session started today)
            if (entry.TodaySessionStart.Date == DateTime.Today)
            {
                entry.TodayDuration += now - entry.TodaySessionStart;
            }
            
            entry.IsActive = false;
        }
    }

    // ===== ACTIVITY STATE =====
    
    private void UpdateActivityState()
    {
        var idleSeconds = GetIdleTimeSeconds();
        var hasRevitOpen = _activeRevitProcesses.Any();
        
        ActivityState newState;
        if (!hasRevitOpen)
            newState = ActivityState.Offline;
        else if (idleSeconds > 300) // 5 minutes idle
            newState = ActivityState.Idle;
        else
            newState = ActivityState.Active;
        
        if (newState != _currentActivityState)
        {
            var previousState = _currentActivityState;
            _currentActivityState = newState;
            
            ActivityStateChanged?.Invoke(this, new ActivityStateEventArgs
            {
                PreviousState = previousState,
                CurrentState = newState,
                IdleSeconds = idleSeconds,
                ActiveRevitCount = _activeRevitProcesses.Count
            });
        }
    }

    private int GetIdleTimeSeconds()
    {
        try
        {
            var lastInputInfo = new LASTINPUTINFO();
            lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
            
            if (GetLastInputInfo(ref lastInputInfo))
            {
                uint idleTime = (uint)Environment.TickCount - lastInputInfo.dwTime;
                return (int)(idleTime / 1000);
            }
        }
        catch { }
        return 0;
    }

    // ===== PUBLIC ACCESSORS =====
    
    public List<RevitProcessInfo> GetActiveRevitProcesses() => _activeRevitProcesses.Values.ToList();
    
    /// <summary>
    /// Get the currently active (foreground) Revit project name
    /// </summary>
    public string GetCurrentForegroundProject() => _isRevitForeground ? _currentForegroundProject : "";
    
    /// <summary>
    /// Check if Revit is currently the foreground application
    /// </summary>
    public bool IsRevitInForeground() => _isRevitForeground;
    
    /// <summary>
    /// Get current activity state string (for UI display)
    /// </summary>
    public string GetActivityStateString()
    {
        var idleSeconds = GetIdleTimeSeconds();
        if (idleSeconds > 300) return "Idle";
        if (_wasInMeeting && _meetingGraceCounter > 0) return "InMeeting";
        if (_isRevitForeground) return "Active";
        if (_activeRevitProcesses.Count > 0) return "Active";
        return "Offline";
    }
    
    public Dictionary<string, MonitorProjectTimeEntry> GetProjectTimes() => new(_projectTimes);
    
    public ActivityState GetCurrentActivityState() => _currentActivityState;
    
    public int GetIdleSeconds() => GetIdleTimeSeconds();
    
    public RealTimeStats GetRealTimeStats()
    {
        return new RealTimeStats
        {
            ActivityState = _currentActivityState,
            IdleSeconds = GetIdleTimeSeconds(),
            ActiveRevitCount = _activeRevitProcesses.Count,
            ActiveProjects = _activeRevitProcesses.Values
                .Where(p => !string.IsNullOrEmpty(p.ProjectName))
                .Select(p => p.ProjectName)
                .Distinct()
                .ToList(),
            ProjectDurations = _projectTimes.Values
                .Select(p => new ProjectDuration
                {
                    ProjectName = p.ProjectName,
                    RevitVersion = p.RevitVersion,
                    StartTime = p.StartTime,
                    Duration = p.IsActive 
                        ? p.TotalDuration + (DateTime.Now - p.StartTime)
                        : p.TotalDuration,
                    IsActive = p.IsActive
                })
                .ToList()
        };
    }
    
    // ===== MEETING/APP DETECTION =====
    
    // Store meeting start times separately (key = appName, value = meeting start time)
    private readonly Dictionary<string, DateTime> _meetingStartTimes = new();
    
    // Store accumulated meeting time for today (ONLY from actual verified meetings)
    private TimeSpan _totalMeetingTime = TimeSpan.Zero;
    private DateTime _lastMeetingUpdateTime = DateTime.MinValue;
    private bool _wasInMeeting = false;
    private string _currentMeetingApp = "";
    private int _meetingGraceCounter = 0; // Prevent detection flicker (NOT for time accumulation)
    
    /// <summary>
    /// Detect running meeting applications (Teams, Zoom, etc.)
    /// STRICT: Only returns IsInActiveMeeting=true when actually in a call
    /// </summary>
    public List<MeetingAppInfo> GetActiveMeetings()
    {
        var meetings = new List<MeetingAppInfo>();
        var now = DateTime.Now;
        bool foundActiveMeeting = false;
        
        try
        {
            var processes = Process.GetProcesses();
            var teamsWindowTitles = new List<string>();
            
            // First, collect ALL window titles from Teams processes
            foreach (var proc in processes)
            {
                try
                {
                    var name = proc.ProcessName;
                    if (name.Contains("Teams", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("ms-teams", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("MSTeams", StringComparison.OrdinalIgnoreCase))
                    {
                        var title = "";
                        try { title = proc.MainWindowTitle; } catch { }
                        if (!string.IsNullOrEmpty(title))
                        {
                            teamsWindowTitles.Add(title);
                        }
                    }
                }
                catch { }
            }
            
            foreach (var proc in processes)
            {
                try
                {
                    var name = proc.ProcessName;
                    
                    // Check if this is a meeting app
                    if (IsMeetingApp(name))
                    {
                        // Try to get window title for more details
                        string windowTitle = "";
                        try { windowTitle = proc.MainWindowTitle; } catch { }
                        
                        // For Teams, check if ANY window shows meeting indicators
                        bool isTeamsApp = name.Contains("Teams", StringComparison.OrdinalIgnoreCase) ||
                                          name.Contains("ms-teams", StringComparison.OrdinalIgnoreCase) ||
                                          name.Contains("MSTeams", StringComparison.OrdinalIgnoreCase);
                        
                        bool isInMeeting = false;
                        
                        if (isTeamsApp && teamsWindowTitles.Count > 0)
                        {
                            // Check ALL Teams window titles for meeting indicators
                            foreach (var title in teamsWindowTitles)
                            {
                                if (IsActivelyInMeeting(name, title))
                                {
                                    isInMeeting = true;
                                    windowTitle = title; // Use the matching title
                                    break;
                                }
                            }
                        }
                        else
                        {
                            // For non-Teams apps, use normal check
                            isInMeeting = IsActivelyInMeeting(name, windowTitle);
                        }
                        
                        var friendlyName = GetFriendlyAppName(name);
                        
                        if (isInMeeting)
                        {
                            foundActiveMeeting = true;
                            _meetingGraceCounter = 10; // Grace period: 10 cycles (~10 seconds) to handle brief title flicker
                            
                            // Track meeting start time - only set once when meeting starts
                            if (!_meetingStartTimes.ContainsKey(friendlyName))
                            {
                                _meetingStartTimes[friendlyName] = now;
                                _lastMeetingUpdateTime = now;
                            }
                            
                            // Accumulate meeting time ONLY if verified in meeting
                            if (_wasInMeeting && _currentMeetingApp == friendlyName && _lastMeetingUpdateTime > DateTime.MinValue)
                            {
                                var elapsed = now - _lastMeetingUpdateTime;
                                // Sanity check: only add if reasonable time elapsed (0.5 sec to 5 min)
                                if (elapsed.TotalSeconds >= 0.5 && elapsed.TotalMinutes < 5)
                                {
                                    _totalMeetingTime += elapsed;
                                }
                            }
                            
                            _lastMeetingUpdateTime = now;
                            _wasInMeeting = true;
                            _currentMeetingApp = friendlyName;
                            
                            meetings.Add(new MeetingAppInfo
                            {
                                AppName = friendlyName,
                                ProcessName = name,
                                WindowTitle = windowTitle,
                                IsInActiveMeeting = true,
                                StartTime = _meetingStartTimes[friendlyName]
                            });
                        }
                        else
                        {
                            // Not detected in meeting
                            // Check if window shows a CLEAR non-meeting activity
                            bool clearlyNotInMeeting = 
                                windowTitle.Contains("Chat", StringComparison.OrdinalIgnoreCase) ||
                                windowTitle.Contains("Calendar", StringComparison.OrdinalIgnoreCase) ||
                                windowTitle.Contains("Files", StringComparison.OrdinalIgnoreCase) ||
                                windowTitle.Contains("Activity", StringComparison.OrdinalIgnoreCase) ||
                                windowTitle.Contains("Teams |", StringComparison.OrdinalIgnoreCase) ||
                                windowTitle.Contains("Calls |", StringComparison.OrdinalIgnoreCase);
                            
                            // Only use grace period if window doesn't clearly show non-meeting
                            // If user switched to Chat/Calendar/etc, immediately end meeting detection
                            bool stillInGracePeriod = !clearlyNotInMeeting && 
                                                       _meetingGraceCounter > 0 && 
                                                       _currentMeetingApp == friendlyName;
                            
                            if (stillInGracePeriod)
                            {
                                // During grace period - keep detection active but DON'T accumulate time
                                foundActiveMeeting = true;
                                
                                meetings.Add(new MeetingAppInfo
                                {
                                    AppName = friendlyName,
                                    ProcessName = name,
                                    WindowTitle = windowTitle,
                                    IsInActiveMeeting = true, // Detection still active
                                    StartTime = _meetingStartTimes.GetValueOrDefault(friendlyName, now)
                                });
                                
                                // DON'T accumulate time during grace period
                            }
                            else
                            {
                                // Grace period expired OR user clearly not in meeting - truly end meeting
                                if (_meetingStartTimes.ContainsKey(friendlyName))
                                {
                                    _meetingStartTimes.Remove(friendlyName);
                                }
                                if (_currentMeetingApp == friendlyName)
                                {
                                    _wasInMeeting = false;
                                    _currentMeetingApp = "";
                                    _lastMeetingUpdateTime = DateTime.MinValue;
                                    _meetingGraceCounter = 0; // Reset grace counter immediately
                                }
                                
                                // Report as NOT in meeting
                                meetings.Add(new MeetingAppInfo
                                {
                                    AppName = friendlyName,
                                    ProcessName = name,
                                    WindowTitle = windowTitle,
                                    IsInActiveMeeting = false,
                                    StartTime = now
                                });
                            }
                        }
                    }
                }
                catch { }
            }
        }
        catch { }
        
        // Decrement grace counter if no active meeting found this cycle
        if (!foundActiveMeeting)
        {
            if (_meetingGraceCounter > 0)
            {
                _meetingGraceCounter--;
            }
            
            if (_meetingGraceCounter == 0 && _wasInMeeting)
            {
                // Grace period fully expired - end all meeting tracking
                _wasInMeeting = false;
                _currentMeetingApp = "";
                _lastMeetingUpdateTime = DateTime.MinValue;
                _meetingStartTimes.Clear(); // Clear all meeting start times
            }
        }
        
        return meetings.DistinctBy(m => m.AppName).ToList();
    }
    
    /// <summary>
    /// Get total accumulated meeting time for today (ACCURATE - only verified meeting time)
    /// </summary>
    public TimeSpan GetTodayMeetingTime()
    {
        // Return the accumulated meeting time (from verified meetings only)
        // If currently in a meeting, add the current session time
        if (_wasInMeeting && _lastMeetingUpdateTime > DateTime.MinValue)
        {
            // Add current meeting session to the total
            var currentSessionTime = DateTime.Now - _lastMeetingUpdateTime;
            // Sanity check
            if (currentSessionTime.TotalMinutes < 5)
            {
                return _totalMeetingTime + currentSessionTime;
            }
        }
        return _totalMeetingTime;
    }
    
    /// <summary>
    /// Get current active meeting duration (real-time display for current meeting only)
    /// </summary>
    public TimeSpan GetCurrentMeetingDuration()
    {
        if (_wasInMeeting && !string.IsNullOrEmpty(_currentMeetingApp) && _meetingStartTimes.TryGetValue(_currentMeetingApp, out var startTime))
        {
            return DateTime.Now - startTime;
        }
        return TimeSpan.Zero;
    }
    
    /// <summary>
    /// Reset daily counters (call at midnight or day change)
    /// </summary>
    public void ResetDailyCounters()
    {
        _totalMeetingTime = TimeSpan.Zero;
        _totalIdleTime = TimeSpan.Zero;
        _totalRevitActiveTime = TimeSpan.Zero;
        _meetingStartTimes.Clear();
        _wasInMeeting = false;
        _lastMeetingUpdateTime = DateTime.MinValue;
        _currentMeetingApp = "";
        _meetingGraceCounter = 0;
        _projectActiveTime.Clear();
        _currentForegroundProject = "";
        _isRevitForeground = false;
        
        // Clear hourly activity buckets
        _hourlyRevitMinutes.Clear();
        _hourlyMeetingMinutes.Clear();
        _hourlyIdleMinutes.Clear();
        _hourlyOtherMinutes.Clear();
        _lastActivityUpdateTime = DateTime.MinValue;
        _lastActivityType = "Other";
    }
    
    /// <summary>
    /// Reset all daily tracking for new day (called automatically at midnight)
    /// </summary>
    private void ResetForNewDay()
    {
        // Clear all daily counters
        ResetDailyCounters();
        
        // Reset idle tracking
        _isCurrentlyIdle = false;
        _idleStartTime = DateTime.MinValue;
        
        // Reset project times for new day
        foreach (var project in _projectTimes.Values)
        {
            project.ResetForNewDay();
        }
        
        // Clear app activities
        _appActivities.Clear();
        
        // Update tracking day
        _currentTrackingDay = DateTime.Today;
    }
    
    private bool IsMeetingApp(string processName)
    {
        return MeetingApps.Any(app => processName.Contains(app, StringComparison.OrdinalIgnoreCase));
    }
    
    private bool IsActivelyInMeeting(string processName, string windowTitle)
    {
        // ULTRA STRICT detection - ONLY return true if DEFINITELY in a call/meeting
        // Default is FALSE - must prove we're in a meeting
        if (string.IsNullOrEmpty(windowTitle)) return false;
        
        // Normalize window title
        var title = windowTitle.Trim();
        
        // ===== MICROSOFT TEAMS DETECTION (ULTRA STRICT) =====
        if (processName.Contains("Teams", StringComparison.OrdinalIgnoreCase) || 
            processName.Contains("ms-teams", StringComparison.OrdinalIgnoreCase) ||
            processName.Contains("MSTeams", StringComparison.OrdinalIgnoreCase))
        {
            // ========== ABSOLUTE EXCLUSIONS ==========
            // If ANY of these patterns are found, DEFINITELY NOT in a meeting
            
            // Pattern: "Microsoft Teams - Chat | PersonName" or "Chat | Microsoft Teams"
            // This is the EXACT issue - Chat windows being detected as meetings
            if (title.Contains("Chat", StringComparison.OrdinalIgnoreCase))
                return false;
            
            if (title.Contains("Chats", StringComparison.OrdinalIgnoreCase))
                return false;
            
            // Any navigation/app section
            if (title.Contains("Calendar", StringComparison.OrdinalIgnoreCase))
                return false;
            if (title.Contains("Activity", StringComparison.OrdinalIgnoreCase))
                return false;
            if (title.Contains("Files", StringComparison.OrdinalIgnoreCase))
                return false;
            if (title.Contains("Apps", StringComparison.OrdinalIgnoreCase))
                return false;
            if (title.Contains("People", StringComparison.OrdinalIgnoreCase))
                return false;
            if (title.Contains("Search", StringComparison.OrdinalIgnoreCase))
                return false;
            if (title.Contains("Assignments", StringComparison.OrdinalIgnoreCase))
                return false;
            if (title.Contains("Calls |", StringComparison.OrdinalIgnoreCase)) // Calls list, not in a call
                return false;
            if (title.Contains("Teams |", StringComparison.OrdinalIgnoreCase))
                return false;
            
            // Just the app name
            if (title.Equals("Microsoft Teams", StringComparison.OrdinalIgnoreCase))
                return false;
            if (title.Equals("Microsoft Teams (work or school)", StringComparison.OrdinalIgnoreCase))
                return false;
            if (title.StartsWith("Microsoft Teams -", StringComparison.OrdinalIgnoreCase) && 
                !title.Contains("Call", StringComparison.OrdinalIgnoreCase) &&
                !title.Contains("Meeting", StringComparison.OrdinalIgnoreCase))
                return false;
            
            // ========== POSITIVE CALL INDICATORS ==========
            // ONLY these specific patterns indicate an ACTUAL call
            
            // Pattern 1: Explicit "Call with" text
            if (title.Contains("Call with", StringComparison.OrdinalIgnoreCase))
                return true;
            
            // Pattern 2: "In a call" or "In call" status (usually in window title during call)
            if (title.Contains("In a call", StringComparison.OrdinalIgnoreCase))
                return true;
            if (title.Contains("In call", StringComparison.OrdinalIgnoreCase))
                return true;
            
            // Pattern 3: "Calling..." or "Calling" - initiating a call
            if (title.Contains("Calling", StringComparison.OrdinalIgnoreCase))
                return true;
            
            // Pattern 4: "Ringing" - receiving a call
            if (title.Contains("Ringing", StringComparison.OrdinalIgnoreCase))
                return true;
            
            // Pattern 5: Call duration showing (like "01:41" or "5:30" or "1:23:45")
            // Teams shows call duration in window title during active calls
            if (System.Text.RegularExpressions.Regex.IsMatch(title, @"^\d{1,2}:\d{2}(:\d{2})?$"))
                return true;
            
            // Pattern 6: Screen sharing/presenting (only during actual calls)
            if (title.Contains("Screen sharing", StringComparison.OrdinalIgnoreCase))
                return true;
            if (title.Contains("You're presenting", StringComparison.OrdinalIgnoreCase))
                return true;
            if (title.Contains("Presenting", StringComparison.OrdinalIgnoreCase) && 
                title.Contains("Microsoft Teams", StringComparison.OrdinalIgnoreCase))
                return true;
            
            // Pattern 7: Meeting window with specific format
            // "Meeting Title | Microsoft Teams" during active meeting
            // BUT NOT calendar entries or meeting chats
            if (title.EndsWith("| Microsoft Teams", StringComparison.OrdinalIgnoreCase))
            {
                var beforePipe = title.Replace("| Microsoft Teams", "").Trim();
                // Must contain "Meeting" or "Call" and NOT contain exclusions
                if ((beforePipe.Contains("Meeting", StringComparison.OrdinalIgnoreCase) ||
                     beforePipe.Contains("Call", StringComparison.OrdinalIgnoreCase)) &&
                    !beforePipe.Contains("Chat", StringComparison.OrdinalIgnoreCase) &&
                    !beforePipe.Contains("Calendar", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            
            // Pattern 8: Connected status (only during actual calls)
            if (title.Contains("Connected", StringComparison.OrdinalIgnoreCase))
                return true;
            
            // Pattern 9: Compact call window
            if (title.Contains("compact", StringComparison.OrdinalIgnoreCase) && 
                title.Contains("call", StringComparison.OrdinalIgnoreCase))
                return true;
            
            // Pattern 10: "On hold" status
            if (title.Contains("On hold", StringComparison.OrdinalIgnoreCase))
                return true;
            
            // Pattern 11: "Reconnecting" status
            if (title.Contains("Reconnecting", StringComparison.OrdinalIgnoreCase))
                return true;
            
            // DEFAULT: NOT in a meeting
            // If we didn't find a positive call indicator, assume NOT in meeting
            return false;
        }
        
        // ===== ZOOM DETECTION =====
        if (processName.Contains("Zoom", StringComparison.OrdinalIgnoreCase))
        {
            // Zoom meeting patterns
            var zoomMeetingPatterns = new[] {
                "Zoom Meeting",
                "Meeting in progress",
                "Screen sharing",
                "You are in a meeting"
            };
            
            // Zoom NOT in meeting
            if (title.Equals("Zoom", StringComparison.OrdinalIgnoreCase) ||
                title.Contains("Home", StringComparison.OrdinalIgnoreCase) ||
                title.Contains("Sign In", StringComparison.OrdinalIgnoreCase))
                return false;
            
            return zoomMeetingPatterns.Any(p => title.Contains(p, StringComparison.OrdinalIgnoreCase));
        }
        
        // ===== WEBEX DETECTION =====
        if (processName.Contains("webex", StringComparison.OrdinalIgnoreCase))
        {
            if (title.Contains("Meeting", StringComparison.OrdinalIgnoreCase) ||
                title.Contains("Call", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }
        
        // ===== SLACK HUDDLE DETECTION =====
        if (processName.Contains("Slack", StringComparison.OrdinalIgnoreCase))
        {
            // Slack huddles have specific patterns
            if (title.Contains("Huddle", StringComparison.OrdinalIgnoreCase) ||
                title.Contains("Call", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }
        
        // ===== GENERIC DETECTION FOR OTHER APPS =====
        // Very strict - require explicit call/meeting indicators
        var genericCallPatterns = new[] {
            "Meeting in progress",
            "In a call",
            "In call",
            "Video call",
            "Screen sharing",
            "Presenting"
        };
        
        return genericCallPatterns.Any(p => title.Contains(p, StringComparison.OrdinalIgnoreCase));
    }
    
    private string GetFriendlyAppName(string processName)
    {
        if (processName.Contains("Teams", StringComparison.OrdinalIgnoreCase)) return "Microsoft Teams";
        if (processName.Contains("Zoom", StringComparison.OrdinalIgnoreCase)) return "Zoom";
        if (processName.Contains("webex", StringComparison.OrdinalIgnoreCase)) return "Cisco Webex";
        if (processName.Contains("Slack", StringComparison.OrdinalIgnoreCase)) return "Slack";
        if (processName.Contains("Discord", StringComparison.OrdinalIgnoreCase)) return "Discord";
        if (processName.Contains("Skype", StringComparison.OrdinalIgnoreCase)) return "Skype";
        return processName;
    }
    
    private DateTime GetProcessStartTime(Process proc)
    {
        try { return proc.StartTime; }
        catch { return DateTime.Now; }
    }
    
    /// <summary>
    /// Track application activity - STRICT tracking for Revit and Meetings
    /// Uses hourly buckets for accurate, mutually exclusive time tracking
    /// </summary>
    public void TrackAppActivity()
    {
        try
        {
            // Use Dubai time (UTC+4) as the canonical timezone for all TLA offices
            var now = Services.UserCredentialService.NowDubai;
            var today = now.Date;
            var currentHour = now.Hour;
            var idleSeconds = GetIdleTimeSeconds();
            
            // ========== DAY CHANGE DETECTION - AUTO RESET AT MIDNIGHT (DXB) ==========
            if (today > _currentTrackingDay)
            {
                // New day! Reset all daily counters
                ResetForNewDay();
                _currentTrackingDay = today;
            }
            
            // Initialize hourly bucket if needed
            if (!_hourlyRevitMinutes.ContainsKey(currentHour))
            {
                _hourlyRevitMinutes[currentHour] = 0;
                _hourlyMeetingMinutes[currentHour] = 0;
                _hourlyIdleMinutes[currentHour] = 0;
                _hourlyOtherMinutes[currentHour] = 0;
            }
            
            // Calculate time elapsed since last update (max 2 minutes to prevent huge jumps)
            double elapsedMinutes = 0;
            var timeSinceLastUpdate = now - _lastActivityUpdateTime;
            
            // Only track if at least 0.5 seconds have passed and same day
            if (timeSinceLastUpdate.TotalSeconds >= 0.5 && _lastActivityUpdateTime.Date == today)
            {
                elapsedMinutes = Math.Min(2.0, timeSinceLastUpdate.TotalMinutes);
            }
            else if (_lastActivityUpdateTime.Date != today)
            {
                // New day - reset and start fresh
                _lastActivityUpdateTime = now;
            }
            _lastActivityUpdateTime = now;
            
            // ========== DETERMINE CURRENT ACTIVITY TYPE (MUTUALLY EXCLUSIVE) ==========
            // Priority: Idle > Meeting > Revit > Other
            
            string currentActivity = "Other";
            
            // Check 1: Is user IDLE? (no input for 5 minutes)
            bool isUserIdle = idleSeconds >= 300;
            
            // Check 2: Is user in a REAL meeting? (actual call, not just Teams open)
            var meetings = GetActiveMeetings();
            bool isInRealMeeting = meetings.Any(m => m.IsInActiveMeeting);
            
            // Check 3: Is Revit foreground AND user recently active?
            bool isRevitForeground = IsRevitForegroundWindow();
            bool isUserActiveForRevit = idleSeconds < REVIT_IDLE_THRESHOLD_SECONDS;
            bool isActivelyUsingRevit = isRevitForeground && isUserActiveForRevit && !isUserIdle;
            
            // Determine activity (mutually exclusive - only ONE can be true)
            if (isUserIdle)
            {
                currentActivity = "Idle";
            }
            else if (isInRealMeeting)
            {
                currentActivity = "Meeting";
            }
            else if (isActivelyUsingRevit)
            {
                currentActivity = "Revit";
            }
            else
            {
                currentActivity = "Other";
            }
            
            // ========== UPDATE HOURLY BUCKETS ==========
            // Add elapsed time to the appropriate bucket
            if (elapsedMinutes > 0)
            {
                switch (currentActivity)
                {
                    case "Revit":
                        _hourlyRevitMinutes[currentHour] += elapsedMinutes;
                        break;
                    case "Meeting":
                        _hourlyMeetingMinutes[currentHour] += elapsedMinutes;
                        break;
                    case "Idle":
                        _hourlyIdleMinutes[currentHour] += elapsedMinutes;
                        break;
                    case "Other":
                        _hourlyOtherMinutes[currentHour] += elapsedMinutes;
                        break;
                }
                
                // Cap each bucket at 60 minutes max
                _hourlyRevitMinutes[currentHour] = Math.Min(60, _hourlyRevitMinutes[currentHour]);
                _hourlyMeetingMinutes[currentHour] = Math.Min(60, _hourlyMeetingMinutes[currentHour]);
                _hourlyIdleMinutes[currentHour] = Math.Min(60, _hourlyIdleMinutes[currentHour]);
                _hourlyOtherMinutes[currentHour] = Math.Min(60, _hourlyOtherMinutes[currentHour]);
            }
            
            _lastActivityType = currentActivity;
            
            // ========== UPDATE LEGACY TRACKING (for backward compatibility) ==========
            
            // Track overall idle time (5 minute threshold)
            if (idleSeconds > 300)
            {
                if (!_isCurrentlyIdle)
                {
                    _isCurrentlyIdle = true;
                    _idleStartTime = now.AddSeconds(-idleSeconds);
                    
                    // Stop Revit active time when user goes idle
                    if (_isRevitForeground)
                    {
                        var elapsed = now - _revitActiveStartTime;
                        _totalRevitActiveTime += elapsed;
                        
                        // Also update project-specific active time
                        if (!string.IsNullOrEmpty(_currentForegroundProject))
                        {
                            if (!_projectActiveTime.ContainsKey(_currentForegroundProject))
                                _projectActiveTime[_currentForegroundProject] = TimeSpan.Zero;
                            _projectActiveTime[_currentForegroundProject] += elapsed;
                        }
                        
                        _isRevitForeground = false;
                    }
                }
            }
            else
            {
                if (_isCurrentlyIdle)
                {
                    // Was idle, now active - record idle period
                    _totalIdleTime += now - _idleStartTime;
                    _isCurrentlyIdle = false;
                }
            }
            
            // Track Revit foreground for legacy systems
            var currentProject = GetCurrentForegroundRevitProject();
            
            if (isActivelyUsingRevit)
            {
                if (!_isRevitForeground)
                {
                    _isRevitForeground = true;
                    _revitActiveStartTime = now;
                    _currentForegroundProject = currentProject;
                    _lastProjectActiveUpdate = now;
                }
                else
                {
                    // Check if project changed
                    if (!string.IsNullOrEmpty(currentProject) && currentProject != _currentForegroundProject)
                    {
                        var elapsed = now - _lastProjectActiveUpdate;
                        if (!string.IsNullOrEmpty(_currentForegroundProject))
                        {
                            if (!_projectActiveTime.ContainsKey(_currentForegroundProject))
                                _projectActiveTime[_currentForegroundProject] = TimeSpan.Zero;
                            _projectActiveTime[_currentForegroundProject] += elapsed;
                        }
                        
                        _currentForegroundProject = currentProject;
                        _lastProjectActiveUpdate = now;
                    }
                }
            }
            else
            {
                if (_isRevitForeground)
                {
                    var elapsed = now - _revitActiveStartTime;
                    _totalRevitActiveTime += elapsed;
                    
                    if (!string.IsNullOrEmpty(_currentForegroundProject))
                    {
                        var projectElapsed = now - _lastProjectActiveUpdate;
                        if (!_projectActiveTime.ContainsKey(_currentForegroundProject))
                            _projectActiveTime[_currentForegroundProject] = TimeSpan.Zero;
                        _projectActiveTime[_currentForegroundProject] += projectElapsed;
                    }
                    
                    _isRevitForeground = false;
                    _currentForegroundProject = "";
                }
            }
            
            // Update meeting app activities
            foreach (var meeting in meetings.Where(m => m.IsInActiveMeeting))
            {
                var key = $"meeting_{meeting.AppName}".ToLower();
                
                if (!_appActivities.ContainsKey(key) || !_appActivities[key].IsActive)
                {
                    _appActivities[key] = new AppActivityEntry
                    {
                        AppName = meeting.AppName,
                        ActivityType = "Meeting",
                        StartTime = meeting.StartTime,
                        IsActive = true
                    };
                }
            }
            
            // End meeting tracking for apps no longer in actual meeting
            var activeMeetingApps = meetings.Where(m => m.IsInActiveMeeting).Select(m => m.AppName).ToHashSet();
            foreach (var entry in _appActivities.Values.Where(a => a.ActivityType == "Meeting" && a.IsActive))
            {
                if (!activeMeetingApps.Contains(entry.AppName))
                {
                    entry.EndTime = now;
                    entry.IsActive = false;
                }
            }
        }
        catch { }
    }
    
    /// <summary>
    /// Get the project name from the currently foreground Revit window
    /// </summary>
    private string GetCurrentForegroundRevitProject()
    {
        try
        {
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) return "";
            
            GetWindowThreadProcessId(foregroundWindow, out uint processId);
            
            if (_activeRevitProcesses.TryGetValue((int)processId, out var revitInfo))
            {
                return revitInfo.ProjectName ?? "";
            }
        }
        catch { }
        return "";
    }
    
    /// <summary>
    /// Check if Revit is the foreground (active) window
    /// </summary>
    public bool IsRevitForegroundWindow()
    {
        try
        {
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) return false;
            
            GetWindowThreadProcessId(foregroundWindow, out uint processId);
            
            // Check if the foreground process is one of our tracked Revit processes
            if (_activeRevitProcesses.ContainsKey((int)processId))
            {
                // Get the window title to extract project name
                var windowTitle = GetWindowTitle(foregroundWindow);
                if (!string.IsNullOrEmpty(windowTitle))
                {
                    // Extract project name from Revit window title
                    // Format: "Autodesk Revit 2025.4 - PROJECT-NAME.rvt - View: ..."
                    var extractedProject = ExtractProjectFromRevitTitle(windowTitle);
                    if (!string.IsNullOrEmpty(extractedProject) && extractedProject != _currentForegroundProject)
                    {
                        _currentForegroundProject = extractedProject;
                    }
                }
                return true;
            }
            
            // Fallback: Check window title directly for Revit patterns
            // This handles cases where process ID tracking might be out of sync
            var title = GetWindowTitle(foregroundWindow);
            if (!string.IsNullOrEmpty(title) && 
                (title.Contains("Autodesk Revit", StringComparison.OrdinalIgnoreCase) ||
                 title.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase) ||
                 title.Contains(".rvt -", StringComparison.OrdinalIgnoreCase)))
            {
                // This is a Revit window - extract project and return true
                var extractedProject = ExtractProjectFromRevitTitle(title);
                if (!string.IsNullOrEmpty(extractedProject) && extractedProject != _currentForegroundProject)
                {
                    _currentForegroundProject = extractedProject;
                }
                return true;
            }
            
            return false;
        }
        catch
        {
            return false;
        }
    }
    
    /// <summary>
    /// Extract project name from Revit window title
    /// </summary>
    private string ExtractProjectFromRevitTitle(string windowTitle)
    {
        if (string.IsNullOrEmpty(windowTitle)) return "";
        
        try
        {
            // Pattern 1: "Autodesk Revit 2025.4 - PROJECT-NAME.rvt - 3D View: ..."
            // Pattern 2: "Autodesk Revit 2024 - PROJECT.rvt - Floor Plan: ..."
            
            // Find the .rvt part
            var rvtIndex = windowTitle.IndexOf(".rvt", StringComparison.OrdinalIgnoreCase);
            if (rvtIndex > 0)
            {
                // Find the last " - " before .rvt
                var titleUpToRvt = windowTitle.Substring(0, rvtIndex);
                var lastDash = titleUpToRvt.LastIndexOf(" - ");
                if (lastDash > 0)
                {
                    var projectName = titleUpToRvt.Substring(lastDash + 3).Trim();
                    return projectName;
                }
            }
            
            // Fallback: try to find project in active processes
            foreach (var proc in _activeRevitProcesses.Values)
            {
                if (!string.IsNullOrEmpty(proc.ProjectName) && windowTitle.Contains(proc.ProjectName, StringComparison.OrdinalIgnoreCase))
                {
                    return proc.ProjectName;
                }
            }
        }
        catch { }
        
        return "";
    }
    
    /// <summary>
    /// Get window title from handle
    /// </summary>
    private string GetWindowTitle(IntPtr hWnd)
    {
        try
        {
            int length = GetWindowTextLength(hWnd);
            if (length == 0) return "";
            
            var builder = new System.Text.StringBuilder(length + 1);
            GetWindowText(hWnd, builder, builder.Capacity);
            return builder.ToString();
        }
        catch
        {
            return "";
        }
    }
    
    /// <summary>
    /// Get total ACTIVE Revit time today (only counts when Revit foreground + user active)
    /// </summary>
    public TimeSpan GetTodayActiveRevitTime()
    {
        var total = _totalRevitActiveTime;
        if (_isRevitForeground)
        {
            total += DateTime.Now - _revitActiveStartTime;
        }
        return total;
    }
    
    /// <summary>
    /// Get comprehensive activity breakdown for today
    /// STRICT: Uses hourly activity buckets for accurate, mutually exclusive time tracking
    /// </summary>
    public DailyActivityBreakdown GetDailyActivityBreakdown()
    {
        var now = Services.UserCredentialService.NowDubai;
        var today = now.Date;
        var idleSeconds = GetIdleTimeSeconds();
        
        // Calculate totals from hourly buckets (ACCURATE)
        double totalRevitMinutes = 0;
        double totalMeetingMinutes = 0;
        double totalIdleMinutes = 0;
        double totalOtherMinutes = 0;
        
        for (int hour = 0; hour <= now.Hour; hour++)
        {
            totalRevitMinutes += _hourlyRevitMinutes.GetValueOrDefault(hour, 0);
            totalMeetingMinutes += _hourlyMeetingMinutes.GetValueOrDefault(hour, 0);
            totalIdleMinutes += _hourlyIdleMinutes.GetValueOrDefault(hour, 0);
            totalOtherMinutes += _hourlyOtherMinutes.GetValueOrDefault(hour, 0);
        }
        
        var breakdown = new DailyActivityBreakdown
        {
            Date = today,
            RevitMinutes = Math.Round(totalRevitMinutes, 1),
            MeetingMinutes = Math.Round(totalMeetingMinutes, 1),
            TotalIdleMinutes = Math.Round(totalIdleMinutes, 1),
            OtherMinutes = Math.Round(totalOtherMinutes, 1)
        };
        
        // Build hourly breakdown using tracked activities
        BuildHourlyBreakdown(breakdown, now, today);
        
        return breakdown;
    }
    
    /// <summary>
    /// Build hourly breakdown using ONLY verified activity data
    /// </summary>
    private void BuildHourlyBreakdown(DailyActivityBreakdown breakdown, DateTime now, DateTime today)
    {
        for (int hour = 0; hour < 24; hour++)
        {
            var hourStart = today.AddHours(hour);
            var hourEnd = hourStart.AddHours(1);
            
            if (hourEnd > now) hourEnd = now;
            if (hourStart > now) continue;
            
            // Calculate actual elapsed minutes in this hour (max 60, or less if current hour)
            var elapsedMinutesInHour = (hourEnd - hourStart).TotalMinutes;
            
            var hourlyData = new HourlyActivity
            {
                Hour = hour,
                HourLabel = $"{(hour == 0 ? 12 : hour > 12 ? hour - 12 : hour)} {(hour < 12 ? "AM" : "PM")}"
            };
            
            // ========== USE HOURLY BUCKETS (ACCURATE, MUTUALLY EXCLUSIVE) ==========
            // Get tracked time from hourly buckets (already mutually exclusive)
            
            double revitMinutes = _hourlyRevitMinutes.GetValueOrDefault(hour, 0);
            double meetingMinutes = _hourlyMeetingMinutes.GetValueOrDefault(hour, 0);
            double idleMinutes = _hourlyIdleMinutes.GetValueOrDefault(hour, 0);
            double otherMinutes = _hourlyOtherMinutes.GetValueOrDefault(hour, 0);
            
            // Sanity check - total tracked time cannot exceed elapsed time
            var totalTracked = revitMinutes + meetingMinutes + idleMinutes + otherMinutes;
            
            if (totalTracked > elapsedMinutesInHour + 0.5) // Allow small rounding error
            {
                // Scale down proportionally
                var scale = elapsedMinutesInHour / totalTracked;
                revitMinutes *= scale;
                meetingMinutes *= scale;
                idleMinutes *= scale;
                otherMinutes *= scale;
            }
            else if (totalTracked < elapsedMinutesInHour - 1 && elapsedMinutesInHour >= 1)
            {
                // Not enough tracked - fill with "Other" for past hours, or current activity for current hour
                var untracked = elapsedMinutesInHour - totalTracked;
                if (hour < now.Hour)
                {
                    // Past hour - attribute untracked to "Other"
                    otherMinutes += untracked;
                }
                else
                {
                    // Current hour - attribute based on last activity type
                    switch (_lastActivityType)
                    {
                        case "Revit": revitMinutes += untracked; break;
                        case "Meeting": meetingMinutes += untracked; break;
                        case "Idle": idleMinutes += untracked; break;
                        default: otherMinutes += untracked; break;
                    }
                }
            }
            
            // Cap individual values at 60 and ensure non-negative
            hourlyData.RevitMinutes = Math.Round(Math.Max(0, Math.Min(60, revitMinutes)), 1);
            hourlyData.MeetingMinutes = Math.Round(Math.Max(0, Math.Min(60, meetingMinutes)), 1);
            hourlyData.IdleMinutes = Math.Round(Math.Max(0, Math.Min(60, idleMinutes)), 1);
            hourlyData.OtherMinutes = Math.Round(Math.Max(0, Math.Min(60, otherMinutes)), 1);
            
            // Final sanity check - total cannot exceed 60 (or elapsed for current hour)
            var finalTotal = hourlyData.RevitMinutes + hourlyData.MeetingMinutes + 
                            hourlyData.IdleMinutes + hourlyData.OtherMinutes;
            if (finalTotal > elapsedMinutesInHour + 0.5)
            {
                var scale = elapsedMinutesInHour / finalTotal;
                hourlyData.RevitMinutes = Math.Round(hourlyData.RevitMinutes * scale, 1);
                hourlyData.MeetingMinutes = Math.Round(hourlyData.MeetingMinutes * scale, 1);
                hourlyData.IdleMinutes = Math.Round(hourlyData.IdleMinutes * scale, 1);
                hourlyData.OtherMinutes = Math.Round(hourlyData.OtherMinutes * scale, 1);
            }
            
            hourlyData.IsWorkHour = hour >= 8 && hour < 18;
            // Overtime: before 8AM or at/after 6PM and user was active
            hourlyData.IsOvertime = (hour < 8 || hour >= 18) && (hourlyData.RevitMinutes > 0 || hourlyData.MeetingMinutes > 0 || hourlyData.OtherMinutes > 0);
            
            breakdown.HourlyBreakdown.Add(hourlyData);
        }
    }
    
    /// <summary>
    /// Get active time for a specific project (STRICT - foreground + user activity)
    /// </summary>
    public TimeSpan GetProjectActiveTime(string projectName)
    {
        if (string.IsNullOrEmpty(projectName)) return TimeSpan.Zero;
        
        var total = _projectActiveTime.GetValueOrDefault(projectName, TimeSpan.Zero);
        
        // Add current session if this project is active
        if (_isRevitForeground && _currentForegroundProject == projectName)
        {
            total += DateTime.Now - _lastProjectActiveUpdate;
        }
        
        return total;
    }
    
    /// <summary>
    /// Get all project active times (STRICT - only foreground + user activity time)
    /// Returns dictionary of project name to active hours
    /// </summary>
    public Dictionary<string, double> GetAllProjectActiveHours()
    {
        var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        
        foreach (var kvp in _projectActiveTime)
        {
            var hours = kvp.Value.TotalHours;
            
            // Add current session time if this project is currently active
            if (_isRevitForeground && _currentForegroundProject == kvp.Key)
            {
                hours += (DateTime.Now - _lastProjectActiveUpdate).TotalHours;
            }
            
            if (hours > 0)
            {
                result[kvp.Key] = Math.Round(hours, 2);
            }
        }
        
        // Also add currently active project if not in dictionary
        if (_isRevitForeground && !string.IsNullOrEmpty(_currentForegroundProject) && !result.ContainsKey(_currentForegroundProject))
        {
            var hours = (DateTime.Now - _lastProjectActiveUpdate).TotalHours;
            if (hours > 0)
            {
                result[_currentForegroundProject] = Math.Round(hours, 2);
            }
        }
        
        return result;
    }
    
    /// <summary>
    /// Check if currently tracking active Revit time
    /// </summary>
    public bool IsRevitActivelyTracking => _isRevitForeground;
    
    /// <summary>
    /// Get the current foreground project being tracked
    /// </summary>
    public string CurrentActiveProject => _currentForegroundProject;
    
    /// <summary>
    /// Get total idle time for TODAY ONLY (resets at midnight)
    /// </summary>
    public TimeSpan GetTodayIdleTime()
    {
        var today = DateTime.Today;
        
        // Auto-reset if tracking a different day
        if (_currentTrackingDay != today)
        {
            ResetForNewDay();
            _currentTrackingDay = today;
        }
        
        var idleTime = _totalIdleTime;
        
        // Add current idle session if active
        if (_isCurrentlyIdle && _idleStartTime > DateTime.MinValue)
        {
            // Only count idle time from today
            var idleStart = _idleStartTime.Date < today ? today : _idleStartTime;
            var currentIdle = DateTime.Now - idleStart;
            
            if (currentIdle.TotalHours >= 0)
            {
                idleTime += currentIdle;
            }
        }
        
        // Sanity check - idle time can't exceed hours since midnight
        var hoursSinceMidnight = (DateTime.Now - today).TotalHours;
        var maxIdleTime = TimeSpan.FromHours(Math.Min(hoursSinceMidnight, 24));
        
        if (idleTime > maxIdleTime)
        {
            idleTime = maxIdleTime;
        }
        
        // Ensure non-negative
        if (idleTime.TotalSeconds < 0)
        {
            idleTime = TimeSpan.Zero;
        }
        
        return idleTime;
    }
    
    /// <summary>
    /// Reset daily tracking (call at midnight or on new day)
    /// </summary>
    public void ResetDailyTracking()
    {
        _totalIdleTime = TimeSpan.Zero;
        _isCurrentlyIdle = false;
        _idleStartTime = DateTime.MinValue;
        _appActivities.Clear();
    }

    private void CheckStatus(object? state)
    {
        try
        {
            var status = ReadLoginState();

            // ── Revit fallback ────────────────────────────────────────────────
            // If loginstate.json couldn't be read but Revit is actively running,
            // treat the user as logged in. This covers non-standard install paths,
            // permission issues, and newer Autodesk Identity Manager layouts.
            if (!status.IsLoggedIn)
            {
                bool revitRunning = Process.GetProcessesByName("Revit").Length > 0;
                if (revitRunning)
                {
                    status.IsLoggedIn = true;
                    // Use the email we already know from config / previous successful read
                    if (string.IsNullOrEmpty(status.Email))
                        status.Email = _lastStatus.Email;
                    System.Diagnostics.Debug.WriteLine("[IDMonitor] loginstate.json not found — Revit.exe is running, marking as LoggedIn");
                }
                else
                {
                    // Log which paths were checked so it's visible in Debug tab
                    var checkedPaths = string.Join("; ", _possiblePaths.Where(p => !File.Exists(p)).Take(5));
                    System.Diagnostics.Debug.WriteLine($"[IDMonitor] Not logged in. Paths checked: {checkedPaths}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[IDMonitor] Logged in as {status.Email} via {_activeLoginStatePath}");
            }

            // Always notify on first check or if status changed
            if (_lastStatus.LastChecked == default ||
                status.IsLoggedIn != _lastStatus.IsLoggedIn ||
                status.Email != _lastStatus.Email)
            {
                _lastStatus = status;
                StatusChanged?.Invoke(this, status);
            }
        }
        catch { }
    }

    private AutodeskStatus ReadLoginState()
    {
        var status = new AutodeskStatus { LastChecked = DateTime.Now };

        // Try all possible paths
        foreach (var path in _possiblePaths)
        {
            try
            {
                if (!File.Exists(path))
                    continue;

                var json = File.ReadAllText(path);
                if (string.IsNullOrWhiteSpace(json))
                    continue;

                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                // Try different JSON structures used by various Autodesk versions
                
                // Structure 1: { "loggedInUser": { "email": "...", "userId": "..." } }
                if (root.TryGetProperty("loggedInUser", out var user) && user.ValueKind == JsonValueKind.Object)
                {
                    status.IsLoggedIn = true;
                    if (user.TryGetProperty("userId", out var userId))
                        status.UserId = userId.GetString();
                    
                    // Try multiple email field names
                    string? emailValue = TryGetEmailFromObject(user);
                    if (!string.IsNullOrEmpty(emailValue))
                        status.Email = emailValue;
                    
                    if (user.TryGetProperty("displayName", out var displayName))
                        status.DisplayName = displayName.GetString();
                    if (user.TryGetProperty("userName", out var userName))
                        status.DisplayName ??= userName.GetString();
                    
                    // If no email found but userId looks like email, use it
                    if (string.IsNullOrEmpty(status.Email) && !string.IsNullOrEmpty(status.UserId) && status.UserId.Contains("@"))
                        status.Email = status.UserId;
                    
                    _activeLoginStatePath = path;
                    if (!string.IsNullOrEmpty(status.Email))
                        return status;
                }

                // Structure 2: { "userId": "...", "email": "..." } (direct properties)
                if (root.TryGetProperty("userId", out var directUserId) && 
                    !string.IsNullOrEmpty(directUserId.GetString()))
                {
                    status.IsLoggedIn = true;
                    status.UserId = directUserId.GetString();
                    
                    // Try multiple email field names
                    string? emailValue = TryGetEmailFromObject(root);
                    if (!string.IsNullOrEmpty(emailValue))
                        status.Email = emailValue;
                    
                    if (root.TryGetProperty("displayName", out var displayName))
                        status.DisplayName = displayName.GetString();
                    
                    // If no email found but userId looks like email, use it
                    if (string.IsNullOrEmpty(status.Email) && status.UserId.Contains("@"))
                        status.Email = status.UserId;
                    
                    _activeLoginStatePath = path;
                    if (!string.IsNullOrEmpty(status.Email))
                        return status;
                }

                // Structure 3: { "user": { ... } }
                if (root.TryGetProperty("user", out var userObj) && userObj.ValueKind == JsonValueKind.Object)
                {
                    status.IsLoggedIn = true;
                    if (userObj.TryGetProperty("userId", out var uid))
                        status.UserId = uid.GetString();
                    
                    // Try multiple email field names
                    string? emailValue = TryGetEmailFromObject(userObj);
                    if (!string.IsNullOrEmpty(emailValue))
                        status.Email = emailValue;
                    
                    if (userObj.TryGetProperty("displayName", out var dn))
                        status.DisplayName = dn.GetString();
                    
                    // If no email found but userId looks like email, use it
                    if (string.IsNullOrEmpty(status.Email) && !string.IsNullOrEmpty(status.UserId) && status.UserId.Contains("@"))
                        status.Email = status.UserId;
                    
                    _activeLoginStatePath = path;
                    if (!string.IsNullOrEmpty(status.Email))
                        return status;
                }

                // Structure 4: Deep search - recursively find any email-like value
                var foundEmail = FindEmailInJson(root);
                if (!string.IsNullOrEmpty(foundEmail))
                {
                    status.IsLoggedIn = true;
                    status.Email = foundEmail;
                    _activeLoginStatePath = path;
                    return status;
                }
            }
            catch { }
        }

        return status;
    }

    // Helper method to try multiple email field names
    private string? TryGetEmailFromObject(JsonElement obj)
    {
        // Common email field names used by Autodesk
        string[] emailFields = { "email", "Email", "mail", "Mail", "userEmail", "UserEmail", 
                                  "emailAddress", "EmailAddress", "login", "Login", 
                                  "username", "UserName", "userName", "upn", "UPN" };
        
        foreach (var field in emailFields)
        {
            if (obj.TryGetProperty(field, out var prop) && prop.ValueKind == JsonValueKind.String)
            {
                var val = prop.GetString();
                if (!string.IsNullOrEmpty(val) && val.Contains("@") && val.Contains("."))
                    return val;
            }
        }
        return null;
    }

    // Deep search for email in JSON
    private string? FindEmailInJson(JsonElement element)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.String:
                var val = element.GetString() ?? "";
                if (val.Contains("@") && val.Contains(".") && !val.Contains(" "))
                    return val;
                break;
            case JsonValueKind.Object:
                // First try known email fields
                var email = TryGetEmailFromObject(element);
                if (!string.IsNullOrEmpty(email))
                    return email;
                // Then search all properties
                foreach (var prop in element.EnumerateObject())
                {
                    var found = FindEmailInJson(prop.Value);
                    if (!string.IsNullOrEmpty(found))
                        return found;
                }
                break;
            case JsonValueKind.Array:
                foreach (var item in element.EnumerateArray())
                {
                    var found = FindEmailInJson(item);
                    if (!string.IsNullOrEmpty(found))
                        return found;
                }
                break;
        }
        return null;
    }

    public AutodeskStatus GetCurrentStatus() => ReadLoginState();

    public string? GetActiveLoginStatePath() => _activeLoginStatePath;

    public void Dispose()
    {
        _watcher?.Dispose();
        _timer?.Dispose();
        _revitTimer?.Dispose();
        
        // End all active project times
        foreach (var entry in _projectTimes.Values.Where(p => p.IsActive))
        {
            entry.EndTime = DateTime.Now;
            entry.TotalDuration += entry.EndTime - entry.StartTime;
            entry.IsActive = false;
        }
    }
}

// ===== EVENT ARGS =====

public class RevitActivityEventArgs : EventArgs
{
    public RevitEventType EventType { get; set; }
    public RevitProcessInfo Process { get; set; } = new();
    public List<RevitProcessInfo> AllActiveProcesses { get; set; } = new();
}

public class ActivityStateEventArgs : EventArgs
{
    public ActivityState PreviousState { get; set; }
    public ActivityState CurrentState { get; set; }
    public int IdleSeconds { get; set; }
    public int ActiveRevitCount { get; set; }
}

// ===== ENUMS =====

public enum RevitEventType
{
    Opened,
    Closed,
    ProjectChanged
}

public enum ActivityState
{
    Active,
    Idle,
    Offline
}

// ===== MODELS =====

public class RevitProcessInfo
{
    public int ProcessId { get; set; }
    public string ProcessName { get; set; } = "";
    public string WindowTitle { get; set; } = "";
    public string RevitVersion { get; set; } = "";
    public string ProjectName { get; set; } = "";
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
    public DateTime LastActivity { get; set; }
    
    // For TODAY's duration calculation
    public DateTime TodayStartTime => StartTime.Date < DateTime.Today ? DateTime.Today : StartTime;
    
    // Duration shows ONLY today's time, not accumulated across days
    public TimeSpan Duration
    {
        get
        {
            var endTime = EndTime > DateTime.MinValue ? EndTime : DateTime.Now;
            var startTime = TodayStartTime;
            
            // If end time is before today, duration is zero for today
            if (endTime.Date < DateTime.Today)
                return TimeSpan.Zero;
                
            // Cap end time at now
            if (endTime > DateTime.Now)
                endTime = DateTime.Now;
                
            var duration = endTime - startTime;
            
            // Sanity check - can't be negative or more than 24 hours
            if (duration.TotalHours < 0)
                return TimeSpan.Zero;
            if (duration.TotalHours > 24)
                return TimeSpan.FromHours(24);
                
            return duration;
        }
    }
    
    public string DurationText
    {
        get
        {
            var d = Duration;
            return $"{(int)d.TotalHours:D2}:{d.Minutes:D2}:{d.Seconds:D2}";
        }
    }
}

public class MonitorProjectTimeEntry
{
    public string ProjectName { get; set; } = "";
    public string RevitVersion { get; set; } = "";
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
    public TimeSpan TotalDuration { get; set; }
    public bool IsActive { get; set; }
    
    // Track TODAY's accumulated duration separately
    public TimeSpan TodayDuration { get; set; } = TimeSpan.Zero;
    public DateTime TodaySessionStart { get; set; }
    
    // For TODAY's duration calculation
    public DateTime TodayStartTime => StartTime.Date < DateTime.Today ? DateTime.Today : StartTime;
    
    public TimeSpan CurrentDuration
    {
        get
        {
            // Start with today's accumulated duration
            var duration = TodayDuration;
            
            // Add current session if active
            if (IsActive && TodaySessionStart > DateTime.MinValue && TodaySessionStart.Date == DateTime.Today)
            {
                duration += DateTime.Now - TodaySessionStart;
            }
            
            // Sanity check - can't exceed 24 hours
            if (duration.TotalHours > 24)
                duration = TimeSpan.FromHours(24);
            if (duration.TotalHours < 0)
                duration = TimeSpan.Zero;
                
            return duration;
        }
    }
    
    public string DurationText
    {
        get
        {
            var d = CurrentDuration;
            return $"{(int)d.TotalHours:D2}:{d.Minutes:D2}:{d.Seconds:D2}";
        }
    }
    
    // Reset for new day
    public void ResetForNewDay()
    {
        TodayDuration = TimeSpan.Zero;
        if (IsActive)
        {
            TodaySessionStart = DateTime.Now;
        }
    }
}

public class ProjectDuration
{
    public string ProjectName { get; set; } = "";
    public string RevitVersion { get; set; } = "";
    public DateTime StartTime { get; set; }
    public TimeSpan Duration { get; set; }
    public bool IsActive { get; set; }
    
    public string DurationText
    {
        get
        {
            var d = Duration;
            return $"{(int)d.TotalHours:D2}:{d.Minutes:D2}:{d.Seconds:D2}";
        }
    }
}

public class RealTimeStats
{
    public ActivityState ActivityState { get; set; }
    public int IdleSeconds { get; set; }
    public int ActiveRevitCount { get; set; }
    public List<string> ActiveProjects { get; set; } = new();
    public List<ProjectDuration> ProjectDurations { get; set; } = new();
    
    public string ActivityStateText => ActivityState switch
    {
        ActivityState.Active => "🟢 Active",
        ActivityState.Idle => "🟡 Idle",
        ActivityState.Offline => "⚫ Offline",
        _ => "Unknown"
    };
    
    public string IdleText => IdleSeconds > 60 
        ? $"{IdleSeconds / 60}m {IdleSeconds % 60}s idle" 
        : IdleSeconds > 0 ? $"{IdleSeconds}s idle" : "";
}
