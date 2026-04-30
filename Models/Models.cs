using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;

namespace AutodeskIDMonitor.Models;

public partial class UserSession : ObservableObject
{
    [ObservableProperty] private string _personName = "";
    [ObservableProperty] private string _autodeskEmail = "";
    [ObservableProperty] private string _windowsUser = "";
    [ObservableProperty] private string _machineId = "";
    [ObservableProperty] private DateTime _lastActivity;
    [ObservableProperty] private bool _isLoggedIn;
    [ObservableProperty] private string _details = "-";
    [ObservableProperty] private bool _isCurrentUser;
    [ObservableProperty] private string _clientVersion = "";
    [ObservableProperty] private string _idStatus = "-";
    [ObservableProperty] private string _expectedEmail = "";
    [ObservableProperty] private string _revitVersion = "";
    [ObservableProperty] private string _currentProject = "";
    [ObservableProperty] private int _revitSessionCount;
    [ObservableProperty] private string _coWorkers = "";
    [ObservableProperty] private string _sharedIdInfo = "-";
    [ObservableProperty] private string _originalPersonName = "";
    
    public List<RevitProjectInfo> OpenProjects { get; set; } = new();
    
    public bool HasDetails => !string.IsNullOrEmpty(Details) && Details != "-";
    public bool HasSharedId => !string.IsNullOrEmpty(SharedIdInfo) && SharedIdInfo != "-";
}

// ===== TIME TRACKING MODELS =====

public partial class WorkSession : ObservableObject
{
    [ObservableProperty] private string _id = Guid.NewGuid().ToString();
    [ObservableProperty] private string _userId = "";
    [ObservableProperty] private string _userName = "";
    [ObservableProperty] private string _machineId = "";
    [ObservableProperty] private string _projectName = "";
    [ObservableProperty] private DateTime _startTime;
    [ObservableProperty] private DateTime _endTime;
    [ObservableProperty] private string _applicationName = "";
    [ObservableProperty] private string _revitVersion = "";
    [ObservableProperty] private bool _isActive;
    
    public TimeSpan Duration => EndTime > StartTime ? EndTime - StartTime : TimeSpan.Zero;
    public double Hours => Duration.TotalHours;
}

public partial class DailyWorkRecord : ObservableObject
{
    [ObservableProperty] private DateTime _date;
    [ObservableProperty] private string _userId = "";
    [ObservableProperty] private string _userName = "";
    [ObservableProperty] private string _machineId = "";
    [ObservableProperty] private DateTime _firstActivity;
    [ObservableProperty] private DateTime _lastActivity;
    [ObservableProperty] private double _totalHours;
    [ObservableProperty] private double _regularHours;
    [ObservableProperty] private double _overtimeHours;
    [ObservableProperty] private string _projectsWorked = "";
    [ObservableProperty] private int _projectCount;
    [ObservableProperty] private bool _isOnline;
    
    public List<WorkSession> Sessions { get; set; } = new();
    public List<string> ApplicationsUsed { get; set; } = new();
}

public partial class ProjectWorkSummary : ObservableObject
{
    [ObservableProperty] private string _projectName = "";
    [ObservableProperty] private double _totalHours;
    [ObservableProperty] private int _userCount;
    [ObservableProperty] private string _usersWorking = "";
    [ObservableProperty] private DateTime _lastActivity;
    
    public List<UserProjectTime> UserTimes { get; set; } = new();
}

public partial class UserProjectTime : ObservableObject
{
    [ObservableProperty] private string _userName = "";
    [ObservableProperty] private string _userId = "";
    [ObservableProperty] private double _hours;
    [ObservableProperty] private DateTime _lastActivity;
}

// ===== ENHANCED TIME TRACKING MODELS =====

public partial class UserDailyReport : ObservableObject
{
    [ObservableProperty] private string _userId = "";
    [ObservableProperty] private string _userName = "";
    [ObservableProperty] private string _machineId = "";
    [ObservableProperty] private DateTime _date;
    [ObservableProperty] private DateTime _firstActivity;
    [ObservableProperty] private DateTime _lastActivity;
    [ObservableProperty] private double _totalWorkHours;
    [ObservableProperty] private double _revitHours;
    [ObservableProperty] private double _idleHours;
    [ObservableProperty] private double _breakHours;
    [ObservableProperty] private double _meetingHours;
    [ObservableProperty] private double _otherAppHours;
    [ObservableProperty] private double _overtimeHours;
    [ObservableProperty] private double _regularHours;
    
    public List<ProjectTimeEntry> ProjectBreakdown { get; set; } = new();
    public List<ActivityTimeEntry> ActivityBreakdown { get; set; } = new();
    public List<HourlyActivity> HourlyBreakdown { get; set; } = new();
}

public partial class ProjectTimeEntry : ObservableObject
{
    [ObservableProperty] private string _projectName = "";
    [ObservableProperty] private double _hours;
    [ObservableProperty] private double _percentage;
    [ObservableProperty] private DateTime _startTime;
    [ObservableProperty] private DateTime _endTime;
    [ObservableProperty] private string _revitVersion = "";
    [ObservableProperty] private bool _isActive;
    
    public string DurationText => $"{TimeSpan.FromHours(Hours):hh\\:mm}";
    public TimeSpan CurrentDuration => IsActive ? (DateTime.Now - StartTime) : (EndTime - StartTime);
}

public partial class ActivityTimeEntry : ObservableObject
{
    [ObservableProperty] private string _activityType = ""; // Revit, Teams, Idle, Break, Other
    [ObservableProperty] private string _applicationName = "";
    [ObservableProperty] private double _hours;
    [ObservableProperty] private double _percentage;
    [ObservableProperty] private string _color = "#00B4D8";
}

public partial class HourlyActivity : ObservableObject
{
    [ObservableProperty] private int _hour; // 0-23
    [ObservableProperty] private string _hourLabel = ""; // "8 AM", "9 AM", etc.
    [ObservableProperty] private double _revitMinutes;
    [ObservableProperty] private double _meetingMinutes;
    [ObservableProperty] private double _idleMinutes;
    [ObservableProperty] private double _otherMinutes;
    [ObservableProperty] private bool _isWorkHour;
    [ObservableProperty] private bool _isOvertime;
}

public partial class ApplicationActivity : ObservableObject
{
    [ObservableProperty] private string _applicationName = "";
    [ObservableProperty] private string _windowTitle = "";
    [ObservableProperty] private DateTime _startTime;
    [ObservableProperty] private DateTime _endTime;
    [ObservableProperty] private bool _isActive;
    [ObservableProperty] private string _activityType = "Other"; // Revit, Meeting, Communication, Other
    
    public double Hours => IsActive ? (DateTime.Now - StartTime).TotalHours : (EndTime - StartTime).TotalHours;
}

public partial class UserDetailedProfile : ObservableObject
{
    [ObservableProperty] private string _userId = "";
    [ObservableProperty] private string _userName = "";
    [ObservableProperty] private string _autodeskEmail = "";
    [ObservableProperty] private string _machineId = "";
    [ObservableProperty] private DateTime _lastOnline;
    [ObservableProperty] private bool _isCurrentlyOnline;
    [ObservableProperty] private double _todayHours;
    [ObservableProperty] private double _weekHours;
    [ObservableProperty] private double _monthHours;
    [ObservableProperty] private double _todayOvertime;
    [ObservableProperty] private string _currentProject = "";
    [ObservableProperty] private string _currentApplication = "";
    [ObservableProperty] private string _revitVersion = "";
    
    public List<string> ActiveProjects { get; set; } = new();
    public List<string> ApplicationsOpen { get; set; } = new();
    public List<DailyWorkRecord> RecentDays { get; set; } = new();
}

public partial class CalendarDayEntry : ObservableObject
{
    [ObservableProperty] private DateTime _date;
    [ObservableProperty] private double _totalHours;
    [ObservableProperty] private double _overtimeHours;
    [ObservableProperty] private int _projectCount;
    [ObservableProperty] private string _projectsSummary = "";
    [ObservableProperty] private bool _hasData;
    [ObservableProperty] private bool _isToday;
    [ObservableProperty] private bool _isSelected;
    
    public string DayNumber => Date.Day.ToString();
    public string HoursDisplay => HasData ? $"{TotalHours:F1}h" : "";
    public string OvertimeDisplay => OvertimeHours > 0 ? $"+{OvertimeHours:F1}h OT" : "";
}

public class LocalDataStore
{
    public DateTime LastUpdated { get; set; }
    public List<UserDetailedProfile> UserProfiles { get; set; } = new();
    public List<DailyWorkRecord> DailyRecords { get; set; } = new();
    public List<WorkSession> ActiveSessions { get; set; } = new();
    public Dictionary<string, List<WorkSession>> SessionHistory { get; set; } = new();
    public List<HistoryEntry> ActivityHistory { get; set; } = new();
}

public class TimeTrackingConfig
{
    public TimeSpan WorkDayStart { get; set; } = new TimeSpan(8, 0, 0);   // 8:00 AM
    public TimeSpan WorkDayEnd { get; set; } = new TimeSpan(18, 0, 0);    // 6:00 PM
    public double MandatoryBreakHours { get; set; } = 1.0;
    public double StandardWorkHours { get; set; } = 9.0;  // 8AM-6PM = 10hrs minus 1hr mandatory break
}

// ===== REAL-TIME DASHBOARD MODELS =====

public partial class LiveUserActivity : ObservableObject
{
    [ObservableProperty] private string _userId = "";
    [ObservableProperty] private string _userName = "";
    [ObservableProperty] private string _machineId = "";
    [ObservableProperty] private string _autodeskEmail = "";
    [ObservableProperty] private string _activityStatus = "Offline";
    [ObservableProperty] private string _activityStatusIcon = "⚫";
    [ObservableProperty] private bool _isRevitOpen;
    [ObservableProperty] private int _revitInstanceCount;
    [ObservableProperty] private string _currentProject = "";
    [ObservableProperty] private string _revitVersion = "";
    [ObservableProperty] private DateTime _sessionStartTime;
    [ObservableProperty] private string _sessionDuration = "00:00:00";
    [ObservableProperty] private string _todayTotalHours = "0.0h";
    [ObservableProperty] private int _idleSeconds;
    [ObservableProperty] private string _idleText = "";
    [ObservableProperty] private DateTime _lastActivity;
    
    public List<string> ActiveProjects { get; set; } = new();
    public List<LiveProjectTime> ProjectTimes { get; set; } = new();
    
    public bool IsOnline => ActivityStatus != "Offline";
    public bool IsActive => ActivityStatus == "Active";
    public bool IsIdle => ActivityStatus == "Idle";
    
    /// <summary>
    /// Display projects as bullet list with line breaks
    /// </summary>
    public string ProjectsDisplay
    {
        get
        {
            if (ActiveProjects != null && ActiveProjects.Count > 0)
            {
                var lines = ActiveProjects
                    .Where(p => !string.IsNullOrEmpty(p))
                    .Select(p => $"• {p}")
                    .ToList();
                return lines.Count > 0 ? string.Join("\n", lines) : CurrentProject ?? "-";
            }
            return string.IsNullOrEmpty(CurrentProject) ? "-" : CurrentProject;
        }
    }
}

public partial class LiveProjectTime : ObservableObject
{
    [ObservableProperty] private string _projectName = "";
    [ObservableProperty] private string _revitVersion = "";
    [ObservableProperty] private DateTime _startTime;
    [ObservableProperty] private DateTime _endTime;
    [ObservableProperty] private string _duration = "00:00:00";
    [ObservableProperty] private bool _isActive;
    [ObservableProperty] private int _userCount;
    [ObservableProperty] private string _usersWorking = "";
    
    public TimeSpan DurationSpan => IsActive 
        ? DateTime.Now - StartTime 
        : (EndTime > DateTime.MinValue ? EndTime - StartTime : TimeSpan.Zero);
}

public partial class LiveProjectSummary : ObservableObject
{
    [ObservableProperty] private string _projectName = "";
    [ObservableProperty] private int _activeUserCount;
    [ObservableProperty] private string _activeUsers = "";
    [ObservableProperty] private string _totalHoursToday = "0.0h";
    [ObservableProperty] private DateTime _lastActivity;
    [ObservableProperty] private bool _hasActiveUsers;
    [ObservableProperty] private string _revitVersion = "";
    [ObservableProperty] private double _totalHoursWorked;
    
    public List<string> UserNames { get; set; } = new();
    public List<string> RevitVersions { get; set; } = new();
}

public partial class DashboardStats : ObservableObject
{
    [ObservableProperty] private int _totalUsers;
    [ObservableProperty] private int _activeUsers;
    [ObservableProperty] private int _idleUsers;
    [ObservableProperty] private int _offlineUsers;
    [ObservableProperty] private int _activeProjects;
    [ObservableProperty] private int _totalRevitInstances;
    [ObservableProperty] private string _totalWorkHoursToday = "0.0h";
    [ObservableProperty] private string _currentTime = "";
    [ObservableProperty] private string _workDayStatus = "";
}

public partial class HistoryEntry : ObservableObject
{
    [ObservableProperty] private DateTime _timestamp;
    [ObservableProperty] private string _eventType = "";
    [ObservableProperty] private string _personName = "";
    [ObservableProperty] private string _autodeskEmail = "";
    [ObservableProperty] private string _previousEmail = "";
    [ObservableProperty] private string _machineId = "";
}

public partial class UserProfile : ObservableObject
{
    [ObservableProperty] private string _name = "";
    [ObservableProperty] private string _email = "";
    [ObservableProperty] private string _password = "";
    [ObservableProperty] private bool _isActive;
    
    public string StatusText => IsActive ? "Active" : "Inactive";
    
    partial void OnIsActiveChanged(bool value)
    {
        OnPropertyChanged(nameof(StatusText));
    }
}

public partial class ProjectInfo : ObservableObject
{
    [ObservableProperty] private string _userName = "";
    [ObservableProperty] private string _projectsWithVersions = "";  // Combined: "Project1 (Revit 2024), Project2 (Revit 2025)"
    [ObservableProperty] private int _sessionCount;              
    [ObservableProperty] private string _machineName = "";
    [ObservableProperty] private DateTime _lastActivity;
    [ObservableProperty] private string _coWorkers = "";
    
    public bool HasProjects => !string.IsNullOrEmpty(ProjectsWithVersions) && ProjectsWithVersions != "-";
}

public class AppConfig
{
    public string Country { get; set; } = "UAE";
    public string Office { get; set; } = "Dubai";
    public int MonitoringIntervalSeconds { get; set; } = 5;
    public bool CloudLoggingEnabled { get; set; } = true;
    public string CloudApiUrl { get; set; } = "https://jdfzpnreoitpdhttielk.supabase.co";
    public string CloudApiKey { get; set; } = "REPLACE_WITH_SUPABASE_ANON_KEY";
    public bool NetworkLoggingEnabled { get; set; } = false;
    public string NetworkLogPath { get; set; } = "";
    public bool StartMinimized { get; set; } = true;
    public bool AutoStartEnabled { get; set; } = true;
}

public class AdminConfig
{
    public string PasswordHash { get; set; } = "";
}

public class RevitProjectInfo
{
    public string ProjectName { get; set; } = "";
    public string RevitVersion { get; set; } = "";
}

public class CloudSessionInfo
{
    public string WindowsUser { get; set; } = "";
    public string WindowsDisplayName { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string MachineName { get; set; } = "";
    public string AutodeskEmail { get; set; } = "";
    public bool IsLoggedIn { get; set; }
    public DateTime LastSeen { get; set; }
    public DateTime LastUpdate { get; set; }
    public string ClientVersion { get; set; } = "";
    public string RevitVersion { get; set; } = "";
    public string CurrentProject { get; set; } = "";
    public List<RevitProjectInfo> OpenProjects { get; set; } = new();  // All open projects with versions
    public int RevitSessionCount { get; set; }
    
    // Activity tracking data
    public double TodayRevitHours { get; set; }
    public double TodayMeetingHours { get; set; }
    public double TodayIdleHours { get; set; }
    public double TodayOtherHours { get; set; }
    public double TodayOvertimeHours { get; set; }
    public double TodayTotalHours { get; set; }
    public bool IsInMeeting { get; set; }
    public string MeetingApp { get; set; } = "";
    public int IdleSeconds { get; set; }
    public string ActivityState { get; set; } = "Offline";
    
    /// <summary>
    /// Get projects as bullet list for display
    /// </summary>
    public string ProjectsAsBulletList
    {
        get
        {
            if (OpenProjects == null || OpenProjects.Count == 0)
            {
                return string.IsNullOrEmpty(CurrentProject) ? "-" : CurrentProject;
            }
            
            // Format as bullet list with line breaks
            var lines = OpenProjects
                .Where(p => !string.IsNullOrEmpty(p.ProjectName))
                .Select(p => $"• {p.ProjectName}")
                .ToList();
            
            return lines.Count > 0 ? string.Join("\n", lines) : "-";
        }
    }
    
    /// <summary>
    /// Get first/main project for single-line display
    /// </summary>
    public string MainProject
    {
        get
        {
            if (OpenProjects != null && OpenProjects.Count > 0)
            {
                var first = OpenProjects.FirstOrDefault(p => !string.IsNullOrEmpty(p.ProjectName));
                if (first != null)
                {
                    var suffix = OpenProjects.Count > 1 ? $" (+{OpenProjects.Count - 1} more)" : "";
                    return first.ProjectName + suffix;
                }
            }
            return string.IsNullOrEmpty(CurrentProject) ? "-" : CurrentProject;
        }
    }

    public string GetDisplayName()
    {
        if (!string.IsNullOrEmpty(WindowsDisplayName)) return WindowsDisplayName;
        if (!string.IsNullOrEmpty(DisplayName)) return DisplayName;
        return WindowsUser;
    }
}

public class EmailUsageSummary
{
    public string AutodeskEmail { get; set; } = "";
    public int UserCount { get; set; }
    public List<string> UserNames { get; set; } = new();
    public List<string> MachineNames { get; set; } = new();
}

public class AutodeskStatus
{
    public bool IsLoggedIn { get; set; }
    public string? Email { get; set; }
    public string? UserId { get; set; }
    public string? DisplayName { get; set; }
    public DateTime LastChecked { get; set; }
}

// ===== MEETING & ACTIVITY TRACKING MODELS =====

public class MeetingAppInfo
{
    public string AppName { get; set; } = "";
    public string ProcessName { get; set; } = "";
    public string WindowTitle { get; set; } = "";
    public bool IsInActiveMeeting { get; set; }
    public DateTime StartTime { get; set; }
    
    public TimeSpan Duration => DateTime.Now - StartTime;
}

public class AppActivityEntry
{
    public string AppName { get; set; } = "";
    public string ActivityType { get; set; } = "Other"; // Revit, Meeting, Communication, Other
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
    public bool IsActive { get; set; }
    
    public TimeSpan Duration => IsActive ? (DateTime.Now - StartTime) : (EndTime - StartTime);
}

public class DailyActivityBreakdown
{
    public DateTime Date { get; set; }
    public double RevitMinutes { get; set; }
    public double MeetingMinutes { get; set; }
    public double TotalIdleMinutes { get; set; }
    public double OtherMinutes { get; set; }
    
    public double RevitHours => RevitMinutes / 60.0;
    public double MeetingHours => MeetingMinutes / 60.0;
    public double IdleHours => TotalIdleMinutes / 60.0;
    public double OtherHours => OtherMinutes / 60.0;
    
    public double TotalWorkHours => RevitHours + MeetingHours;
    
    public List<HourlyActivity> HourlyBreakdown { get; set; } = new();
}

/// <summary>
/// Historical daily activity record - persisted on server
/// </summary>
public class DailyActivityRecord
{
    public string WindowsUser { get; set; } = "";
    public string WindowsDisplayName { get; set; } = "";
    public string MachineName { get; set; } = "";
    public DateTime Date { get; set; }
    public double RevitHours { get; set; }
    public double MeetingHours { get; set; }
    public double IdleHours { get; set; }
    public double OtherHours { get; set; }
    public double OvertimeHours { get; set; }
    public double TotalHours { get; set; }
    public string CurrentProject { get; set; } = "";
    public List<string> ProjectsWorked { get; set; } = new();
    public DateTime FirstActivity { get; set; }
    public DateTime LastActivity { get; set; }
}

/// <summary>
/// Activity summary for a user - used by admin to view all users' time segregation
/// </summary>
public partial class UserActivitySummary : ObservableObject
{
    [ObservableProperty] private string _windowsUser = "";
    [ObservableProperty] private string _windowsDisplayName = "";
    [ObservableProperty] private string _machineName = "";
    [ObservableProperty] private DateTime _date;
    [ObservableProperty] private double _revitHours;
    [ObservableProperty] private double _meetingHours;
    [ObservableProperty] private double _idleHours;
    [ObservableProperty] private double _otherHours;
    [ObservableProperty] private double _overtimeHours;
    [ObservableProperty] private double _totalHours;
    [ObservableProperty] private bool _isInMeeting;
    [ObservableProperty] private string _meetingApp = "";
    [ObservableProperty] private int _idleSeconds;
    [ObservableProperty] private string _activityState = "Offline";
    [ObservableProperty] private string _currentProject = "";
    [ObservableProperty] private bool _isCurrentUser;
    
    // Project-level breakdown list
    public List<ProjectActivityBreakdown> ProjectBreakdowns { get; set; } = new();
    
    // Computed properties for UI
    public string StatusIcon => ActivityState switch
    {
        "Active" => "🟢",
        "Idle" => "🟡",
        "InMeeting" => "🟣",
        _ => "⚫"
    };
    
    public string MeetingStatus => IsInMeeting ? $"🟣 {MeetingApp}" : "";
    
    public string IdleText => IdleSeconds > 60 
        ? $"{IdleSeconds / 60}m idle" 
        : IdleSeconds > 0 ? $"{IdleSeconds}s idle" : "";
        
    public double ProductiveHours => RevitHours + MeetingHours;
    
    public string ProjectsListDisplay => ProjectBreakdowns.Count > 0 
        ? string.Join(", ", ProjectBreakdowns.Select(p => p.ProjectName)) 
        : CurrentProject;
}

/// <summary>
/// Project-level activity breakdown for a user
/// </summary>
public class ProjectActivityBreakdown
{
    public string ProjectName { get; set; } = "";
    public double ActiveRevitHours { get; set; }
    public double IdleHours { get; set; }
    public double OtherHours { get; set; }
    public DateTime FirstActivity { get; set; }
    public DateTime LastActivity { get; set; }
    public bool IsCurrentlyActive { get; set; }
    
    public string ActiveTimeDisplay => $"{ActiveRevitHours:F1}h";
    public string StatusDisplay => IsCurrentlyActive ? "🟢 Active" : "⚪ Inactive";
}

/// <summary>
/// Event args for hourly chart bar click
/// </summary>
public class HourlyActivityClickEventArgs : EventArgs
{
    public int Hour { get; set; }
    public string HourLabel { get; set; } = "";
    public HourlyActivity? HourlyData { get; set; }
    public string UserName { get; set; } = "";
}
