using AutodeskIDMonitor.Models;

namespace AutodeskIDMonitor.Services;

public class TimeTrackingService
{
    private readonly LocalStorageService _storage;
    private readonly TimeTrackingConfig _config;
    private readonly Dictionary<string, DateTime> _lastActivityTimes = new();
    private readonly Dictionary<string, HashSet<string>> _trackedProjects = new();

    public TimeTrackingService(LocalStorageService storage)
    {
        _storage = storage;
        _config = new TimeTrackingConfig();
    }

    /// <summary>
    /// Track user activity from a cloud session
    /// </summary>
    public void TrackActivity(CloudSessionInfo session)
    {
        var key = GetKey(session.MachineName, session.WindowsUser);
        var now = DateTime.Now;
        
        // Update profile
        _storage.UpdateProfileFromSession(session);
        
        if (!session.IsLoggedIn)
        {
            // User logged out - end all sessions
            EndAllUserSessions(session.MachineName, session.WindowsUser);
            return;
        }

        // Check for new projects
        var currentProjects = new HashSet<string>();
        
        // First try to get projects from OpenProjects
        if (session.OpenProjects?.Any() == true)
        {
            foreach (var p in session.OpenProjects.Where(p => !string.IsNullOrEmpty(p.ProjectName)))
            {
                currentProjects.Add(p.ProjectName);
            }
        }
        
        // If no projects from OpenProjects, use CurrentProject
        if (currentProjects.Count == 0 && !string.IsNullOrEmpty(session.CurrentProject))
        {
            currentProjects.Add(session.CurrentProject);
        }
        
        if (!_trackedProjects.ContainsKey(key))
            _trackedProjects[key] = new HashSet<string>();
        
        var previousProjects = _trackedProjects[key];
        
        // New projects started
        var newProjects = currentProjects.Except(previousProjects).ToList();
        foreach (var project in newProjects)
        {
            if (!string.IsNullOrEmpty(project))
            {
                _storage.StartWorkSession(
                    session.MachineName,
                    session.WindowsUser,
                    session.GetDisplayName(),
                    project,
                    "Revit",
                    session.RevitVersion);
            }
        }
        
        // Projects closed
        var closedProjects = previousProjects.Except(currentProjects).ToList();
        foreach (var project in closedProjects)
        {
            _storage.EndWorkSession(session.MachineName, session.WindowsUser, project);
        }
        
        _trackedProjects[key] = currentProjects;
        _lastActivityTimes[key] = now;
        
        // Note: We no longer use "Revit Session" as a fallback.
        // We only track time when we have an actual project name.
        // If user has Revit open but no project detected, we don't track generic "Revit Session" time.
    }

    /// <summary>
    /// End all active sessions for a user
    /// </summary>
    public void EndAllUserSessions(string machineId, string userId)
    {
        _storage.EndAllActiveSessions(machineId, userId);
        var key = GetKey(machineId, userId);
        _trackedProjects.Remove(key);
        _lastActivityTimes.Remove(key);
    }

    /// <summary>
    /// Get user's hours for today
    /// </summary>
    public double GetTodayHours(string machineId, string userId)
    {
        var record = _storage.GetUserDailyRecord(machineId, userId, DateTime.Today);
        if (record == null) return 0;
        
        // Add active session time
        var activeSessions = _storage.GetActiveSessions()
            .Where(s => s.MachineId == machineId && s.UserId == userId)
            .Sum(s => (DateTime.Now - s.StartTime).TotalHours);
        
        return record.TotalHours + activeSessions;
    }

    /// <summary>
    /// Get user's hours for this week
    /// </summary>
    public double GetWeekHours(string machineId, string userId)
    {
        var today = DateTime.Today;
        var startOfWeek = today.AddDays(-(int)today.DayOfWeek);
        var records = _storage.GetRecordsForDateRange(startOfWeek, today);
        
        return records
            .Where(r => r.MachineId == machineId && r.UserId == userId)
            .Sum(r => r.TotalHours);
    }

    /// <summary>
    /// Get user's hours for this month
    /// </summary>
    public double GetMonthHours(string machineId, string userId)
    {
        var today = DateTime.Today;
        var startOfMonth = new DateTime(today.Year, today.Month, 1);
        var records = _storage.GetRecordsForDateRange(startOfMonth, today);
        
        return records
            .Where(r => r.MachineId == machineId && r.UserId == userId)
            .Sum(r => r.TotalHours);
    }

    /// <summary>
    /// Get project statistics for a date
    /// </summary>
    public List<ProjectWorkSummary> GetProjectStats(DateTime date)
    {
        return _storage.GetProjectSummaries(date);
    }

    /// <summary>
    /// Get all users working on a project
    /// </summary>
    public List<UserProjectTime> GetUsersOnProject(string projectName, DateTime date)
    {
        var summaries = _storage.GetProjectSummaries(date);
        var project = summaries.FirstOrDefault(p => p.ProjectName == projectName);
        return project?.UserTimes ?? new List<UserProjectTime>();
    }

    /// <summary>
    /// Get detailed profile with updated stats
    /// </summary>
    public UserDetailedProfile? GetDetailedProfile(string machineId, string userId)
    {
        var key = GetKey(machineId, userId);
        var profile = _storage.GetProfile(key);
        
        if (profile != null)
        {
            profile.TodayHours = GetTodayHours(machineId, userId);
            profile.WeekHours = GetWeekHours(machineId, userId);
            profile.MonthHours = GetMonthHours(machineId, userId);
            
            var (regular, overtime) = CalculateOvertimeForDate(DateTime.Today, profile.TodayHours);
            profile.TodayOvertime = overtime;
            
            // Get recent history
            var last30Days = _storage.GetRecordsForDateRange(DateTime.Today.AddDays(-30), DateTime.Today);
            profile.RecentDays = last30Days
                .Where(r => r.MachineId == machineId && r.UserId == userId)
                .OrderByDescending(r => r.Date)
                .Take(30)
                .ToList();
        }
        
        return profile;
    }

    /// <summary>
    /// Calculate overtime for a given date.
    /// Regular = hours within 8AM-6PM window (capped at 9hrs after 1hr mandatory break).
    /// Overtime = any hours beyond the standard 9hr window.
    /// </summary>
    public (double regular, double overtime) CalculateOvertimeForDate(DateTime date, double totalHours)
    {
        var standardHours = _config.StandardWorkHours; // 9.0 (8AM-6PM minus 1hr break)
        var regularHours = Math.Min(totalHours, standardHours);
        var overtimeHours = Math.Max(0, totalHours - standardHours);
        return (regularHours, overtimeHours);
    }

    /// <summary>
    /// Check if work is currently overtime
    /// </summary>
    public bool IsCurrentlyOvertime()
    {
        var now = DateTime.Now.TimeOfDay;
        return now < _config.WorkDayStart || now >= _config.WorkDayEnd;
    }

    /// <summary>
    /// Get work day status message
    /// </summary>
    public string GetWorkDayStatus()
    {
        var now = DateTime.Now.TimeOfDay;
        
        if (now < _config.WorkDayStart)
            return $"Before work hours (starts at {_config.WorkDayStart:hh\\:mm})";
        
        if (now > _config.WorkDayEnd)
            return $"After work hours (ended at {_config.WorkDayEnd:hh\\:mm}) - Overtime";
        
        var remaining = _config.WorkDayEnd - now;
        return $"Work day in progress ({remaining.Hours}h {remaining.Minutes}m remaining)";
    }

    private string GetKey(string machineId, string userId) => $"{machineId}|{userId}".ToLower();
}
