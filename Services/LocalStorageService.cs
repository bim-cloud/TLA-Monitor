using System;
using System.Collections.Generic;
using System.Linq;
using AutodeskIDMonitor.Models;

namespace AutodeskIDMonitor.Services;

/// <summary>
/// LocalStorageService v5 — IN-MEMORY ONLY.
///
/// The v4 implementation wrote four files (cache.json, user_profiles.json,
/// active_sessions.json, plus a daily_records folder) to
/// %LOCALAPPDATA%\TangentIDMonitor\Data on every heartbeat. Over time those
/// daily_records JSONs accumulated and slowed down user machines.
///
/// In v5 we hold transient state in process memory only and let Supabase be
/// the durable store. Public method signatures are preserved so MainViewModel
/// and the rest of the codebase don't need to change.
/// </summary>
public class LocalStorageService
{
    private readonly LocalDataStore _cache = new();

    public LocalStorageService()
    {
        // No directory creation. No file IO. Nothing on disk.
    }

    // ===== USER PROFILES =====

    public List<UserDetailedProfile> GetAllProfiles() => _cache.UserProfiles.ToList();

    public UserDetailedProfile? GetProfile(string uniqueKey) =>
        _cache.UserProfiles.FirstOrDefault(p => GetKey(p.MachineId, p.UserId) == uniqueKey);

    public void SaveProfile(UserDetailedProfile profile)
    {
        var key = GetKey(profile.MachineId, profile.UserId);
        var existing = _cache.UserProfiles.FirstOrDefault(p => GetKey(p.MachineId, p.UserId) == key);
        if (existing != null) _cache.UserProfiles.Remove(existing);
        _cache.UserProfiles.Add(profile);
    }

    public void UpdateProfileFromSession(CloudSessionInfo session)
    {
        var key = GetKey(session.MachineName, session.WindowsUser);
        var profile = _cache.UserProfiles.FirstOrDefault(p => GetKey(p.MachineId, p.UserId) == key);
        if (profile == null)
        {
            profile = new UserDetailedProfile
            {
                MachineId = session.MachineName,
                UserId    = session.WindowsUser,
                CreatedAt = DateTime.Now
            };
            _cache.UserProfiles.Add(profile);
        }

        profile.UserName            = session.GetDisplayName();
        profile.AutodeskEmail       = session.AutodeskEmail;
        profile.IsCurrentlyOnline   = session.IsLoggedIn;
        profile.LastOnline          = session.LastSeen;
        profile.CurrentProject      = session.CurrentProject;
        profile.RevitVersion        = session.RevitVersion;
        profile.ActiveProjects      = session.OpenProjects?.Select(p => p.ProjectName).ToList() ?? new List<string>();

        if (session.RevitSessionCount > 0 && !profile.ApplicationsOpen.Contains("Revit"))
            profile.ApplicationsOpen.Add("Revit");
    }

    // ===== DAILY RECORDS (in-memory) =====

    public List<DailyWorkRecord> GetDailyRecords(DateTime date) =>
        _cache.DailyRecords.Where(r => r.Date.Date == date.Date).ToList();

    public List<DailyWorkRecord> GetRecordsForDateRange(DateTime startDate, DateTime endDate) =>
        _cache.DailyRecords
              .Where(r => r.Date.Date >= startDate.Date && r.Date.Date <= endDate.Date)
              .ToList();

    public DailyWorkRecord? GetUserDailyRecord(string machineId, string userId, DateTime date) =>
        _cache.DailyRecords.FirstOrDefault(r =>
            r.MachineId == machineId && r.UserId == userId && r.Date.Date == date.Date);

    public void SaveDailyRecord(DailyWorkRecord record)
    {
        var existing = GetUserDailyRecord(record.MachineId, record.UserId, record.Date);
        if (existing != null) _cache.DailyRecords.Remove(existing);
        _cache.DailyRecords.Add(record);
    }

    // ===== WORK SESSIONS =====

    private static readonly HashSet<string> InvalidProjectNames =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "Home", "Start", "Revit", "Recent", "New", "Open",
            "Models", "Families", "Templates", "Samples", "Autodesk",
            "Architecture", "Structure", "MEP", "Revit Session"
        };

    public void StartWorkSession(string machineId, string userId, string userName,
        string projectName, string application, string revitVersion = "")
    {
        EnforceDateBoundaries();
        if (string.IsNullOrEmpty(projectName) || InvalidProjectNames.Contains(projectName)) return;

        var today = DateTime.Today;
        var existing = _cache.ActiveSessions.FirstOrDefault(s =>
            s.MachineId == machineId && s.UserId == userId &&
            s.ProjectName == projectName && s.IsActive &&
            s.StartTime.Date == today);
        if (existing != null) return;

        // Drop stale sessions older than 24h
        var cutoff = DateTime.Now.AddHours(-24);
        var stale = _cache.ActiveSessions
            .Where(s => (s.StartTime < cutoff || s.StartTime.Date < today) && s.IsActive)
            .ToList();
        foreach (var s in stale)
        {
            s.EndTime  = s.StartTime.Date < today
                ? s.StartTime.Date.AddDays(1).AddSeconds(-1)
                : s.StartTime.AddMinutes(30);
            s.IsActive = false;
            _cache.ActiveSessions.Remove(s);
        }

        _cache.ActiveSessions.Add(new WorkSession
        {
            UserId          = userId,
            UserName        = userName,
            MachineId       = machineId,
            ProjectName     = projectName,
            ApplicationName = application,
            RevitVersion    = revitVersion,
            StartTime       = DateTime.Now,
            IsActive        = true
        });
    }

    public void EndWorkSession(string machineId, string userId, string projectName)
    {
        var session = _cache.ActiveSessions.FirstOrDefault(s =>
            s.MachineId == machineId && s.UserId == userId &&
            s.ProjectName == projectName && s.IsActive);
        if (session == null) return;

        session.EndTime = DateTime.Now;
        session.IsActive = false;

        var key = GetKey(machineId, userId);
        if (!_cache.SessionHistory.ContainsKey(key))
            _cache.SessionHistory[key] = new List<WorkSession>();
        _cache.SessionHistory[key].Add(session);
        _cache.ActiveSessions.Remove(session);

        UpdateDailyRecordFromSession(session);
    }

    public void EndAllActiveSessions(string machineId, string userId)
    {
        var sessions = _cache.ActiveSessions
            .Where(s => s.MachineId == machineId && s.UserId == userId && s.IsActive)
            .ToList();
        foreach (var s in sessions) EndWorkSession(machineId, userId, s.ProjectName);
    }

    public List<WorkSession> GetActiveSessions() =>
        _cache.ActiveSessions.Where(s => s.IsActive).ToList();

    public List<WorkSession> GetUserSessionHistory(string machineId, string userId,
        DateTime? fromDate = null, DateTime? toDate = null)
    {
        var key = GetKey(machineId, userId);
        if (!_cache.SessionHistory.ContainsKey(key)) return new List<WorkSession>();
        var sessions = _cache.SessionHistory[key].AsEnumerable();
        if (fromDate.HasValue) sessions = sessions.Where(s => s.StartTime >= fromDate.Value);
        if (toDate.HasValue)   sessions = sessions.Where(s => s.StartTime <= toDate.Value);
        return sessions.OrderByDescending(s => s.StartTime).ToList();
    }

    // ===== TIME CALCULATIONS =====

    public void UpdateDailyRecordFromSession(WorkSession session)
    {
        var date = session.StartTime.Date;
        var record = GetUserDailyRecord(session.MachineId, session.UserId, date)
                  ?? new DailyWorkRecord
                  {
                      Date          = date,
                      UserId        = session.UserId,
                      UserName      = session.UserName,
                      MachineId     = session.MachineId,
                      FirstActivity = session.StartTime
                  };

        record.Sessions.Add(session);
        record.LastActivity = session.EndTime > record.LastActivity ? session.EndTime : record.LastActivity;
        CalculateDailyTotals(record);
        SaveDailyRecord(record);
    }

    public void CalculateDailyTotals(DailyWorkRecord record)
    {
        var totalMinutes = record.Sessions.Sum(s => s.Duration.TotalMinutes);
        record.TotalHours = totalMinutes / 60.0;

        var standardHours = new TimeTrackingConfig().StandardWorkHours;
        record.RegularHours  = Math.Min(record.TotalHours, standardHours);
        record.OvertimeHours = Math.Max(0, record.TotalHours - standardHours);

        var projects = record.Sessions
            .Select(s => s.ProjectName)
            .Where(p => !string.IsNullOrEmpty(p))
            .Distinct()
            .ToList();
        record.ProjectsWorked = string.Join(", ", projects);
        record.ProjectCount   = projects.Count;
    }

    // ===== PROJECT SUMMARIES =====

    public List<ProjectWorkSummary> GetProjectSummaries(DateTime date)
    {
        var records = GetDailyRecords(date);
        return records
            .SelectMany(r => r.Sessions)
            .GroupBy(s => s.ProjectName)
            .Where(g => !string.IsNullOrEmpty(g.Key))
            .Select(g => new ProjectWorkSummary
            {
                ProjectName  = g.Key,
                TotalHours   = g.Sum(s => s.Duration.TotalHours),
                UserCount    = g.Select(s => s.UserId).Distinct().Count(),
                UsersWorking = string.Join(", ", g.Select(s => s.UserName).Distinct()),
                LastActivity = g.Max(s => s.EndTime),
                UserTimes = g.GroupBy(s => s.UserId)
                             .Select(ug => new UserProjectTime
                             {
                                 UserId       = ug.Key,
                                 UserName     = ug.First().UserName,
                                 Hours        = ug.Sum(s => s.Duration.TotalHours),
                                 LastActivity = ug.Max(s => s.EndTime)
                             }).ToList()
            })
            .OrderByDescending(p => p.TotalHours)
            .ToList();
    }

    // ===== CALENDAR =====

    public List<CalendarDayEntry> GetCalendarMonth(int year, int month, string? userId = null, string? machineId = null)
    {
        var entries = new List<CalendarDayEntry>();
        var first = new DateTime(year, month, 1);
        var last  = first.AddMonths(1).AddDays(-1);
        for (var d = first; d <= last; d = d.AddDays(1))
        {
            var records = GetDailyRecords(d);
            if (!string.IsNullOrEmpty(userId))
                records = records.Where(r => r.UserId == userId).ToList();
            if (!string.IsNullOrEmpty(machineId))
                records = records.Where(r => r.MachineId == machineId).ToList();

            entries.Add(new CalendarDayEntry
            {
                Date          = d,
                TotalHours    = records.Sum(r => r.TotalHours),
                OvertimeHours = records.Sum(r => r.OvertimeHours),
                IsWorkDay     = d.DayOfWeek != DayOfWeek.Friday && d.DayOfWeek != DayOfWeek.Saturday,
                HasActivity   = records.Any()
            });
        }
        return entries;
    }

    // ===== HOUSEKEEPING =====

    public DateTime GetLastSyncTime() => _cache.LastUpdated;

    public List<UserDetailedProfile> GetCachedProfiles() => _cache.UserProfiles.ToList();

    public void ClearDailyRecords(DateTime date)
    {
        var toRemove = _cache.DailyRecords.Where(r => r.Date.Date == date.Date).ToList();
        foreach (var r in toRemove) _cache.DailyRecords.Remove(r);
    }

    public void ClearAllActiveSessions() => _cache.ActiveSessions.Clear();

    public int CleanupStaleSessions(int maxAgeHours = 24)
    {
        var cutoff = DateTime.Now.AddHours(-maxAgeHours);
        var stale = _cache.ActiveSessions.Where(s => s.StartTime < cutoff).ToList();
        foreach (var s in stale)
        {
            s.IsActive = false;
            s.EndTime = s.StartTime.AddMinutes(30);
            _cache.ActiveSessions.Remove(s);
        }
        return stale.Count;
    }

    public void ResetAllTimeData()
    {
        _cache.DailyRecords.Clear();
        _cache.ActiveSessions.Clear();
        _cache.SessionHistory.Clear();
    }

    public void EnforceDateBoundaries()
    {
        var today = DateTime.Today;
        var crossDay = _cache.ActiveSessions
            .Where(s => s.IsActive && s.StartTime.Date < today)
            .ToList();
        foreach (var s in crossDay)
        {
            s.EndTime  = s.StartTime.Date.AddDays(1).AddSeconds(-1);
            s.IsActive = false;
            UpdateDailyRecordFromSession(s);
            _cache.ActiveSessions.Remove(s);
        }
    }

    public void RecalculateDailyData(DateTime date)
    {
        foreach (var r in GetDailyRecords(date)) CalculateDailyTotals(r);
    }

    // ===== USER DAILY REPORT (returns whatever the in-memory cache has) =====

    public UserDailyReport GetUserDailyReport(string machineId, string userId, DateTime date)
    {
        var record = GetUserDailyRecord(machineId, userId, date);
        if (record == null)
        {
            return new UserDailyReport
            {
                UserId    = userId,
                MachineId = machineId,
                Date      = date
            };
        }
        return new UserDailyReport
        {
            UserId         = record.UserId,
            UserName       = record.UserName,
            MachineId      = record.MachineId,
            Date           = record.Date,
            FirstActivity  = record.FirstActivity,
            LastActivity   = record.LastActivity,
            TotalWorkHours = record.TotalHours,
            RegularHours   = record.RegularHours,
            OvertimeHours  = record.OvertimeHours,
            ProjectBreakdown = record.Sessions
                .GroupBy(s => s.ProjectName)
                .Where(g => !string.IsNullOrEmpty(g.Key))
                .Select(g => new ProjectTimeEntry
                {
                    ProjectName = g.Key,
                    Hours       = g.Sum(s => s.Duration.TotalHours),
                    StartTime   = g.Min(s => s.StartTime),
                    EndTime     = g.Max(s => s.EndTime),
                }).ToList()
        };
    }

    public void ClearTodayAndRestart()
    {
        ClearDailyRecords(DateTime.Today);
        var today = _cache.ActiveSessions.Where(s => s.StartTime.Date == DateTime.Today).ToList();
        foreach (var s in today) _cache.ActiveSessions.Remove(s);
    }

    // ===== ACTIVITY HISTORY (kept in memory only) =====

    public void SaveActivityHistory(IEnumerable<HistoryEntry> entries)
    {
        _cache.ActivityHistory.Clear();
        _cache.ActivityHistory.AddRange(entries);
    }

    public List<HistoryEntry> LoadActivityHistory() => _cache.ActivityHistory.ToList();

    public void AddHistoryEntry(HistoryEntry entry)
    {
        _cache.ActivityHistory.Add(entry);
        // Keep only last 500 entries to bound memory
        if (_cache.ActivityHistory.Count > 500)
            _cache.ActivityHistory.RemoveRange(0, _cache.ActivityHistory.Count - 500);
    }

    public void DeleteDailyData(DateTime date) => ClearDailyRecords(date);

    // ===== PER-USER ACTIVITY BREAKDOWN =====

    private readonly Dictionary<string, UserDailyActivityRecord> _userDailyActivity = new();

    public void SaveUserDailyActivity(string userId, DateTime date, DailyActivityBreakdown breakdown)
    {
        var key = $"{userId}|{date:yyyy-MM-dd}";
        _userDailyActivity[key] = new UserDailyActivityRecord
        {
            UserId          = userId,
            Date            = date,
            RevitMinutes    = breakdown.RevitHours   * 60.0,
            MeetingMinutes  = breakdown.MeetingHours * 60.0,
            IdleMinutes     = breakdown.IdleHours    * 60.0,
            OtherMinutes    = breakdown.OtherHours   * 60.0,
            HourlyBreakdown = breakdown.HourlyBreakdown?
                .Select(h => new HourlyActivityRecord
                {
                    Hour           = h.Hour,
                    RevitMinutes   = h.RevitMinutes,
                    MeetingMinutes = h.MeetingMinutes,
                    IdleMinutes    = h.IdleMinutes,
                    OtherMinutes   = h.OtherMinutes
                }).ToList() ?? new List<HourlyActivityRecord>(),
            LastUpdated = DateTime.Now
        };
    }

    public UserDailyActivityRecord? GetUserDailyActivity(string userId, DateTime date)
    {
        var key = $"{userId}|{date:yyyy-MM-dd}";
        return _userDailyActivity.TryGetValue(key, out var rec) ? rec : null;
    }

    public List<string> GetUsersWithActivityOnDate(DateTime date)
    {
        var dateStr = date.ToString("yyyy-MM-dd");
        return _userDailyActivity.Keys
            .Where(k => k.EndsWith($"|{dateStr}"))
            .Select(k => k.Split('|')[0])
            .Distinct()
            .ToList();
    }

    public List<UserDailyActivityRecord> GetAllUserActivityForDate(DateTime date)
    {
        var dateStr = date.ToString("yyyy-MM-dd");
        return _userDailyActivity
            .Where(kv => kv.Key.EndsWith($"|{dateStr}"))
            .Select(kv => kv.Value)
            .ToList();
    }

    private static string GetKey(string machineId, string userId) => $"{machineId}|{userId}";
}

public class UserDailyActivityRecord
{
    public string UserId { get; set; } = "";
    public DateTime Date { get; set; }
    public double RevitMinutes { get; set; }
    public double MeetingMinutes { get; set; }
    public double IdleMinutes { get; set; }
    public double OtherMinutes { get; set; }
    public List<HourlyActivityRecord> HourlyBreakdown { get; set; } = new();
    public DateTime LastUpdated { get; set; }

    public double TotalProductiveMinutes => RevitMinutes + MeetingMinutes;
    public double TotalMinutes => RevitMinutes + MeetingMinutes + IdleMinutes + OtherMinutes;
}

public class HourlyActivityRecord
{
    public int Hour { get; set; }
    public double RevitMinutes { get; set; }
    public double MeetingMinutes { get; set; }
    public double IdleMinutes { get; set; }
    public double OtherMinutes { get; set; }
}
