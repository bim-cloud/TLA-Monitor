using System.IO;
using System.Linq;
using System.Text.Json;
using AutodeskIDMonitor.Models;

namespace AutodeskIDMonitor.Services;

public class LocalStorageService
{
    private readonly string _dataFolder;
    private readonly string _profilesFile;
    private readonly string _dailyRecordsFolder;
    private readonly string _sessionsFile;
    private readonly string _cacheFile;
    private LocalDataStore _cache = new();
    
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    public LocalStorageService()
    {
        _dataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TangentIDMonitor", "Data");
        
        Directory.CreateDirectory(_dataFolder);
        
        _profilesFile = Path.Combine(_dataFolder, "user_profiles.json");
        _dailyRecordsFolder = Path.Combine(_dataFolder, "daily_records");
        _sessionsFile = Path.Combine(_dataFolder, "active_sessions.json");
        _cacheFile = Path.Combine(_dataFolder, "cache.json");
        
        Directory.CreateDirectory(_dailyRecordsFolder);
        LoadCache();
    }

    private void LoadCache()
    {
        try
        {
            if (File.Exists(_cacheFile))
            {
                var json = File.ReadAllText(_cacheFile);
                _cache = JsonSerializer.Deserialize<LocalDataStore>(json, JsonOptions) ?? new LocalDataStore();
            }
        }
        catch { _cache = new LocalDataStore(); }
    }

    private void SaveCache()
    {
        try
        {
            _cache.LastUpdated = DateTime.Now;
            var json = JsonSerializer.Serialize(_cache, JsonOptions);
            File.WriteAllText(_cacheFile, json);
        }
        catch { }
    }

    // ===== USER PROFILES =====
    
    public List<UserDetailedProfile> GetAllProfiles()
    {
        return _cache.UserProfiles.ToList();
    }

    public UserDetailedProfile? GetProfile(string uniqueKey)
    {
        return _cache.UserProfiles.FirstOrDefault(p => GetKey(p.MachineId, p.UserId) == uniqueKey);
    }

    public void SaveProfile(UserDetailedProfile profile)
    {
        var key = GetKey(profile.MachineId, profile.UserId);
        var existing = _cache.UserProfiles.FirstOrDefault(p => GetKey(p.MachineId, p.UserId) == key);
        
        if (existing != null)
        {
            _cache.UserProfiles.Remove(existing);
        }
        _cache.UserProfiles.Add(profile);
        SaveCache();
    }

    public void UpdateProfileFromSession(CloudSessionInfo session)
    {
        var key = GetKey(session.MachineName, session.WindowsUser);
        var profile = _cache.UserProfiles.FirstOrDefault(p => GetKey(p.MachineId, p.UserId) == key);
        
        if (profile == null)
        {
            profile = new UserDetailedProfile
            {
                UserId = session.WindowsUser,
                MachineId = session.MachineName
            };
            _cache.UserProfiles.Add(profile);
        }
        
        profile.UserName = session.GetDisplayName();
        profile.AutodeskEmail = session.AutodeskEmail;
        profile.IsCurrentlyOnline = session.IsLoggedIn;
        profile.LastOnline = session.LastSeen;
        profile.CurrentProject = session.CurrentProject;
        profile.RevitVersion = session.RevitVersion;
        profile.ActiveProjects = session.OpenProjects?.Select(p => p.ProjectName).ToList() ?? new List<string>();
        
        if (session.RevitSessionCount > 0)
        {
            if (!profile.ApplicationsOpen.Contains("Revit"))
                profile.ApplicationsOpen.Add("Revit");
        }
        
        SaveCache();
    }

    // ===== DAILY RECORDS =====
    
    public List<DailyWorkRecord> GetDailyRecords(DateTime date)
    {
        var fileName = GetDailyFileName(date);
        var filePath = Path.Combine(_dailyRecordsFolder, fileName);
        
        if (File.Exists(filePath))
        {
            try
            {
                var json = File.ReadAllText(filePath);
                return JsonSerializer.Deserialize<List<DailyWorkRecord>>(json, JsonOptions) ?? new List<DailyWorkRecord>();
            }
            catch { }
        }
        
        // Return from cache if not on disk
        return _cache.DailyRecords.Where(r => r.Date.Date == date.Date).ToList();
    }

    public List<DailyWorkRecord> GetRecordsForDateRange(DateTime startDate, DateTime endDate)
    {
        var records = new List<DailyWorkRecord>();
        for (var date = startDate.Date; date <= endDate.Date; date = date.AddDays(1))
        {
            records.AddRange(GetDailyRecords(date));
        }
        return records;
    }

    public DailyWorkRecord? GetUserDailyRecord(string machineId, string userId, DateTime date)
    {
        var records = GetDailyRecords(date);
        return records.FirstOrDefault(r => r.MachineId == machineId && r.UserId == userId);
    }

    public void SaveDailyRecord(DailyWorkRecord record)
    {
        var date = record.Date.Date;
        var records = GetDailyRecords(date);
        
        var existing = records.FirstOrDefault(r => r.MachineId == record.MachineId && r.UserId == record.UserId);
        if (existing != null)
        {
            records.Remove(existing);
        }
        records.Add(record);
        
        var fileName = GetDailyFileName(date);
        var filePath = Path.Combine(_dailyRecordsFolder, fileName);
        var json = JsonSerializer.Serialize(records, JsonOptions);
        File.WriteAllText(filePath, json);
        
        // Update cache
        var cacheExisting = _cache.DailyRecords.FirstOrDefault(r => 
            r.Date.Date == date && r.MachineId == record.MachineId && r.UserId == record.UserId);
        if (cacheExisting != null)
            _cache.DailyRecords.Remove(cacheExisting);
        _cache.DailyRecords.Add(record);
        SaveCache();
    }

    // ===== WORK SESSIONS =====
    
    public void StartWorkSession(string machineId, string userId, string userName, string projectName, string application, string revitVersion = "")
    {
        // First, enforce date boundaries for any stale sessions
        EnforceDateBoundaries();
        
        // Filter out invalid project names
        var invalidNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Home", "Start", "Revit", "Recent", "New", "Open",
            "Models", "Families", "Templates", "Samples", "Autodesk",
            "Architecture", "Structure", "MEP", "Revit Session"
        };
        
        if (string.IsNullOrEmpty(projectName) || invalidNames.Contains(projectName))
            return;
        
        // Check if there's already an active session for this project today
        var today = DateTime.Today;
        var existingSession = _cache.ActiveSessions.FirstOrDefault(s => 
            s.MachineId == machineId && s.UserId == userId && 
            s.ProjectName == projectName && s.IsActive &&
            s.StartTime.Date == today);
        
        if (existingSession != null)
        {
            // Session already exists for today, just return
            return;
        }
        
        // Clean up stale sessions (older than 24 hours or from previous days)
        var cutoff = DateTime.Now.AddHours(-24);
        var staleSessions = _cache.ActiveSessions
            .Where(s => (s.StartTime < cutoff || s.StartTime.Date < today) && s.IsActive)
            .ToList();
        foreach (var stale in staleSessions)
        {
            // End at midnight if from previous day, otherwise 30 min after start
            if (stale.StartTime.Date < today)
            {
                stale.EndTime = stale.StartTime.Date.AddDays(1).AddSeconds(-1);
            }
            else
            {
                stale.EndTime = stale.StartTime.AddMinutes(30);
            }
            stale.IsActive = false;
            _cache.ActiveSessions.Remove(stale);
        }
        
        var session = new WorkSession
        {
            UserId = userId,
            UserName = userName,
            MachineId = machineId,
            ProjectName = projectName,
            ApplicationName = application,
            RevitVersion = revitVersion,
            StartTime = DateTime.Now,
            IsActive = true
        };
        
        _cache.ActiveSessions.Add(session);
        SaveCache();
    }

    public void EndWorkSession(string machineId, string userId, string projectName)
    {
        var session = _cache.ActiveSessions.FirstOrDefault(s => 
            s.MachineId == machineId && s.UserId == userId && 
            s.ProjectName == projectName && s.IsActive);
        
        if (session != null)
        {
            session.EndTime = DateTime.Now;
            session.IsActive = false;
            
            // Move to history
            var key = GetKey(machineId, userId);
            if (!_cache.SessionHistory.ContainsKey(key))
                _cache.SessionHistory[key] = new List<WorkSession>();
            _cache.SessionHistory[key].Add(session);
            
            _cache.ActiveSessions.Remove(session);
            SaveCache();
            
            // Update daily record
            UpdateDailyRecordFromSession(session);
        }
    }

    public void EndAllActiveSessions(string machineId, string userId)
    {
        var sessions = _cache.ActiveSessions
            .Where(s => s.MachineId == machineId && s.UserId == userId && s.IsActive)
            .ToList();
        
        foreach (var session in sessions)
        {
            EndWorkSession(machineId, userId, session.ProjectName);
        }
    }

    public List<WorkSession> GetActiveSessions()
    {
        return _cache.ActiveSessions.Where(s => s.IsActive).ToList();
    }

    public List<WorkSession> GetUserSessionHistory(string machineId, string userId, DateTime? fromDate = null, DateTime? toDate = null)
    {
        var key = GetKey(machineId, userId);
        if (!_cache.SessionHistory.ContainsKey(key))
            return new List<WorkSession>();
        
        var sessions = _cache.SessionHistory[key];
        
        if (fromDate.HasValue)
            sessions = sessions.Where(s => s.StartTime >= fromDate.Value).ToList();
        if (toDate.HasValue)
            sessions = sessions.Where(s => s.StartTime <= toDate.Value).ToList();
        
        return sessions.OrderByDescending(s => s.StartTime).ToList();
    }

    // ===== TIME CALCULATIONS =====
    
    public void UpdateDailyRecordFromSession(WorkSession session)
    {
        var date = session.StartTime.Date;
        var record = GetUserDailyRecord(session.MachineId, session.UserId, date);
        
        if (record == null)
        {
            record = new DailyWorkRecord
            {
                Date = date,
                UserId = session.UserId,
                UserName = session.UserName,
                MachineId = session.MachineId,
                FirstActivity = session.StartTime
            };
        }
        
        record.Sessions.Add(session);
        record.LastActivity = session.EndTime > record.LastActivity ? session.EndTime : record.LastActivity;
        
        // Recalculate totals
        CalculateDailyTotals(record);
        SaveDailyRecord(record);
    }

    public void CalculateDailyTotals(DailyWorkRecord record)
    {
        var totalMinutes = record.Sessions.Sum(s => s.Duration.TotalMinutes);
        record.TotalHours = totalMinutes / 60.0;
        
        // Standard hours: 8 AM to 6 PM = 10 hours - 1 hour break = 9 hours
        var config = new TimeTrackingConfig();
        var standardHours = config.StandardWorkHours;
        
        record.RegularHours = Math.Min(record.TotalHours, standardHours);
        record.OvertimeHours = Math.Max(0, record.TotalHours - standardHours);
        
        // Update project list
        var projects = record.Sessions.Select(s => s.ProjectName).Where(p => !string.IsNullOrEmpty(p)).Distinct().ToList();
        record.ProjectsWorked = string.Join(", ", projects);
        record.ProjectCount = projects.Count;
        
        // Update applications
        record.ApplicationsUsed = record.Sessions.Select(s => s.ApplicationName).Where(a => !string.IsNullOrEmpty(a)).Distinct().ToList();
    }

    // ===== REPORTING =====
    
    public List<ProjectWorkSummary> GetProjectSummaries(DateTime date)
    {
        var records = GetDailyRecords(date);
        var allSessions = new List<WorkSession>();
        
        // Get completed sessions from records
        allSessions.AddRange(records.SelectMany(r => r.Sessions));
        
        // Add active sessions with calculated time
        if (date.Date == DateTime.Today)
        {
            var activeSessions = _cache.ActiveSessions.Where(s => s.IsActive && s.StartTime.Date == date.Date).ToList();
            foreach (var active in activeSessions)
            {
                // Calculate current duration for active session
                var sessionCopy = new WorkSession
                {
                    Id = active.Id,
                    UserId = active.UserId,
                    UserName = active.UserName,
                    MachineId = active.MachineId,
                    ProjectName = active.ProjectName,
                    ApplicationName = active.ApplicationName,
                    RevitVersion = active.RevitVersion,
                    StartTime = active.StartTime,
                    EndTime = DateTime.Now, // Use current time for active sessions
                    IsActive = true
                };
                allSessions.Add(sessionCopy);
            }
        }
        
        var projectGroups = allSessions
            .Where(s => !string.IsNullOrEmpty(s.ProjectName))
            .GroupBy(s => s.ProjectName);
        
        return projectGroups.Select(g => new ProjectWorkSummary
        {
            ProjectName = g.Key,
            TotalHours = Math.Round(g.Sum(s => s.Hours), 1),
            UserCount = g.Select(s => s.UserId).Distinct().Count(),
            UsersWorking = string.Join(", ", g.Select(s => s.UserName).Distinct()),
            LastActivity = g.Max(s => s.EndTime > DateTime.MinValue ? s.EndTime : DateTime.Now),
            UserTimes = g.GroupBy(s => s.UserId).Select(ug => new UserProjectTime
            {
                UserId = ug.Key,
                UserName = ug.First().UserName,
                Hours = Math.Round(ug.Sum(s => s.Hours), 1),
                LastActivity = ug.Max(s => s.EndTime > DateTime.MinValue ? s.EndTime : DateTime.Now)
            }).ToList()
        }).ToList();
    }

    public List<CalendarDayEntry> GetCalendarMonth(int year, int month, string? userId = null, string? machineId = null)
    {
        var entries = new List<CalendarDayEntry>();
        var startDate = new DateTime(year, month, 1);
        var endDate = startDate.AddMonths(1).AddDays(-1);
        
        for (var date = startDate; date <= endDate; date = date.AddDays(1))
        {
            var records = GetDailyRecords(date);
            
            if (!string.IsNullOrEmpty(userId) && !string.IsNullOrEmpty(machineId))
            {
                records = records.Where(r => r.UserId == userId && r.MachineId == machineId).ToList();
            }
            
            var entry = new CalendarDayEntry
            {
                Date = date,
                TotalHours = records.Sum(r => r.TotalHours),
                OvertimeHours = records.Sum(r => r.OvertimeHours),
                ProjectCount = records.SelectMany(r => r.Sessions).Select(s => s.ProjectName).Distinct().Count(),
                HasData = records.Any(),
                IsToday = date.Date == DateTime.Today
            };
            
            entry.ProjectsSummary = string.Join(", ", records.SelectMany(r => r.Sessions).Select(s => s.ProjectName).Distinct().Take(3));
            entries.Add(entry);
        }
        
        return entries;
    }

    // ===== OFFLINE SUPPORT =====
    
    public DateTime GetLastSyncTime()
    {
        return _cache.LastUpdated;
    }

    public List<UserDetailedProfile> GetCachedProfiles()
    {
        return _cache.UserProfiles.ToList();
    }

    // ===== CLEANUP METHODS =====
    
    /// <summary>
    /// Clear all time tracking data for a specific date
    /// </summary>
    public void ClearDailyRecords(DateTime date)
    {
        var fileName = GetDailyFileName(date);
        var filePath = Path.Combine(_dailyRecordsFolder, fileName);
        
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
        
        // Also clear from cache
        _cache.DailyRecords.RemoveAll(r => r.Date.Date == date.Date);
        SaveCache();
    }
    
    /// <summary>
    /// Clear all active sessions and stale data
    /// </summary>
    public void ClearAllActiveSessions()
    {
        _cache.ActiveSessions.Clear();
        SaveCache();
    }
    
    /// <summary>
    /// Clean up stale sessions older than specified hours
    /// </summary>
    public int CleanupStaleSessions(int maxAgeHours = 24)
    {
        var cutoff = DateTime.Now.AddHours(-maxAgeHours);
        var staleSessions = _cache.ActiveSessions
            .Where(s => s.StartTime < cutoff && s.IsActive)
            .ToList();
        
        int cleaned = 0;
        foreach (var stale in staleSessions)
        {
            // End the session with a reasonable duration (30 min max for stale)
            stale.EndTime = stale.StartTime.AddMinutes(30);
            stale.IsActive = false;
            _cache.ActiveSessions.Remove(stale);
            cleaned++;
        }
        
        if (cleaned > 0)
            SaveCache();
        
        return cleaned;
    }
    
    /// <summary>
    /// Reset all time tracking data (use with caution)
    /// </summary>
    public void ResetAllTimeData()
    {
        _cache.ActiveSessions.Clear();
        _cache.SessionHistory.Clear();
        _cache.DailyRecords.Clear();
        SaveCache();
        
        // Delete all daily record files
        if (Directory.Exists(_dailyRecordsFolder))
        {
            foreach (var file in Directory.GetFiles(_dailyRecordsFolder, "daily_*.json"))
            {
                try { File.Delete(file); } catch { }
            }
        }
    }
    
    /// <summary>
    /// End sessions that span midnight (date boundary enforcement)
    /// </summary>
    public void EnforceDateBoundaries()
    {
        var today = DateTime.Today;
        var sessionsToFix = _cache.ActiveSessions
            .Where(s => s.IsActive && s.StartTime.Date < today)
            .ToList();
        
        foreach (var session in sessionsToFix)
        {
            // End the session at midnight of the day it started
            var midnight = session.StartTime.Date.AddDays(1);
            session.EndTime = midnight.AddSeconds(-1); // 23:59:59
            session.IsActive = false;
            
            // Save to history
            UpdateDailyRecordFromSession(session);
            _cache.ActiveSessions.Remove(session);
        }
        
        if (sessionsToFix.Count > 0)
            SaveCache();
    }
    
    /// <summary>
    /// Recalculate and fix all time data for a specific date
    /// </summary>
    public void RecalculateDailyData(DateTime date)
    {
        var records = GetDailyRecords(date);
        
        foreach (var record in records)
        {
            // Validate sessions - remove invalid ones
            record.Sessions = record.Sessions
                .Where(s => s.StartTime.Date == date.Date && 
                            s.Duration.TotalHours < 12 && // No session longer than 12 hours
                            s.Duration.TotalMinutes > 0) // Must have some duration
                .ToList();
            
            // Recalculate totals
            CalculateDailyTotals(record);
            SaveDailyRecord(record);
        }
    }
    
    /// <summary>
    /// Get detailed user daily report with activity breakdown
    /// </summary>
    public UserDailyReport GetUserDailyReport(string machineId, string userId, DateTime date)
    {
        var record = GetUserDailyRecord(machineId, userId, date);
        var config = new TimeTrackingConfig();
        
        var report = new UserDailyReport
        {
            UserId = userId,
            MachineId = machineId,
            Date = date
        };
        
        if (record == null)
            return report;
        
        report.UserName = record.UserName;
        report.FirstActivity = record.FirstActivity;
        report.LastActivity = record.LastActivity;
        report.TotalWorkHours = record.TotalHours;
        report.RegularHours = record.RegularHours;
        report.OvertimeHours = record.OvertimeHours;
        
        // Calculate project breakdown
        var projectGroups = record.Sessions
            .Where(s => !string.IsNullOrEmpty(s.ProjectName))
            .GroupBy(s => s.ProjectName);
        
        double totalProjectHours = record.Sessions.Sum(s => s.Hours);
        
        foreach (var group in projectGroups.OrderByDescending(g => g.Sum(s => s.Hours)))
        {
            var projectHours = group.Sum(s => s.Hours);
            report.ProjectBreakdown.Add(new ProjectTimeEntry
            {
                ProjectName = group.Key,
                Hours = Math.Round(projectHours, 2),
                Percentage = totalProjectHours > 0 ? Math.Round(projectHours / totalProjectHours * 100, 1) : 0,
                StartTime = group.Min(s => s.StartTime),
                EndTime = group.Max(s => s.EndTime),
                RevitVersion = group.First().RevitVersion
            });
        }
        
        // Add activity breakdown
        report.RevitHours = record.Sessions
            .Where(s => s.ApplicationName?.Contains("Revit") == true)
            .Sum(s => s.Hours);
        
        report.ActivityBreakdown.Add(new ActivityTimeEntry
        {
            ActivityType = "Revit",
            ApplicationName = "Autodesk Revit",
            Hours = Math.Round(report.RevitHours, 2),
            Percentage = report.TotalWorkHours > 0 ? Math.Round(report.RevitHours / report.TotalWorkHours * 100, 1) : 0,
            Color = "#00B4D8"
        });
        
        // Build hourly breakdown
        for (int hour = 0; hour < 24; hour++)
        {
            var hourStart = date.AddHours(hour);
            var hourEnd = hourStart.AddHours(1);
            
            var sessionsInHour = record.Sessions
                .Where(s => s.StartTime < hourEnd && s.EndTime > hourStart)
                .ToList();
            
            double revitMinutes = 0;
            foreach (var session in sessionsInHour)
            {
                var overlapStart = session.StartTime > hourStart ? session.StartTime : hourStart;
                var overlapEnd = session.EndTime < hourEnd ? session.EndTime : hourEnd;
                if (overlapEnd > overlapStart)
                {
                    revitMinutes += (overlapEnd - overlapStart).TotalMinutes;
                }
            }
            
            var isWorkHour = hour >= config.WorkDayStart.Hours && hour < config.WorkDayEnd.Hours;
            var isOvertime = hour >= config.WorkDayEnd.Hours && revitMinutes > 0;
            
            report.HourlyBreakdown.Add(new HourlyActivity
            {
                Hour = hour,
                HourLabel = $"{(hour == 0 ? 12 : hour > 12 ? hour - 12 : hour)} {(hour < 12 ? "AM" : "PM")}",
                RevitMinutes = Math.Round(revitMinutes, 1),
                IsWorkHour = isWorkHour,
                IsOvertime = isOvertime
            });
        }
        
        return report;
    }
    
    /// <summary>
    /// Clear ONLY today's corrupted data and start fresh
    /// </summary>
    public void ClearTodayAndRestart()
    {
        var today = DateTime.Today;
        
        // Clear today's file
        ClearDailyRecords(today);
        
        // End all active sessions without saving to history
        _cache.ActiveSessions.Clear();
        
        // Clear session history for today
        foreach (var key in _cache.SessionHistory.Keys.ToList())
        {
            _cache.SessionHistory[key] = _cache.SessionHistory[key]
                .Where(s => s.StartTime.Date != today)
                .ToList();
        }
        
        SaveCache();
    }
    
    // ===== ACTIVITY HISTORY =====
    
    /// <summary>
    /// Save activity history entries
    /// </summary>
    public void SaveActivityHistory(IEnumerable<HistoryEntry> entries)
    {
        _cache.ActivityHistory = entries.Take(500).ToList(); // Keep last 500 entries
        SaveCache();
    }
    
    /// <summary>
    /// Load saved activity history
    /// </summary>
    public List<HistoryEntry> LoadActivityHistory()
    {
        return _cache.ActivityHistory ?? new List<HistoryEntry>();
    }
    
    /// <summary>
    /// Add a single history entry
    /// </summary>
    public void AddHistoryEntry(HistoryEntry entry)
    {
        _cache.ActivityHistory.Insert(0, entry);
        
        // Keep only last 500 entries
        while (_cache.ActivityHistory.Count > 500)
        {
            _cache.ActivityHistory.RemoveAt(_cache.ActivityHistory.Count - 1);
        }
        
        SaveCache();
    }
    
    /// <summary>
    /// Delete all time tracking data for a specific date
    /// </summary>
    public void DeleteDailyData(DateTime date)
    {
        try
        {
            var filePath = Path.Combine(_dailyRecordsFolder, GetDailyFileName(date));
            
            if (File.Exists(filePath))
            {
                // Create backup before deletion
                var backupPath = filePath + $".backup_{DateTime.Now:yyyyMMdd_HHmmss}";
                File.Copy(filePath, backupPath);
                
                // Delete the file
                File.Delete(filePath);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error deleting daily data: {ex.Message}");
            throw;
        }
    }
    
    // ===== USER ACTIVITY STORAGE (DAILY, HOURLY) =====
    
    private string GetUserActivityFolder()
    {
        var folder = Path.Combine(_dataFolder, "user_activity");
        Directory.CreateDirectory(folder);
        return folder;
    }
    
    /// <summary>
    /// Save user's daily activity breakdown (call periodically and at end of day)
    /// </summary>
    public void SaveUserDailyActivity(string userId, DateTime date, DailyActivityBreakdown breakdown)
    {
        try
        {
            var folder = GetUserActivityFolder();
            var fileName = $"activity_{userId.Replace("\\", "_")}_{date:yyyy-MM-dd}.json";
            var filePath = Path.Combine(folder, fileName);
            
            var record = new UserDailyActivityRecord
            {
                UserId = userId,
                Date = date.Date,
                RevitMinutes = breakdown.RevitMinutes,
                MeetingMinutes = breakdown.MeetingMinutes,
                IdleMinutes = breakdown.TotalIdleMinutes,
                OtherMinutes = breakdown.OtherMinutes,
                HourlyBreakdown = breakdown.HourlyBreakdown.Select(h => new HourlyActivityRecord
                {
                    Hour = h.Hour,
                    RevitMinutes = h.RevitMinutes,
                    MeetingMinutes = h.MeetingMinutes,
                    IdleMinutes = h.IdleMinutes,
                    OtherMinutes = h.OtherMinutes
                }).ToList(),
                LastUpdated = DateTime.Now
            };
            
            var json = JsonSerializer.Serialize(record, JsonOptions);
            File.WriteAllText(filePath, json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error saving user activity: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Get user's daily activity for a specific date
    /// </summary>
    public UserDailyActivityRecord? GetUserDailyActivity(string userId, DateTime date)
    {
        try
        {
            var folder = GetUserActivityFolder();
            var fileName = $"activity_{userId.Replace("\\", "_")}_{date:yyyy-MM-dd}.json";
            var filePath = Path.Combine(folder, fileName);
            
            if (File.Exists(filePath))
            {
                var json = File.ReadAllText(filePath);
                return JsonSerializer.Deserialize<UserDailyActivityRecord>(json, JsonOptions);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error loading user activity: {ex.Message}");
        }
        return null;
    }
    
    /// <summary>
    /// Get all users who have activity data for a specific date
    /// </summary>
    public List<string> GetUsersWithActivityOnDate(DateTime date)
    {
        var users = new List<string>();
        try
        {
            var folder = GetUserActivityFolder();
            var dateStr = date.ToString("yyyy-MM-dd");
            
            foreach (var file in Directory.GetFiles(folder, $"activity_*_{dateStr}.json"))
            {
                var fileName = Path.GetFileNameWithoutExtension(file);
                // Extract userId from "activity_USERID_DATE"
                var parts = fileName.Split('_');
                if (parts.Length >= 2)
                {
                    users.Add(parts[1].Replace("_", "\\"));
                }
            }
        }
        catch { }
        return users;
    }
    
    /// <summary>
    /// Get activity data for all users on a specific date
    /// </summary>
    public List<UserDailyActivityRecord> GetAllUserActivityForDate(DateTime date)
    {
        var records = new List<UserDailyActivityRecord>();
        
        foreach (var userId in GetUsersWithActivityOnDate(date))
        {
            var record = GetUserDailyActivity(userId, date);
            if (record != null)
            {
                records.Add(record);
            }
        }
        
        return records;
    }

    // ===== HELPERS =====
    
    private string GetKey(string machineId, string userId) => $"{machineId}|{userId}".ToLower();
    private string GetDailyFileName(DateTime date) => $"daily_{date:yyyy-MM-dd}.json";
}

/// <summary>
/// Record of a user's daily activity (stored per user, per day)
/// </summary>
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

/// <summary>
/// Hourly activity record for storage
/// </summary>
public class HourlyActivityRecord
{
    public int Hour { get; set; }
    public double RevitMinutes { get; set; }
    public double MeetingMinutes { get; set; }
    public double IdleMinutes { get; set; }
    public double OtherMinutes { get; set; }
}

/// <summary>
/// User profile settings stored locally
/// </summary>
public class UserProfileSettings
{
    public string DisplayName { get; set; } = "";
    public string Email { get; set; } = "";
    public bool SetupCompleted { get; set; } = false;
    public DateTime FirstSetupDate { get; set; }
    public DateTime LastModified { get; set; }
}

/// <summary>
/// Extension methods for LocalStorageService - User Profile
/// </summary>
public static class LocalStorageExtensions
{
    private static readonly string SettingsFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TangentIDMonitor", "Data", "user_settings.json");
    
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };
    
    /// <summary>
    /// Check if first-time setup is needed
    /// </summary>
    public static bool NeedsFirstTimeSetup(this LocalStorageService storage)
    {
        try
        {
            if (!File.Exists(SettingsFile))
                return true;
            
            var json = File.ReadAllText(SettingsFile);
            var settings = JsonSerializer.Deserialize<UserProfileSettings>(json, JsonOptions);
            
            return settings == null || !settings.SetupCompleted || string.IsNullOrEmpty(settings.DisplayName);
        }
        catch
        {
            return true;
        }
    }
    
    /// <summary>
    /// Get user profile settings
    /// </summary>
    public static UserProfileSettings GetUserProfileSettings(this LocalStorageService storage)
    {
        try
        {
            if (File.Exists(SettingsFile))
            {
                var json = File.ReadAllText(SettingsFile);
                return JsonSerializer.Deserialize<UserProfileSettings>(json, JsonOptions) ?? new UserProfileSettings();
            }
        }
        catch { }
        
        return new UserProfileSettings();
    }
    
    /// <summary>
    /// Save user profile settings
    /// </summary>
    public static void SaveUserProfileSettings(this LocalStorageService storage, string displayName, string email)
    {
        try
        {
            var settings = new UserProfileSettings
            {
                DisplayName = displayName,
                Email = email,
                SetupCompleted = true,
                FirstSetupDate = DateTime.Now,
                LastModified = DateTime.Now
            };
            
            // Load existing if present to preserve first setup date
            if (File.Exists(SettingsFile))
            {
                try
                {
                    var existing = JsonSerializer.Deserialize<UserProfileSettings>(File.ReadAllText(SettingsFile), JsonOptions);
                    if (existing != null && existing.FirstSetupDate > DateTime.MinValue)
                    {
                        settings.FirstSetupDate = existing.FirstSetupDate;
                    }
                }
                catch { }
            }
            
            var json = JsonSerializer.Serialize(settings, JsonOptions);
            File.WriteAllText(SettingsFile, json);
        }
        catch { }
    }
    
    /// <summary>
    /// Auto-correct Autodesk email format
    /// </summary>
    public static string AutoCorrectAutodeskEmail(string rawEmail, string displayName = "")
    {
        if (string.IsNullOrEmpty(rawEmail))
            return "";
        
        rawEmail = rawEmail.Trim();
        
        // If already a valid email format, return as-is
        if (rawEmail.Contains("@") && rawEmail.Contains("."))
        {
            return rawEmail.ToLower();
        }
        
        // Try to construct proper email from partial name
        var cleanName = System.Text.RegularExpressions.Regex.Replace(rawEmail.ToLower(), @"[^a-z0-9]", "");
        
        // If we have a display name, try to use it to format the email
        if (!string.IsNullOrEmpty(displayName))
        {
            var nameParts = displayName.Trim().ToLower().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (nameParts.Length >= 2)
            {
                // Format as firstname.lastname@domain
                return $"{nameParts[0]}.{nameParts[nameParts.Length - 1]}@tangentlandscape.com";
            }
        }
        
        // Try common name patterns
        string[] commonFirstNames = { "amna", "anshu", "john", "jane", "mike", "sarah", "david", "lisa", 
            "ahmed", "mohammed", "fatima", "ali", "omar", "layla", "noor", "hassan", "hussein",
            "abinshan", "adhithyan", "afsal", "akshaya", "ananthu", "aparna", "athira", "jesto",
            "laura", "mohamed", "rajeev", "amna" };
        
        foreach (var firstName in commonFirstNames)
        {
            if (cleanName.StartsWith(firstName) && cleanName.Length > firstName.Length)
            {
                var lastName = cleanName.Substring(firstName.Length);
                return $"{firstName}.{lastName}@tangentlandscape.com";
            }
        }
        
        // Default: add domain to cleaned name
        return $"{cleanName}@tangentlandscape.com";
    }
}
