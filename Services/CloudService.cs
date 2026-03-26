using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using AutodeskIDMonitor.Models;

namespace AutodeskIDMonitor.Services;

public class CloudService
{
    private readonly HttpClient _httpClient;
    
    public CloudService()
    {
        _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        UpdateApiKey();
    }

    public void UpdateApiKey()
    {
        _httpClient.DefaultRequestHeaders.Remove("X-API-Key");
        var apiKey = ConfigService.Instance.Config.CloudApiKey;
        if (!string.IsNullOrEmpty(apiKey))
        {
            _httpClient.DefaultRequestHeaders.Add("X-API-Key", apiKey);
        }
    }

    private string BaseUrl => ConfigService.Instance.Config.CloudApiUrl?.TrimEnd('/') ?? "";

    public async Task<bool> SendStatusAsync(string windowsUser, string displayName, string machineName,
        string autodeskEmail, bool isLoggedIn, string country, string office, string version)
    {
        if (!ConfigService.Instance.Config.CloudLoggingEnabled || string.IsNullOrEmpty(BaseUrl))
            return false;

        try
        {
            var status = new
            {
                WindowsUser = windowsUser,
                WindowsDisplayName = displayName,
                MachineName = machineName,
                AutodeskEmail = autodeskEmail,
                IsLoggedIn = isLoggedIn,
                Country = country,
                Office = office,
                ClientVersion = version,
                LastUpdate = DateTime.UtcNow
            };

            var response = await _httpClient.PostAsJsonAsync($"{BaseUrl}/api/status", status);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Send real-time activity status including Revit processes and activity state
    /// </summary>
    public async Task<bool> SendActivityStatusAsync(string windowsUser, string displayName, string machineName,
        string autodeskEmail, bool isLoggedIn, string country, string office, string version,
        string activityState, int idleSeconds, List<RevitProcessInfo> revitProcesses, 
        Dictionary<string, MonitorProjectTimeEntry> projectTimes,
        DailyActivityBreakdown? activityBreakdown = null, bool isInMeeting = false, string meetingApp = "")
    {
        if (!ConfigService.Instance.Config.CloudLoggingEnabled || string.IsNullOrEmpty(BaseUrl))
            return false;

        try
        {
            // Build OpenProjects list with project names and versions
            var openProjects = revitProcesses
                .Where(p => !string.IsNullOrEmpty(p.ProjectName))
                .Select(p => new RevitProjectInfo
                { 
                    ProjectName = p.ProjectName, 
                    RevitVersion = p.RevitVersion ?? ""
                })
                .ToList();

            var projectDurations = projectTimes.Values
                .Select(p => new
                {
                    ProjectName = p.ProjectName,
                    RevitVersion = p.RevitVersion,
                    StartTime = p.StartTime,
                    Duration = p.DurationText,
                    TotalSeconds = p.CurrentDuration.TotalSeconds,
                    IsActive = p.IsActive
                })
                .ToList();

            // Get current project from revit processes or project times
            var currentProject = revitProcesses.FirstOrDefault(p => !string.IsNullOrEmpty(p.ProjectName))?.ProjectName ?? "";
            if (string.IsNullOrEmpty(currentProject) && projectTimes.Values.Any(p => p.IsActive))
            {
                currentProject = projectTimes.Values.FirstOrDefault(p => p.IsActive)?.ProjectName ?? "";
            }

            var status = new
            {
                WindowsUser = windowsUser,
                WindowsDisplayName = displayName,
                MachineName = machineName,
                AutodeskEmail = autodeskEmail,
                IsLoggedIn = isLoggedIn,
                Country = country,
                Office = office,
                ClientVersion = version,
                LastUpdate = Services.UserCredentialService.NowDubai,
                // Real-time activity data
                ActivityState = activityState,
                IdleSeconds = idleSeconds,
                RevitSessionCount = revitProcesses.Count > 0 ? revitProcesses.Count : (openProjects.Count > 0 ? openProjects.Count : 0),
                OpenProjects = openProjects,
                ProjectDurations = projectDurations,
                CurrentProject = currentProject,
                RevitVersion = revitProcesses.FirstOrDefault()?.RevitVersion ?? "",
                // Activity breakdown
                TodayRevitHours = activityBreakdown?.RevitHours ?? 0,
                TodayMeetingHours = activityBreakdown?.MeetingHours ?? 0,
                TodayIdleHours = activityBreakdown?.IdleHours ?? 0,
                TodayOtherHours = activityBreakdown?.OtherHours ?? 0,
                TodayOvertimeHours = Math.Max(0, (activityBreakdown?.TotalWorkHours ?? 0) - 9.0),
                TodayTotalHours = activityBreakdown?.TotalWorkHours ?? 0,
                IsInMeeting = isInMeeting,
                MeetingApp = meetingApp,
                // Include hourly breakdown for server sync
                HourlyBreakdown = GetHourlyBreakdownForSync(activityBreakdown)
            };

            var response = await _httpClient.PostAsJsonAsync($"{BaseUrl}/api/status", status);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }
    
    /// <summary>
    /// Helper to get hourly breakdown for server sync
    /// </summary>
    private static List<HourlyBreakdownDto> GetHourlyBreakdownForSync(DailyActivityBreakdown? breakdown)
    {
        if (breakdown?.HourlyBreakdown == null)
            return new List<HourlyBreakdownDto>();
            
        return breakdown.HourlyBreakdown.Select(h => new HourlyBreakdownDto
        {
            Hour = h.Hour,
            RevitMinutes = Math.Round(h.RevitMinutes, 1),
            MeetingMinutes = Math.Round(h.MeetingMinutes, 1),
            IdleMinutes = Math.Round(h.IdleMinutes, 1),
            OtherMinutes = Math.Round(h.OtherMinutes, 1)
        }).ToList();
    }
    
    /// <summary>
    /// Get all users' activity summaries from server (admin only)
    /// </summary>
    public async Task<List<UserActivitySummary>> GetAllUsersActivityAsync()
    {
        if (string.IsNullOrEmpty(BaseUrl))
            return new List<UserActivitySummary>();

        try
        {
            var response = await _httpClient.GetAsync($"{BaseUrl}/api/activity/all");
            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync();
                var options = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                return System.Text.Json.JsonSerializer.Deserialize<List<UserActivitySummary>>(json, options) 
                    ?? new List<UserActivitySummary>();
            }
        }
        catch { }
        
        return new List<UserActivitySummary>();
    }
    
    /// <summary>
    /// Get activity data for a specific date (historical or current)
    /// </summary>
    public async Task<(List<UserActivitySummary> Activities, bool IsLive, string Message)> GetActivityByDateAsync(DateTime date)
    {
        if (string.IsNullOrEmpty(BaseUrl))
            return (new List<UserActivitySummary>(), false, "Server not configured");

        try
        {
            var dateStr = date.ToString("yyyy-MM-dd");
            var response = await _httpClient.GetAsync($"{BaseUrl}/api/activity/date/{dateStr}");
            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync();
                var options = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                var root = doc.RootElement;
                
                var isLive = root.TryGetProperty("isLive", out var liveEl) && liveEl.GetBoolean();
                var message = root.TryGetProperty("message", out var msgEl) ? msgEl.GetString() ?? "" : "";
                
                var activities = new List<UserActivitySummary>();
                if (root.TryGetProperty("activities", out var activitiesEl))
                {
                    activities = System.Text.Json.JsonSerializer.Deserialize<List<UserActivitySummary>>(
                        activitiesEl.GetRawText(), options) ?? new List<UserActivitySummary>();
                }
                
                return (activities, isLive, message);
            }
        }
        catch (Exception ex)
        {
            return (new List<UserActivitySummary>(), false, $"Error: {ex.Message}");
        }
        
        return (new List<UserActivitySummary>(), false, "Failed to fetch data");
    }

    /// <summary>
    /// Send Revit project status (for Revit add-in integration)
    /// </summary>
    public async Task<bool> SendRevitStatusAsync(string windowsUser, string machineName, 
        string revitVersion, int sessionCount, List<RevitProjectInfo> openProjects)
    {
        if (!ConfigService.Instance.Config.CloudLoggingEnabled || string.IsNullOrEmpty(BaseUrl))
            return false;

        try
        {
            var status = new
            {
                WindowsUser = windowsUser,
                MachineName = machineName,
                RevitVersion = revitVersion,
                RevitSessionCount = sessionCount,
                OpenProjects = openProjects,
                CurrentProject = openProjects.FirstOrDefault()?.ProjectName ?? ""
            };

            var response = await _httpClient.PostAsJsonAsync($"{BaseUrl}/api/revit/status", status);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    public async Task<List<CloudSessionInfo>> GetSessionsAsync()
    {
        if (string.IsNullOrEmpty(BaseUrl))
            return new List<CloudSessionInfo>();

        try
        {
            var response = await _httpClient.GetAsync($"{BaseUrl}/api/sessions");
            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync();
                return JsonSerializer.Deserialize<List<CloudSessionInfo>>(json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new();
            }
        }
        catch { }

        return new List<CloudSessionInfo>();
    }

    public async Task<List<EmailUsageSummary>> GetEmailUsageSummaryAsync()
    {
        if (string.IsNullOrEmpty(BaseUrl))
            return new List<EmailUsageSummary>();

        try
        {
            var response = await _httpClient.GetAsync($"{BaseUrl}/api/admin/email-usage");
            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadFromJsonAsync<List<EmailUsageSummary>>() ?? new();
            }
        }
        catch { }

        return new List<EmailUsageSummary>();
    }

    /// <summary>
    /// Calculate email usage from sessions data (client-side fallback)
    /// </summary>
    public List<EmailUsageSummary> CalculateEmailUsageFromSessions(List<CloudSessionInfo> sessions, 
        string? currentUserEmail, string currentUserName, string currentMachine)
    {
        var result = new List<EmailUsageSummary>();
        
        // Group sessions by email (only logged-in users with valid emails)
        var emailGroups = sessions
            .Where(s => s.IsLoggedIn && !string.IsNullOrEmpty(s.AutodeskEmail) && s.AutodeskEmail.Contains("@"))
            .GroupBy(s => s.AutodeskEmail.ToLower())
            .ToList();

        // Also consider current user
        if (!string.IsNullOrEmpty(currentUserEmail) && currentUserEmail.Contains("@"))
        {
            var currentEmailLower = currentUserEmail.ToLower();
            var existingGroup = emailGroups.FirstOrDefault(g => g.Key == currentEmailLower);
            
            // Check if current user is already in the group
            bool currentUserInGroup = existingGroup?.Any(s => 
                s.MachineName.Equals(currentMachine, StringComparison.OrdinalIgnoreCase)) ?? false;

            if (existingGroup != null && !currentUserInGroup)
            {
                // Add current user to existing group count
                var summary = new EmailUsageSummary
                {
                    AutodeskEmail = currentUserEmail,
                    UserCount = existingGroup.Count() + 1,
                    UserNames = existingGroup.Select(s => s.GetDisplayName()).Append(currentUserName).Distinct().ToList(),
                    MachineNames = existingGroup.Select(s => s.MachineName).Append(currentMachine).Distinct().ToList()
                };
                result.Add(summary);
            }
        }

        foreach (var group in emailGroups)
        {
            // Skip if already added with current user
            if (result.Any(r => r.AutodeskEmail.Equals(group.Key, StringComparison.OrdinalIgnoreCase)))
                continue;

            var sessions_list = group.ToList();
            
            // Count unique users (by machine + windows user combination)
            var uniqueUsers = sessions_list
                .Select(s => $"{s.MachineName}|{s.WindowsUser}".ToLower())
                .Distinct()
                .Count();

            if (uniqueUsers > 1 || sessions_list.Count > 1)
            {
                result.Add(new EmailUsageSummary
                {
                    AutodeskEmail = group.First().AutodeskEmail,
                    UserCount = Math.Max(uniqueUsers, sessions_list.Count),
                    UserNames = sessions_list.Select(s => s.GetDisplayName()).Distinct().ToList(),
                    MachineNames = sessions_list.Select(s => s.MachineName).Distinct().ToList()
                });
            }
        }

        return result;
    }
    
    /// <summary>
    /// Delete historical activity data for a specific date
    /// </summary>
    public async Task<bool> DeleteActivityDataAsync(DateTime date)
    {
        if (string.IsNullOrEmpty(BaseUrl))
            return false;

        try
        {
            var dateStr = date.ToString("yyyy-MM-dd");
            var request = new HttpRequestMessage(HttpMethod.Delete, $"{BaseUrl}/api/activity/delete/{dateStr}");
            var response = await _httpClient.SendAsync(request);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }
    
    /// <summary>
    /// Get historical activity data for a date range
    /// </summary>
    public async Task<List<DailyActivityRecord>> GetHistoricalActivityAsync(DateTime startDate, DateTime endDate)
    {
        if (string.IsNullOrEmpty(BaseUrl))
            return new List<DailyActivityRecord>();

        try
        {
            var response = await _httpClient.GetAsync(
                $"{BaseUrl}/api/activity/history?startDate={startDate:yyyy-MM-dd}&endDate={endDate:yyyy-MM-dd}");
            
            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync();
                var options = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                return System.Text.Json.JsonSerializer.Deserialize<List<DailyActivityRecord>>(json, options) 
                    ?? new List<DailyActivityRecord>();
            }
        }
        catch { }
        
        return new List<DailyActivityRecord>();
    }
    
    /// <summary>
    /// Export activity report as CSV
    /// </summary>
    public async Task<string?> ExportReportAsync(DateTime startDate, DateTime endDate, string reportType = "all")
    {
        if (string.IsNullOrEmpty(BaseUrl))
            return null;

        try
        {
            var response = await _httpClient.GetAsync(
                $"{BaseUrl}/api/export/excel?startDate={startDate:yyyy-MM-dd}&endDate={endDate:yyyy-MM-dd}&reportType={reportType}");
            
            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadAsStringAsync();
            }
        }
        catch { }
        
        return null;
    }
}

/// <summary>
/// DTO for hourly breakdown data transfer
/// </summary>
public class HourlyBreakdownDto
{
    public int Hour { get; set; }
    public double RevitMinutes { get; set; }
    public double MeetingMinutes { get; set; }
    public double IdleMinutes { get; set; }
    public double OtherMinutes { get; set; }
}
