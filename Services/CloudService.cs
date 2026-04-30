using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using AutodeskIDMonitor.Models;

namespace AutodeskIDMonitor.Services;

/// <summary>
/// CloudService v5 — talks directly to Supabase (PostgREST) instead of the
/// old Oracle Cloud Flask server. All write paths go through a single
/// `upsert_activity` RPC so one heartbeat = one HTTP call.
///
/// Configuration (AppConfig):
///   CloudApiUrl   -> https://jdfzpnreoitpdhttielk.supabase.co
///   CloudApiKey   -> Supabase anon key (or, ideally, a per-machine signed key)
/// </summary>
public class CloudService
{
    private readonly HttpClient _httpClient;

    public CloudService()
    {
        _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        UpdateApiKey();
    }

    public void UpdateApiKey()
    {
        _httpClient.DefaultRequestHeaders.Remove("apikey");
        _httpClient.DefaultRequestHeaders.Remove("Authorization");
        var apiKey = ConfigService.Instance.Config.CloudApiKey;
        if (!string.IsNullOrEmpty(apiKey))
        {
            _httpClient.DefaultRequestHeaders.Add("apikey", apiKey);
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
        }
    }

    private string BaseUrl => ConfigService.Instance.Config.CloudApiUrl?.TrimEnd('/') ?? "";
    private string RpcUrl  => $"{BaseUrl}/rest/v1/rpc";
    private string RestUrl => $"{BaseUrl}/rest/v1";

    // -----------------------------------------------------------------------
    // Heartbeat — replaces the old POST /api/status. Single RPC call.
    // -----------------------------------------------------------------------
    public async Task<bool> SendActivityStatusAsync(
        string windowsUser, string displayName, string machineName,
        string autodeskEmail, bool isLoggedIn, string country, string office, string version,
        string activityState, int idleSeconds, List<RevitProcessInfo> revitProcesses,
        Dictionary<string, MonitorProjectTimeEntry> projectTimes,
        DailyActivityBreakdown? activityBreakdown = null,
        bool isInMeeting = false, string meetingApp = "")
    {
        if (!ConfigService.Instance.Config.CloudLoggingEnabled || string.IsNullOrEmpty(BaseUrl))
            return false;

        try
        {
            var openProjects = revitProcesses
                .Where(p => !string.IsNullOrEmpty(p.ProjectName))
                .Select(p => new { projectName = p.ProjectName, revitVersion = p.RevitVersion ?? "" })
                .ToList<object>();

            var projectsWorked = projectTimes.Values
                .Select(p => new
                {
                    projectName  = p.ProjectName,
                    revitVersion = p.RevitVersion,
                    totalSeconds = p.CurrentDuration.TotalSeconds,
                    isActive     = p.IsActive
                })
                .ToList<object>();

            var currentProject =
                revitProcesses.FirstOrDefault(p => !string.IsNullOrEmpty(p.ProjectName))?.ProjectName
                ?? projectTimes.Values.FirstOrDefault(p => p.IsActive)?.ProjectName
                ?? "";

            var hourly = (activityBreakdown?.HourlyBreakdown ?? new List<HourlyActivity>())
                .Select(h => new
                {
                    hour          = h.Hour,
                    revitMinutes  = Math.Round(h.RevitMinutes, 1),
                    meetingMinutes= Math.Round(h.MeetingMinutes, 1),
                    idleMinutes   = Math.Round(h.IdleMinutes, 1),
                    otherMinutes  = Math.Round(h.OtherMinutes, 1)
                })
                .ToList<object>();

            // RPC payload — note the parameter names match the SQL function exactly
            var payload = new
            {
                p_machine_name        = machineName,
                p_windows_user        = windowsUser,
                p_autodesk_email      = autodeskEmail ?? "",
                p_is_logged_in        = isLoggedIn,
                p_activity_state      = activityState ?? "offline",
                p_idle_seconds        = idleSeconds,
                p_revit_session_count = revitProcesses.Count,
                p_current_project     = currentProject,
                p_revit_version       = revitProcesses.FirstOrDefault()?.RevitVersion ?? "",
                p_open_projects       = openProjects,
                p_is_in_meeting       = isInMeeting,
                p_meeting_app         = meetingApp ?? "",
                p_client_version      = version,
                p_revit_hours         = activityBreakdown?.RevitHours   ?? 0,
                p_meeting_hours       = activityBreakdown?.MeetingHours ?? 0,
                p_idle_hours          = activityBreakdown?.IdleHours    ?? 0,
                p_other_hours         = activityBreakdown?.OtherHours   ?? 0,
                p_overtime_hours      = Math.Max(0, (activityBreakdown?.TotalWorkHours ?? 0) - 9.0),
                p_total_hours         = activityBreakdown?.TotalWorkHours ?? 0,
                p_hourly_breakdown    = hourly,
                p_projects_worked     = projectsWorked
            };

            var response = await _httpClient.PostAsJsonAsync($"{RpcUrl}/upsert_activity", payload);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Backwards-compat shim — the old CloudService had a separate
    /// SendStatusAsync overload. Forward it to the unified call so older
    /// callers keep working without changes.
    /// </summary>
    public Task<bool> SendStatusAsync(string windowsUser, string displayName, string machineName,
        string autodeskEmail, bool isLoggedIn, string country, string office, string version)
    {
        return SendActivityStatusAsync(windowsUser, displayName, machineName, autodeskEmail,
            isLoggedIn, country, office, version,
            activityState: isLoggedIn ? "active" : "offline",
            idleSeconds: 0,
            revitProcesses: new List<RevitProcessInfo>(),
            projectTimes: new Dictionary<string, MonitorProjectTimeEntry>());
    }

    // -----------------------------------------------------------------------
    // Reads — now query the v_live_users view directly.
    // -----------------------------------------------------------------------
    public async Task<List<CloudSessionInfo>> GetSessionsAsync()
    {
        if (string.IsNullOrEmpty(BaseUrl)) return new List<CloudSessionInfo>();
        try
        {
            var url = $"{RestUrl}/v_live_users?select=*&order=last_update.desc";
            var response = await _httpClient.GetAsync(url);
            if (!response.IsSuccessStatusCode) return new List<CloudSessionInfo>();

            var json = await response.Content.ReadAsStringAsync();
            return ParseLiveUsers(json);
        }
        catch { return new List<CloudSessionInfo>(); }
    }

    public async Task<List<UserActivitySummary>> GetAllUsersActivityAsync()
    {
        var sessions = await GetSessionsAsync();
        return sessions.Select(SessionToActivitySummary).ToList();
    }

    public async Task<(List<UserActivitySummary> Activities, bool IsLive, string Message)>
        GetActivityByDateAsync(DateTime date)
    {
        if (string.IsNullOrEmpty(BaseUrl))
            return (new List<UserActivitySummary>(), false, "Server not configured");

        var isToday = date.Date == DateTime.UtcNow.Date;
        if (isToday)
        {
            var live = await GetAllUsersActivityAsync();
            return (live, true, "Live data");
        }

        try
        {
            var dateStr = date.ToString("yyyy-MM-dd");
            var url = $"{RestUrl}/daily_activity?select=*,users(display_name,email)&date=eq.{dateStr}";
            var response = await _httpClient.GetAsync(url);
            if (!response.IsSuccessStatusCode)
                return (new List<UserActivitySummary>(), false, "No data for date");

            var json = await response.Content.ReadAsStringAsync();
            var rows = JsonDocument.Parse(json).RootElement;
            var summaries = new List<UserActivitySummary>();
            foreach (var r in rows.EnumerateArray())
            {
                summaries.Add(new UserActivitySummary
                {
                    MachineName    = r.GetProperty("machine_name").GetString() ?? "",
                    DisplayName    = r.TryGetProperty("users", out var u) && u.ValueKind == JsonValueKind.Object
                                     ? u.GetProperty("display_name").GetString() ?? ""
                                     : "",
                    AutodeskEmail  = r.TryGetProperty("users", out var u2) && u2.ValueKind == JsonValueKind.Object
                                     ? u2.GetProperty("email").GetString() ?? ""
                                     : "",
                    TodayRevitHours    = r.GetProperty("revit_hours").GetDouble(),
                    TodayMeetingHours  = r.GetProperty("meeting_hours").GetDouble(),
                    TodayIdleHours     = r.GetProperty("idle_hours").GetDouble(),
                    TodayOtherHours    = r.GetProperty("other_hours").GetDouble(),
                    TodayOvertimeHours = r.GetProperty("overtime_hours").GetDouble(),
                    TodayTotalHours    = r.GetProperty("total_hours").GetDouble(),
                });
            }
            return (summaries, false, $"Historical data for {dateStr}");
        }
        catch (Exception ex)
        {
            return (new List<UserActivitySummary>(), false, $"Error: {ex.Message}");
        }
    }

    public async Task<List<DailyActivityRecord>> GetHistoricalActivityAsync(DateTime startDate, DateTime endDate)
    {
        if (string.IsNullOrEmpty(BaseUrl)) return new List<DailyActivityRecord>();
        try
        {
            var url = $"{RestUrl}/daily_activity?select=*&date=gte.{startDate:yyyy-MM-dd}&date=lte.{endDate:yyyy-MM-dd}";
            var response = await _httpClient.GetAsync(url);
            if (!response.IsSuccessStatusCode) return new List<DailyActivityRecord>();

            var json = await response.Content.ReadAsStringAsync();
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            return JsonSerializer.Deserialize<List<DailyActivityRecord>>(json, options)
                   ?? new List<DailyActivityRecord>();
        }
        catch { return new List<DailyActivityRecord>(); }
    }

    public async Task<bool> DeleteActivityDataAsync(DateTime date)
    {
        if (string.IsNullOrEmpty(BaseUrl)) return false;
        try
        {
            var url = $"{RestUrl}/daily_activity?date=eq.{date:yyyy-MM-dd}";
            var response = await _httpClient.DeleteAsync(url);
            return response.IsSuccessStatusCode;
        }
        catch { return false; }
    }

    // The Excel export endpoint is now generated by the dashboard side; we
    // keep the method here returning null for backwards compat so the WPF
    // code path doesn't break.
    public Task<string?> ExportReportAsync(DateTime _, DateTime __, string ___ = "all")
        => Task.FromResult<string?>(null);

    // The old SendRevitStatusAsync was a niche "Revit add-in only" hook that
    // also went through SendActivityStatusAsync for normal clients — we no
    // longer need a distinct endpoint.
    public Task<bool> SendRevitStatusAsync(string windowsUser, string machineName,
        string revitVersion, int sessionCount, List<RevitProjectInfo> openProjects)
        => Task.FromResult(false);

    public List<EmailUsageSummary> CalculateEmailUsageFromSessions(List<CloudSessionInfo> sessions,
        string? currentUserEmail, string currentUserName, string currentMachine)
    {
        // Behaviour preserved from v4 — duplicate-email detection works against
        // whatever GetSessionsAsync returns, which is now the Supabase view.
        var result = new List<EmailUsageSummary>();
        var emailGroups = sessions
            .Where(s => s.IsLoggedIn && !string.IsNullOrEmpty(s.AutodeskEmail) && s.AutodeskEmail.Contains("@"))
            .GroupBy(s => s.AutodeskEmail.ToLower())
            .ToList();

        foreach (var g in emailGroups)
        {
            var list = g.ToList();
            var unique = list.Select(s => $"{s.MachineName}|{s.WindowsUser}".ToLower()).Distinct().Count();
            if (unique > 1 || list.Count > 1)
            {
                result.Add(new EmailUsageSummary
                {
                    AutodeskEmail = g.First().AutodeskEmail,
                    UserCount     = Math.Max(unique, list.Count),
                    UserNames     = list.Select(s => s.GetDisplayName()).Distinct().ToList(),
                    MachineNames  = list.Select(s => s.MachineName).Distinct().ToList()
                });
            }
        }
        return result;
    }

    public Task<List<EmailUsageSummary>> GetEmailUsageSummaryAsync()
    {
        // Compute client-side from the live sessions list rather than calling
        // a dedicated endpoint.
        return GetSessionsAsync()
            .ContinueWith(t => CalculateEmailUsageFromSessions(t.Result, null, "", ""));
    }

    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------
    private static List<CloudSessionInfo> ParseLiveUsers(string json)
    {
        var list = new List<CloudSessionInfo>();
        using var doc = JsonDocument.Parse(json);
        foreach (var r in doc.RootElement.EnumerateArray())
        {
            var openProjects = new List<RevitProjectInfo>();
            if (r.TryGetProperty("open_projects", out var op) && op.ValueKind == JsonValueKind.Array)
            {
                foreach (var p in op.EnumerateArray())
                {
                    openProjects.Add(new RevitProjectInfo
                    {
                        ProjectName  = p.TryGetProperty("projectName", out var n) ? n.GetString() ?? "" : "",
                        RevitVersion = p.TryGetProperty("revitVersion", out var v) ? v.GetString() ?? "" : ""
                    });
                }
            }

            list.Add(new CloudSessionInfo
            {
                MachineName        = GetStr(r, "machine_name"),
                WindowsUser        = GetStr(r, "windows_user"),
                WindowsDisplayName = GetStr(r, "display_name"),
                DisplayName        = GetStr(r, "display_name"),
                AutodeskEmail      = GetStr(r, "autodesk_email"),
                IsLoggedIn         = GetBool(r, "is_logged_in"),
                LastSeen           = GetDateTime(r, "last_update"),
                LastUpdate         = GetDateTime(r, "last_update"),
                ClientVersion      = GetStr(r, "client_version"),
                RevitVersion       = GetStr(r, "revit_version"),
                CurrentProject     = GetStr(r, "current_project"),
                OpenProjects       = openProjects,
                RevitSessionCount  = GetInt(r, "revit_session_count"),
                ActivityState      = GetStr(r, "activity_state"),
                IdleSeconds        = GetInt(r, "idle_seconds"),
                IsInMeeting        = GetBool(r, "is_in_meeting"),
                MeetingApp         = GetStr(r, "meeting_app"),
                TodayRevitHours    = GetDouble(r, "today_revit_hours"),
                TodayMeetingHours  = GetDouble(r, "today_meeting_hours"),
                TodayIdleHours     = GetDouble(r, "today_idle_hours"),
                TodayTotalHours    = GetDouble(r, "today_total_hours"),
                TodayOvertimeHours = GetDouble(r, "today_overtime_hours"),
            });
        }
        return list;
    }

    private static UserActivitySummary SessionToActivitySummary(CloudSessionInfo s) => new()
    {
        MachineName        = s.MachineName,
        DisplayName        = s.GetDisplayName(),
        AutodeskEmail      = s.AutodeskEmail,
        IsOnline           = s.IsLoggedIn,
        ActivityState      = s.ActivityState,
        CurrentProject     = s.CurrentProject,
        RevitVersion       = s.RevitVersion,
        TodayRevitHours    = s.TodayRevitHours,
        TodayMeetingHours  = s.TodayMeetingHours,
        TodayIdleHours     = s.TodayIdleHours,
        TodayTotalHours    = s.TodayTotalHours,
        TodayOvertimeHours = s.TodayOvertimeHours,
    };

    private static string GetStr(JsonElement e, string n) =>
        e.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";
    private static bool GetBool(JsonElement e, string n) =>
        e.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.True;
    private static int GetInt(JsonElement e, string n) =>
        e.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : 0;
    private static double GetDouble(JsonElement e, string n) =>
        e.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetDouble() : 0;
    private static DateTime GetDateTime(JsonElement e, string n) =>
        e.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String && DateTime.TryParse(v.GetString(), out var d) ? d : DateTime.MinValue;
}

/// <summary>Kept for backwards compatibility with MainViewModel.cs callers.</summary>
public class HourlyBreakdownDto
{
    public int Hour { get; set; }
    public double RevitMinutes { get; set; }
    public double MeetingMinutes { get; set; }
    public double IdleMinutes { get; set; }
    public double OtherMinutes { get; set; }
}
