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
            var userObj = r.TryGetProperty("users", out var u) && u.ValueKind == JsonValueKind.Object ? u : default;
            summaries.Add(new UserActivitySummary
            {
                MachineName = r.GetProperty("machine_name").GetString() ?? "",
                WindowsDisplayName = userObj.ValueKind == JsonValueKind.Object
                    ? userObj.GetProperty("display_name").GetString() ?? "" : "",
                Date = date,
                RevitHours = r.GetProperty("revit_hours").GetDouble(),
                MeetingHours = r.GetProperty("meeting_hours").GetDouble(),
                IdleHours = r.GetProperty("idle_hours").GetDouble(),
                OtherHours = r.GetProperty("other_hours").GetDouble(),
                OvertimeHours = r.GetProperty("overtime_hours").GetDouble(),
                TotalHours = r.GetProperty("total_hours").GetDouble(),
            });
        }
        return (summaries, false, $"Historical data for {dateStr}");
    }
    catch (Exception ex)
    {
        return (new List<UserActivitySummary>(), false, $"Error: {ex.Message}");
    }
}
