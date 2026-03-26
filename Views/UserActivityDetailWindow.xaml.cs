using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using AutodeskIDMonitor.Models;
using AutodeskIDMonitor.Services;

namespace AutodeskIDMonitor.Views;

public partial class UserActivityDetailWindow : Window
{
    private readonly UserActivitySummary _user;
    private readonly DailyActivityBreakdown? _breakdown;
    private readonly ExcelExportService _excelExport;
    private readonly bool _isCurrentUser;
    
    public UserActivityDetailWindow(UserActivitySummary user, DailyActivityBreakdown? localBreakdown = null, bool isCurrentUser = false)
    {
        InitializeComponent();
        _user = user;
        _breakdown = localBreakdown;
        _isCurrentUser = isCurrentUser;
        _excelExport = new ExcelExportService();
        
        LoadUserData();
    }
    
    private void LoadUserData()
    {
        // Header
        var names = _user.WindowsDisplayName.Split(' ');
        var initials = names.Length >= 2 
            ? $"{names[0][0]}{names[1][0]}" 
            : _user.WindowsDisplayName.Length >= 2 
                ? _user.WindowsDisplayName.Substring(0, 2).ToUpper()
                : _user.WindowsDisplayName.ToUpper();
        
        UserInitials.Text = initials;
        UserNameText.Text = _user.WindowsDisplayName;
        DateText.Text = _isCurrentUser 
            ? "📍 Your Activity - Live Updates" 
            : $"Activity for {_user.Date:MMMM dd, yyyy}";
        
        // Status
        UpdateStatus();
        
        // Activity Summary - Use LOCAL breakdown if available for current user
        double revitHours, meetingHours, idleHours, totalHours;
        
        if (_isCurrentUser && _breakdown != null)
        {
            // Use LIVE local data for current user
            revitHours = _breakdown.RevitHours;
            meetingHours = _breakdown.MeetingHours;
            idleHours = _breakdown.IdleHours;
            totalHours = _breakdown.TotalWorkHours;
        }
        else
        {
            // Use user data (from cloud/history)
            revitHours = _user.RevitHours;
            meetingHours = _user.MeetingHours;
            idleHours = _user.IdleHours;
            totalHours = _user.TotalHours;
        }
        
        RevitHoursText.Text = $"{revitHours:F1}h";
        MeetingHoursText.Text = $"{meetingHours:F1}h";
        IdleHoursText.Text = $"{idleHours:F1}h";
        TotalHoursText.Text = $"{totalHours:F1}h";
        
        // Project Breakdown
        LoadProjectBreakdown();
        
        // Hourly Breakdown
        LoadHourlyBreakdown();
        
        // Last updated
        LastUpdatedText.Text = _isCurrentUser 
            ? $"Live data • Last updated: {DateTime.Now:HH:mm:ss}"
            : $"Data from: {_user.Date:yyyy-MM-dd}";
    }
    
    private void UpdateStatus()
    {
        string statusText;
        SolidColorBrush statusBg;
        
        if (_user.IsInMeeting)
        {
            statusText = $"🟣 In Meeting";
            statusBg = new SolidColorBrush(Color.FromRgb(156, 39, 176));
        }
        else if (_user.ActivityState == "Idle")
        {
            statusText = "🟡 Idle";
            statusBg = new SolidColorBrush(Color.FromRgb(255, 152, 0));
        }
        else if (_user.ActivityState == "Active" || _isCurrentUser)
        {
            statusText = "🟢 Active";
            statusBg = new SolidColorBrush(Color.FromRgb(76, 175, 80));
        }
        else
        {
            statusText = "⚫ Offline";
            statusBg = new SolidColorBrush(Color.FromRgb(117, 117, 117));
        }
        
        StatusText.Text = statusText;
        StatusBadge.Background = statusBg;
    }
    
    private void LoadProjectBreakdown()
    {
        var projects = new List<ProjectDisplayItem>();
        
        // Get project breakdowns from user or breakdown
        if (_user.ProjectBreakdowns.Count > 0)
        {
            foreach (var p in _user.ProjectBreakdowns.OrderByDescending(p => p.ActiveRevitHours))
            {
                projects.Add(new ProjectDisplayItem
                {
                    ProjectName = p.ProjectName,
                    ActiveTimeDisplay = $"{p.ActiveRevitHours:F1}h",
                    StatusDisplay = p.IsCurrentlyActive ? "● Active" : "○ Inactive",
                    IsCurrentlyActive = p.IsCurrentlyActive
                });
            }
        }
        else if (!string.IsNullOrEmpty(_user.CurrentProject))
        {
            // Single project from current project field
            projects.Add(new ProjectDisplayItem
            {
                ProjectName = _user.CurrentProject,
                ActiveTimeDisplay = $"{_user.RevitHours:F1}h",
                StatusDisplay = "● Active",
                IsCurrentlyActive = true
            });
        }
        
        if (projects.Count > 0)
        {
            ProjectList.ItemsSource = projects;
            NoProjectsText.Visibility = Visibility.Collapsed;
        }
        else
        {
            NoProjectsText.Visibility = Visibility.Visible;
        }
    }
    
    private void LoadHourlyBreakdown()
    {
        var hourlyItems = new List<HourlyDisplayItem>();
        var hourlyData = _breakdown?.HourlyBreakdown ?? new List<HourlyActivity>();
        
        var currentHour = DateTime.Now.Hour;
        
        for (int hour = 7; hour <= Math.Max(currentHour, 18); hour++)
        {
            var data = hourlyData.FirstOrDefault(h => h.Hour == hour);
            var total = data != null ? (data.RevitMinutes + data.MeetingMinutes) / 60.0 : 0;
            var hasActivity = total > 0.01;
            
            hourlyItems.Add(new HourlyDisplayItem
            {
                HourLabel = $"{(hour > 12 ? hour - 12 : hour)}{(hour < 12 ? "A" : "P")}",
                TotalDisplay = hasActivity ? $"{total:F1}h" : "-",
                BackgroundColor = hasActivity 
                    ? new SolidColorBrush(Color.FromRgb(225, 245, 254))
                    : new SolidColorBrush(Color.FromRgb(250, 250, 250)),
                TextColor = hasActivity
                    ? new SolidColorBrush(Color.FromRgb(21, 101, 192))
                    : new SolidColorBrush(Color.FromRgb(189, 189, 189))
            });
        }
        
        HourlyList.ItemsSource = hourlyItems;
    }
    
    private void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel Files|*.xlsx|All Files|*.*",
            FileName = $"Activity_{_user.WindowsDisplayName.Replace(" ", "_")}_{DateTime.Now:yyyy-MM-dd_HHmm}.xlsx",
            Title = "Export User Activity"
        };
        
        if (dialog.ShowDialog() != true) return;
        
        try
        {
            _excelExport.ExportUserActivity(
                _user.WindowsDisplayName,
                _user.WindowsUser,
                _user.Date,
                _breakdown,
                null,
                dialog.FileName);
            
            var result = MessageBox.Show(
                $"Activity exported successfully!\n\nFile: {dialog.FileName}\n\nWould you like to open it?",
                "Export Complete",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);
            
            if (result == MessageBoxResult.Yes)
            {
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    
    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}

// Helper classes for display binding
public class ProjectDisplayItem
{
    public string ProjectName { get; set; } = "";
    public string ActiveTimeDisplay { get; set; } = "0h";
    public string StatusDisplay { get; set; } = "";
    public bool IsCurrentlyActive { get; set; }
    public SolidColorBrush StatusBackground => IsCurrentlyActive 
        ? new SolidColorBrush(Color.FromRgb(76, 175, 80))  // Green for active
        : new SolidColorBrush(Color.FromRgb(158, 158, 158)); // Gray for inactive
}

public class HourlyDisplayItem
{
    public string HourLabel { get; set; } = "";
    public string TotalDisplay { get; set; } = "-";
    public SolidColorBrush BackgroundColor { get; set; } = Brushes.Transparent;
    public SolidColorBrush TextColor { get; set; } = Brushes.Black;
}
