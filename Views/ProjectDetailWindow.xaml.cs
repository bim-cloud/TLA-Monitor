using System.Windows;
using AutodeskIDMonitor.Models;

namespace AutodeskIDMonitor.Views;

public partial class ProjectDetailWindow : Window
{
    public ProjectDetailWindow(ProjectWorkSummary project)
    {
        InitializeComponent();
        LoadProjectDetails(project);
    }
    
    private void LoadProjectDetails(ProjectWorkSummary project)
    {
        if (project == null) return;
        
        // Set header info
        ProjectNameText.Text = project.ProjectName ?? "Unknown Project";
        TotalHoursText.Text = $"{project.TotalHours:F1}h";
        
        // Set summary stats
        UserCountText.Text = project.UserCount.ToString();
        LastActivityText.Text = project.LastActivity.ToString("HH:mm");
        
        double avgHours = project.UserCount > 0 ? project.TotalHours / project.UserCount : 0;
        AvgHoursText.Text = $"{avgHours:F1}h";
        
        // Calculate contribution percentages and bind to grid
        if (project.UserTimes != null && project.UserTimes.Count > 0)
        {
            var userBreakdownList = project.UserTimes
                .Select(ut => new UserProjectTimeDisplay
                {
                    UserName = ut.UserName ?? "Unknown",
                    UserId = ut.UserId ?? "",
                    Hours = ut.Hours,
                    LastActivity = ut.LastActivity,
                    ContributionPercent = project.TotalHours > 0 ? (ut.Hours / project.TotalHours) * 100 : 0
                })
                .OrderByDescending(u => u.Hours)
                .ToList();
            
            UserBreakdownGrid.ItemsSource = userBreakdownList;
        }
    }
    
    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}

// Display model with contribution percentage
public class UserProjectTimeDisplay
{
    public string UserName { get; set; } = "";
    public string UserId { get; set; } = "";
    public double Hours { get; set; }
    public DateTime LastActivity { get; set; }
    public double ContributionPercent { get; set; }
}
