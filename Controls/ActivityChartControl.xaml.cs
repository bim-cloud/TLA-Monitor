using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using AutodeskIDMonitor.Models;

namespace AutodeskIDMonitor.Controls;

public partial class ActivityChartControl : UserControl
{
    // Colors
    private static readonly SolidColorBrush RevitBrush = new(Color.FromRgb(0, 180, 216));      // Cyan
    private static readonly SolidColorBrush MeetingBrush = new(Color.FromRgb(156, 39, 176));   // Purple
    private static readonly SolidColorBrush IdleBrush = new(Color.FromRgb(189, 189, 189));     // Gray
    private static readonly SolidColorBrush OtherBrush = new(Color.FromRgb(158, 158, 158));    // Dark gray
    private static readonly SolidColorBrush OvertimeBrush = new(Color.FromRgb(230, 81, 0));    // Deep orange
    private static readonly SolidColorBrush WorkHourBgBrush = new(Color.FromRgb(232, 245, 233));  // Light green
    private static readonly SolidColorBrush LunchBreakBgBrush = new(Color.FromRgb(255, 243, 224)); // Light orange
    private static readonly SolidColorBrush NonWorkHourBgBrush = new(Color.FromRgb(250, 250, 250)); // Light gray
    
    // Event for hour click
    public event EventHandler<HourlyActivityClickEventArgs>? HourClicked;
    
    // Current user name
    private string _currentUserName = "Current User";

    public ActivityChartControl()
    {
        InitializeComponent();
    }
    
    public void SetUserName(string userName)
    {
        _currentUserName = userName;
    }

    public void UpdateChart(DailyActivityBreakdown breakdown)
    {
        if (breakdown == null) return;

        ChartGrid.Children.Clear();
        ChartGrid.RowDefinitions.Clear();
        ChartGrid.ColumnDefinitions.Clear();

        // Working hours: 8 AM to 6 PM (with 1-2 PM lunch break)
        int startHour = 7;  // Show from 7 AM
        int endHour = 19;   // Show until 7 PM
        int hourCount = endHour - startHour + 1;

        // Add Y-axis label column
        ChartGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(35) });
        
        // Add columns for each hour
        for (int i = 0; i < hourCount; i++)
        {
            ChartGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        }

        // Create rows: chart area, hour labels
        ChartGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Chart
        ChartGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(30) }); // Hour labels with type

        // Add Y-axis labels (60m, 30m, 0m)
        AddYAxisLabels();

        // Calculate stats
        double totalRevit = 0, totalMeeting = 0, totalIdle = 0, totalOvertime = 0;

        // Add bars for each hour
        for (int i = 0; i < hourCount; i++)
        {
            int hour = startHour + i;
            var hourData = breakdown.HourlyBreakdown?.FirstOrDefault(h => h.Hour == hour);

            // Determine hour type
            bool isWorkHour = hour >= 8 && hour < 18 && hour != 13; // 8AM-6PM except lunch
            bool isLunchBreak = hour == 13; // 1PM-2PM
            bool isOvertime = hour < 8 || hour >= 18; // Before 8AM or after 6PM
            
            // Background color based on hour type
            SolidColorBrush bgBrush;
            string hourType = "";
            
            if (isLunchBreak)
            {
                bgBrush = LunchBreakBgBrush;
                hourType = "LB";
            }
            else if (isWorkHour)
            {
                bgBrush = WorkHourBgBrush;
                hourType = "WH";
            }
            else
            {
                bgBrush = NonWorkHourBgBrush;
                hourType = isOvertime ? "OT" : "";
            }
            
            // Background for this hour
            var bgBorder = new Border
            {
                Background = bgBrush,
                Margin = new Thickness(1, 0, 1, 0),
                CornerRadius = new CornerRadius(4, 4, 0, 0)
            };
            Grid.SetColumn(bgBorder, i + 1);
            Grid.SetRow(bgBorder, 0);
            ChartGrid.Children.Add(bgBorder);

            // Add activity bar if there's data
            if (hourData != null)
            {
                var barContainer = new Grid { Margin = new Thickness(3, 8, 3, 8) };
                
                var stackPanel = new StackPanel
                {
                    Orientation = Orientation.Vertical,
                    VerticalAlignment = VerticalAlignment.Bottom
                };

                double maxMinutes = 60;

                // Calculate minutes (capped at 60)
                double revitMins = Math.Min(hourData.RevitMinutes, maxMinutes);
                double meetingMins = Math.Min(hourData.MeetingMinutes, maxMinutes);
                double idleMins = Math.Min(hourData.IdleMinutes, maxMinutes);
                double otherMins = Math.Min(hourData.OtherMinutes, maxMinutes);
                
                // Track overtime (work outside 8-6 PM or during lunch)
                bool isOvertimeHour = isOvertime || isLunchBreak;
                if (isOvertimeHour)
                {
                    totalOvertime += (revitMins + meetingMins) / 60.0;
                }

                // Revit bar (cyan, or orange if overtime)
                if (revitMins > 0)
                {
                    var brush = isOvertimeHour ? OvertimeBrush : RevitBrush;
                    var bar = CreateClickableBar(brush, revitMins, maxMinutes, "Revit", hourData);
                    stackPanel.Children.Insert(0, bar);
                    if (!isOvertimeHour) totalRevit += revitMins / 60.0;
                }

                // Meeting bar (purple - always productive)
                if (meetingMins > 0)
                {
                    var brush = isOvertimeHour ? OvertimeBrush : MeetingBrush;
                    var bar = CreateClickableBar(brush, meetingMins, maxMinutes, "Meeting", hourData);
                    stackPanel.Children.Insert(0, bar);
                    if (!isOvertimeHour) totalMeeting += meetingMins / 60.0;
                }

                // Idle bar (gray)
                if (idleMins > 0)
                {
                    var bar = CreateClickableBar(IdleBrush, idleMins, maxMinutes, "Idle", hourData);
                    stackPanel.Children.Insert(0, bar);
                    if (!isOvertimeHour) totalIdle += idleMins / 60.0;
                }
                
                // Other bar
                if (otherMins > 0)
                {
                    var bar = CreateClickableBar(OtherBrush, otherMins, maxMinutes, "Other", hourData);
                    stackPanel.Children.Insert(0, bar);
                }

                barContainer.Children.Add(stackPanel);
                Grid.SetColumn(barContainer, i + 1);
                Grid.SetRow(barContainer, 0);
                ChartGrid.Children.Add(barContainer);
            }

            // Hour label with type indicator
            var labelPanel = new StackPanel 
            { 
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            
            // Time label (8A, 9A, etc.)
            var timeLabel = new TextBlock
            {
                Text = FormatHour(hour),
                FontSize = 10,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Color.FromRgb(80, 80, 80)),
                HorizontalAlignment = HorizontalAlignment.Center
            };
            labelPanel.Children.Add(timeLabel);
            
            // Type indicator (WH, LB, OT)
            if (!string.IsNullOrEmpty(hourType))
            {
                var typeLabel = new TextBlock
                {
                    Text = hourType,
                    FontSize = 8,
                    FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Foreground = isLunchBreak ? new SolidColorBrush(Color.FromRgb(230, 81, 0)) 
                               : isWorkHour ? new SolidColorBrush(Color.FromRgb(46, 125, 50))
                               : new SolidColorBrush(Color.FromRgb(230, 81, 0))
                };
                labelPanel.Children.Add(typeLabel);
            }
            
            Grid.SetColumn(labelPanel, i + 1);
            Grid.SetRow(labelPanel, 1);
            ChartGrid.Children.Add(labelPanel);
        }

        // Update summary stats
        UpdateSummaryStats(totalRevit, totalMeeting, totalIdle, totalOvertime);
    }
    
    private Border CreateClickableBar(SolidColorBrush brush, double minutes, double maxMinutes, string category, HourlyActivity hourData)
    {
        // Progressive height based on actual minutes
        double heightRatio = minutes / maxMinutes;
        
        var bar = new Border
        {
            Background = brush,
            Height = heightRatio * 80, // Max height 80 pixels
            MinHeight = minutes > 0 ? 4 : 0,
            CornerRadius = new CornerRadius(3),
            Margin = new Thickness(0, 1, 0, 0),
            Cursor = System.Windows.Input.Cursors.Hand,
            ToolTip = $"{category}: {minutes:F0} min"
        };
        
        // Click handler
        bar.MouseLeftButtonUp += (s, e) =>
        {
            HourClicked?.Invoke(this, new HourlyActivityClickEventArgs
            {
                UserName = _currentUserName,
                HourLabel = FormatHour(hourData.Hour),
                HourlyData = hourData
            });
        };
        
        // Hover effect
        bar.MouseEnter += (s, e) => bar.Opacity = 0.8;
        bar.MouseLeave += (s, e) => bar.Opacity = 1.0;
        
        return bar;
    }
    
    private void AddYAxisLabels()
    {
        var labelContainer = new Grid { Margin = new Thickness(0, 8, 0, 8) };
        labelContainer.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        labelContainer.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        labelContainer.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        
        var label60 = new TextBlock { Text = "60m", FontSize = 9, Foreground = Brushes.Gray, VerticalAlignment = VerticalAlignment.Top };
        var label30 = new TextBlock { Text = "30m", FontSize = 9, Foreground = Brushes.Gray, VerticalAlignment = VerticalAlignment.Center };
        var label0 = new TextBlock { Text = "0m", FontSize = 9, Foreground = Brushes.Gray, VerticalAlignment = VerticalAlignment.Bottom };
        
        Grid.SetRow(label60, 0);
        Grid.SetRow(label30, 1);
        Grid.SetRow(label0, 2);
        
        labelContainer.Children.Add(label60);
        labelContainer.Children.Add(label30);
        labelContainer.Children.Add(label0);
        
        Grid.SetColumn(labelContainer, 0);
        Grid.SetRow(labelContainer, 0);
        ChartGrid.Children.Add(labelContainer);
    }
    
    private string FormatHour(int hour)
    {
        if (hour == 0) return "12A";
        if (hour < 12) return $"{hour}A";
        if (hour == 12) return "12P";
        return $"{hour - 12}P";
    }
    
    private void UpdateSummaryStats(double revitHours, double meetingHours, double idleHours, double overtimeHours)
    {
        try
        {
            double totalWork = revitHours + meetingHours;
            
            TotalHoursText.Text = $"{totalWork:F1}h";
            RevitHoursText.Text = $"{revitHours:F1}h";
            MeetingHoursText.Text = $"{meetingHours:F1}h";
            IdleHoursText.Text = $"{idleHours:F1}h";
            
            // Update overtime if element exists
            if (OvertimeHoursText != null)
            {
                OvertimeHoursText.Text = $"{overtimeHours:F1}h";
            }
        }
        catch { }
    }
}

/// <summary>
/// Event args for hourly activity click
/// </summary>
public class HourlyActivityClickEventArgs : EventArgs
{
    public string UserName { get; set; } = "";
    public string HourLabel { get; set; } = "";
    public HourlyActivity? HourlyData { get; set; }
}
