using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using AutodeskIDMonitor.Models;

namespace AutodeskIDMonitor.Controls;

public partial class MeetingStatusControl : UserControl
{
    private static readonly SolidColorBrush InMeetingBg = new(Color.FromRgb(243, 229, 245)); // Light purple
    private static readonly SolidColorBrush InMeetingBorder = new(Color.FromRgb(206, 147, 216)); // Purple border
    private static readonly SolidColorBrush InMeetingIcon = new(Color.FromRgb(156, 39, 176)); // Purple
    private static readonly SolidColorBrush NoMeetingBg = new(Color.FromRgb(248, 249, 250)); // Light gray
    private static readonly SolidColorBrush NoMeetingBorder = new(Color.FromRgb(233, 236, 239)); // Gray border
    private static readonly SolidColorBrush NoMeetingIcon = new(Color.FromRgb(173, 181, 189)); // Gray
    private static readonly SolidColorBrush AppRunningBg = new(Color.FromRgb(232, 245, 233)); // Light green
    private static readonly SolidColorBrush AppRunningBorder = new(Color.FromRgb(165, 214, 167)); // Green border
    private static readonly SolidColorBrush AppRunningIcon = new(Color.FromRgb(76, 175, 80)); // Green
    private static readonly SolidColorBrush DurationBgActive = new(Color.FromRgb(243, 229, 245)); // Light purple

    private DispatcherTimer? _pulseTimer;
    private bool _isPulsing = false;
    private double _pulsePhase = 0;

    public MeetingStatusControl()
    {
        InitializeComponent();
        
        // Create pulse animation timer
        _pulseTimer = new DispatcherTimer();
        _pulseTimer.Interval = TimeSpan.FromMilliseconds(50);
        _pulseTimer.Tick += PulseTimer_Tick;
    }

    private void PulseTimer_Tick(object? sender, EventArgs e)
    {
        _pulsePhase += 0.05;
        if (_pulsePhase > 1.0) _pulsePhase = 0;
        
        // Animate the pulse ring
        var scale = 1.0 + (_pulsePhase * 0.8);
        var opacity = 0.8 * (1.0 - _pulsePhase);
        
        if (PulseRing.RenderTransform is ScaleTransform st)
        {
            st.ScaleX = scale;
            st.ScaleY = scale;
        }
        PulseRing.Opacity = opacity;
    }

    public void UpdateStatus(List<MeetingAppInfo>? meetings)
    {
        if (meetings == null || !meetings.Any())
        {
            ShowNoMeeting();
            return;
        }

        var activeMeeting = meetings.FirstOrDefault(m => m.IsInActiveMeeting);
        var runningApp = meetings.FirstOrDefault();

        if (activeMeeting != null)
        {
            ShowInMeeting(activeMeeting);
        }
        else if (runningApp != null)
        {
            ShowAppRunning(runningApp);
        }
        else
        {
            ShowNoMeeting();
        }
    }

    private void ShowInMeeting(MeetingAppInfo meeting)
    {
        MainBorder.Background = InMeetingBg;
        MainBorder.BorderBrush = InMeetingBorder;
        IconBorder.Background = InMeetingIcon;
        IconText.Text = GetAppIcon(meeting.AppName);
        
        StatusText.Text = "● In Meeting";
        StatusText.Foreground = InMeetingIcon;
        
        AppNameText.Text = meeting.AppName;
        if (!string.IsNullOrEmpty(meeting.WindowTitle))
        {
            var cleanTitle = CleanWindowTitle(meeting.WindowTitle);
            if (!string.IsNullOrEmpty(cleanTitle))
            {
                AppNameText.Text += $" - {TruncateTitle(cleanTitle, 35)}";
            }
        }
        
        var duration = meeting.Duration;
        DurationText.Text = $"{(int)duration.TotalHours:D2}:{duration.Minutes:D2}:{duration.Seconds:D2}";
        DurationText.Foreground = InMeetingIcon;
        DurationLabel.Text = "Duration";
        DurationBorder.Background = DurationBgActive;
        
        // Start pulse animation
        StartPulseAnimation();
    }

    private void ShowAppRunning(MeetingAppInfo app)
    {
        MainBorder.Background = AppRunningBg;
        MainBorder.BorderBrush = AppRunningBorder;
        IconBorder.Background = AppRunningIcon;
        IconText.Text = GetAppIcon(app.AppName);
        
        StatusText.Text = $"○ {app.AppName} Running";
        StatusText.Foreground = AppRunningIcon;
        
        AppNameText.Text = "Not in an active call";
        
        DurationText.Text = "";
        DurationLabel.Text = "";
        DurationBorder.Background = Brushes.Transparent;
        
        // Stop pulse animation
        StopPulseAnimation();
    }

    private void ShowNoMeeting()
    {
        MainBorder.Background = NoMeetingBg;
        MainBorder.BorderBrush = NoMeetingBorder;
        IconBorder.Background = NoMeetingIcon;
        IconText.Text = "📞";
        
        StatusText.Text = "○ No active meetings";
        StatusText.Foreground = new SolidColorBrush(Color.FromRgb(108, 117, 125));
        
        AppNameText.Text = "Teams, Zoom, Webex, Slack";
        
        DurationText.Text = "";
        DurationLabel.Text = "";
        DurationBorder.Background = Brushes.Transparent;
        
        // Stop pulse animation
        StopPulseAnimation();
    }

    private void StartPulseAnimation()
    {
        if (!_isPulsing && _pulseTimer != null)
        {
            _isPulsing = true;
            _pulsePhase = 0;
            _pulseTimer.Start();
        }
    }

    private void StopPulseAnimation()
    {
        if (_isPulsing && _pulseTimer != null)
        {
            _isPulsing = false;
            _pulseTimer.Stop();
            PulseRing.Opacity = 0;
            if (PulseRing.RenderTransform is ScaleTransform st)
            {
                st.ScaleX = 1;
                st.ScaleY = 1;
            }
        }
    }

    private string GetAppIcon(string appName)
    {
        return appName.ToLower() switch
        {
            var s when s.Contains("teams") => "🟣",
            var s when s.Contains("zoom") => "🔵",
            var s when s.Contains("webex") => "🟢",
            var s when s.Contains("slack") => "🟡",
            var s when s.Contains("discord") => "🟣",
            var s when s.Contains("skype") => "🔵",
            _ => "📞"
        };
    }

    private string CleanWindowTitle(string title)
    {
        if (string.IsNullOrEmpty(title)) return "";
        
        // Remove common suffixes
        var cleanTitle = title
            .Replace(" | Microsoft Teams", "")
            .Replace(" - Microsoft Teams", "")
            .Replace("Microsoft Teams", "")
            .Trim();
        
        return cleanTitle;
    }

    private string TruncateTitle(string title, int maxLength)
    {
        if (string.IsNullOrEmpty(title)) return "";
        if (title.Length <= maxLength) return title;
        return title.Substring(0, maxLength - 3) + "...";
    }

    /// <summary>
    /// Simple update with just meeting status
    /// </summary>
    public void UpdateStatus(bool isInMeeting, string appName = "", TimeSpan duration = default)
    {
        if (isInMeeting)
        {
            var meeting = new MeetingAppInfo
            {
                AppName = appName,
                IsInActiveMeeting = true,
                StartTime = DateTime.Now - duration
            };
            ShowInMeeting(meeting);
        }
        else
        {
            ShowNoMeeting();
        }
    }
}
