using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace AutodeskIDMonitor.Controls;

public partial class ActivityPieChart : UserControl
{
    private static readonly Color[] ChartColors = new[]
    {
        Color.FromRgb(0, 180, 216),    // Revit - Cyan
        Color.FromRgb(156, 39, 176),   // Meeting - Purple
        Color.FromRgb(189, 189, 189),  // Idle - Gray
        Color.FromRgb(255, 152, 0),    // Other - Orange
        Color.FromRgb(76, 175, 80),    // Break - Green
    };

    public ActivityPieChart()
    {
        InitializeComponent();
    }

    public void UpdateChart(List<(string Label, double Hours, Color? CustomColor)> data)
    {
        PieCanvas.Children.Clear();
        LegendPanel.Children.Clear();

        if (data == null || data.Count == 0 || data.Sum(d => d.Hours) == 0)
        {
            // Show empty state
            var emptyText = new TextBlock
            {
                Text = "No data",
                Foreground = Brushes.Gray,
                FontSize = 14,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            Canvas.SetLeft(emptyText, 50);
            Canvas.SetTop(emptyText, 70);
            PieCanvas.Children.Add(emptyText);
            return;
        }

        double total = data.Sum(d => d.Hours);
        double centerX = 75;
        double centerY = 75;
        double radius = 70;
        double startAngle = -90; // Start from top

        int colorIndex = 0;

        foreach (var (label, hours, customColor) in data)
        {
            if (hours <= 0) continue;

            double percentage = hours / total;
            double sweepAngle = percentage * 360;

            // Get color
            var color = customColor ?? ChartColors[colorIndex % ChartColors.Length];
            colorIndex++;

            // Create animated pie slice
            var slice = CreatePieSlice(centerX, centerY, radius, startAngle, sweepAngle, color, label, hours);
            PieCanvas.Children.Add(slice);

            // Add legend item
            AddLegendItem(label, hours, percentage, color);

            startAngle += sweepAngle;
        }

        // Add center circle for donut effect
        var centerCircle = new Ellipse
        {
            Width = 50,
            Height = 50,
            Fill = Brushes.White
        };
        Canvas.SetLeft(centerCircle, centerX - 25);
        Canvas.SetTop(centerCircle, centerY - 25);
        PieCanvas.Children.Add(centerCircle);

        // Add total in center
        var totalText = new TextBlock
        {
            Text = $"{total:F1}h",
            FontWeight = FontWeights.Bold,
            FontSize = 14,
            Foreground = new SolidColorBrush(Color.FromRgb(51, 51, 51))
        };
        Canvas.SetLeft(totalText, centerX - 18);
        Canvas.SetTop(totalText, centerY - 10);
        PieCanvas.Children.Add(totalText);
    }

    private Path CreatePieSlice(double cx, double cy, double r, double startAngle, double sweepAngle, Color color, string label, double hours)
    {
        // Convert angles to radians
        double startRad = startAngle * Math.PI / 180;
        double endRad = (startAngle + sweepAngle) * Math.PI / 180;

        // Calculate start and end points
        double x1 = cx + r * Math.Cos(startRad);
        double y1 = cy + r * Math.Sin(startRad);
        double x2 = cx + r * Math.Cos(endRad);
        double y2 = cy + r * Math.Sin(endRad);

        bool isLargeArc = sweepAngle > 180;

        var pathFigure = new PathFigure
        {
            StartPoint = new Point(cx, cy),
            IsClosed = true
        };

        pathFigure.Segments.Add(new LineSegment(new Point(x1, y1), true));
        pathFigure.Segments.Add(new ArcSegment
        {
            Point = new Point(x2, y2),
            Size = new Size(r, r),
            IsLargeArc = isLargeArc,
            SweepDirection = SweepDirection.Clockwise
        });

        var pathGeometry = new PathGeometry();
        pathGeometry.Figures.Add(pathFigure);

        var path = new Path
        {
            Fill = new SolidColorBrush(color),
            Data = pathGeometry,
            Stroke = Brushes.White,
            StrokeThickness = 2,
            Cursor = System.Windows.Input.Cursors.Hand,
            ToolTip = new ToolTip 
            { 
                Content = $"{label}: {hours:F1} hours",
                Background = new SolidColorBrush(Color.FromRgb(33, 33, 33)),
                Foreground = Brushes.White,
                FontSize = 12,
                Padding = new Thickness(10, 6, 10, 6),
                BorderThickness = new Thickness(0)
            },
            RenderTransformOrigin = new Point(0.5, 0.5),
            Effect = new System.Windows.Media.Effects.DropShadowEffect
            {
                ShadowDepth = 0,
                BlurRadius = 0,
                Opacity = 0,
                Color = color
            }
        };
        
        // Composite transform for sophisticated animations
        var transformGroup = new TransformGroup();
        var scaleTransform = new ScaleTransform(1, 1);
        var translateTransform = new TranslateTransform(0, 0);
        transformGroup.Children.Add(scaleTransform);
        transformGroup.Children.Add(translateTransform);
        path.RenderTransform = transformGroup;
        
        // Calculate direction for "pop out" effect
        double midAngle = (startAngle + sweepAngle / 2) * Math.PI / 180;
        double popX = Math.Cos(midAngle) * 8;
        double popY = Math.Sin(midAngle) * 8;
        
        // Sophisticated hover animations
        path.MouseEnter += (s, e) =>
        {
            // Scale up with bounce
            var scaleAnim = new System.Windows.Media.Animation.DoubleAnimation(1.0, 1.1, TimeSpan.FromMilliseconds(200))
            {
                EasingFunction = new System.Windows.Media.Animation.BackEase { Amplitude = 0.3, EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
            };
            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnim);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnim);
            
            // Pop out in the direction of the slice
            var moveXAnim = new System.Windows.Media.Animation.DoubleAnimation(0, popX, TimeSpan.FromMilliseconds(200))
            {
                EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
            };
            var moveYAnim = new System.Windows.Media.Animation.DoubleAnimation(0, popY, TimeSpan.FromMilliseconds(200))
            {
                EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
            };
            translateTransform.BeginAnimation(TranslateTransform.XProperty, moveXAnim);
            translateTransform.BeginAnimation(TranslateTransform.YProperty, moveYAnim);
            
            // Glow effect
            var shadow = path.Effect as System.Windows.Media.Effects.DropShadowEffect;
            if (shadow != null)
            {
                var shadowAnim = new System.Windows.Media.Animation.DoubleAnimation(0, 0.5, TimeSpan.FromMilliseconds(200));
                var blurAnim = new System.Windows.Media.Animation.DoubleAnimation(0, 15, TimeSpan.FromMilliseconds(200));
                shadow.BeginAnimation(System.Windows.Media.Effects.DropShadowEffect.OpacityProperty, shadowAnim);
                shadow.BeginAnimation(System.Windows.Media.Effects.DropShadowEffect.BlurRadiusProperty, blurAnim);
            }
            
            path.StrokeThickness = 3;
            path.Opacity = 0.95;
        };
        
        path.MouseLeave += (s, e) =>
        {
            // Scale back smoothly
            var scaleAnim = new System.Windows.Media.Animation.DoubleAnimation(1.1, 1.0, TimeSpan.FromMilliseconds(250))
            {
                EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
            };
            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnim);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnim);
            
            // Move back to center
            var moveXAnim = new System.Windows.Media.Animation.DoubleAnimation(popX, 0, TimeSpan.FromMilliseconds(250))
            {
                EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
            };
            var moveYAnim = new System.Windows.Media.Animation.DoubleAnimation(popY, 0, TimeSpan.FromMilliseconds(250))
            {
                EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
            };
            translateTransform.BeginAnimation(TranslateTransform.XProperty, moveXAnim);
            translateTransform.BeginAnimation(TranslateTransform.YProperty, moveYAnim);
            
            // Remove glow
            var shadow = path.Effect as System.Windows.Media.Effects.DropShadowEffect;
            if (shadow != null)
            {
                var shadowAnim = new System.Windows.Media.Animation.DoubleAnimation(0.5, 0, TimeSpan.FromMilliseconds(250));
                var blurAnim = new System.Windows.Media.Animation.DoubleAnimation(15, 0, TimeSpan.FromMilliseconds(250));
                shadow.BeginAnimation(System.Windows.Media.Effects.DropShadowEffect.OpacityProperty, shadowAnim);
                shadow.BeginAnimation(System.Windows.Media.Effects.DropShadowEffect.BlurRadiusProperty, blurAnim);
            }
            
            path.StrokeThickness = 2;
            path.Opacity = 1.0;
        };
        
        // Click animation
        path.MouseLeftButtonDown += (s, e) =>
        {
            var clickAnim = new System.Windows.Media.Animation.DoubleAnimation(1.1, 0.95, TimeSpan.FromMilliseconds(80));
            clickAnim.AutoReverse = true;
            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, clickAnim);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, clickAnim);
        };

        return path;
    }

    // Keep original method for backward compatibility
    private Path CreatePieSlice(double cx, double cy, double r, double startAngle, double sweepAngle, Color color)
    {
        return CreatePieSlice(cx, cy, r, startAngle, sweepAngle, color, "", 0);
    }

    private void AddLegendItem(string label, double hours, double percentage, Color color)
    {
        var panel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Thickness(0, 3, 0, 3)
        };

        var colorBox = new Border
        {
            Width = 14,
            Height = 14,
            Background = new SolidColorBrush(color),
            CornerRadius = new CornerRadius(2),
            Margin = new Thickness(0, 0, 8, 0)
        };

        var labelText = new TextBlock
        {
            Text = $"{label}: {hours:F1}h ({percentage:P0})",
            FontSize = 11,
            Foreground = new SolidColorBrush(Color.FromRgb(102, 102, 102))
        };

        panel.Children.Add(colorBox);
        panel.Children.Add(labelText);
        LegendPanel.Children.Add(panel);
    }

    /// <summary>
    /// Simplified update method
    /// </summary>
    public void UpdateChart(double revitHours, double meetingHours, double idleHours, double otherHours = 0)
    {
        var data = new List<(string Label, double Hours, Color? CustomColor)>
        {
            ("Revit", revitHours, Color.FromRgb(0, 180, 216)),
            ("Meetings", meetingHours, Color.FromRgb(156, 39, 176)),
            ("Idle", idleHours, Color.FromRgb(189, 189, 189)),
        };

        if (otherHours > 0)
        {
            data.Add(("Other", otherHours, Color.FromRgb(255, 152, 0)));
        }

        UpdateChart(data);
    }
}
