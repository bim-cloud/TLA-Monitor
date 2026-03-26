using System.Windows;
using AutodeskIDMonitor.Services;

namespace AutodeskIDMonitor.Views;

public partial class UpdateNotificationWindow : Window
{
    private readonly UpdateService.UpdateAvailableEventArgs _updateInfo;
    private bool _isUpdating = false;

    public UpdateNotificationWindow(UpdateService.UpdateAvailableEventArgs updateInfo)
    {
        InitializeComponent();
        _updateInfo = updateInfo;
        
        // Set version info
        CurrentVersionText.Text = $"v{updateInfo.CurrentVersion}";
        NewVersionText.Text = $"v{updateInfo.LatestVersion}";
        VersionText.Text = $"Version {updateInfo.LatestVersion} is now available";
        
        // Set release notes
        if (!string.IsNullOrEmpty(updateInfo.ReleaseNotes))
        {
            ReleaseNotesText.Text = updateInfo.ReleaseNotes;
        }
        
        // Set file size
        if (updateInfo.FileSizeBytes > 0)
        {
            var sizeMB = updateInfo.FileSizeBytes / 1024.0 / 1024.0;
            FileSizeText.Text = $"Download size: ~{sizeMB:F1} MB";
        }
        else
        {
            FileSizeText.Text = "Download size: ~15 MB";
        }
        
        // If mandatory, disable remind later
        if (updateInfo.IsMandatory)
        {
            RemindLaterButton.IsEnabled = false;
            RemindLaterButton.Content = "Required";
            RemindLaterButton.ToolTip = "This update is required and cannot be skipped";
            Title = "Required Update Available";
        }
        
        // Subscribe to update events
        UpdateService.Instance.UpdateProgress += OnUpdateProgress;
        UpdateService.Instance.UpdateError += OnUpdateError;
    }

    private void OnUpdateProgress(object? sender, UpdateService.UpdateProgressEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            ProgressPanel.Visibility = Visibility.Visible;
            ProgressBar.Value = e.ProgressPercent;
            ProgressText.Text = e.Status;
            
            if (e.ProgressPercent >= 100)
            {
                ProgressText.Text = "Installing... Please wait.";
            }
        });
    }

    private void OnUpdateError(object? sender, string error)
    {
        Dispatcher.Invoke(() =>
        {
            _isUpdating = false;
            UpdateNowButton.IsEnabled = true;
            UpdateNowButton.Content = "⟳ Retry";
            RemindLaterButton.IsEnabled = !_updateInfo.IsMandatory;
            ProgressPanel.Visibility = Visibility.Collapsed;
            
            System.Windows.MessageBox.Show(
                $"Update failed:\n\n{error}\n\nPlease try again or download manually from:\nhttp://141.145.153.32:5000/",
                "Update Error",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        });
    }

    private async void UpdateNow_Click(object sender, RoutedEventArgs e)
    {
        if (_isUpdating) return;
        
        _isUpdating = true;
        UpdateNowButton.IsEnabled = false;
        UpdateNowButton.Content = "Updating...";
        RemindLaterButton.IsEnabled = false;
        ProgressPanel.Visibility = Visibility.Visible;
        ProgressText.Text = "Preparing download...";
        
        var result = await UpdateService.Instance.DownloadAndInstallUpdateAsync(_updateInfo.LatestVersion);
        
        // If successful, the app will close automatically
        // If failed, OnUpdateError will handle it
    }

    private void RemindLater_Click(object sender, RoutedEventArgs e)
    {
        // Remind again in 24 hours
        UpdateService.Instance.SetRemindLater(24);
        DialogResult = false;
        Close();
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        // Prevent closing during update
        if (_isUpdating)
        {
            e.Cancel = true;
            System.Windows.MessageBox.Show(
                "Update is in progress. Please wait for it to complete.",
                "Update in Progress",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        
        // If mandatory update and trying to close without updating
        if (_updateInfo.IsMandatory)
        {
            var result = System.Windows.MessageBox.Show(
                "This update is required to continue using ID Monitor.\n\nAre you sure you want to exit? The application will close.",
                "Required Update",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            
            if (result == MessageBoxResult.Yes)
            {
                // Unsubscribe from events
                UpdateService.Instance.UpdateProgress -= OnUpdateProgress;
                UpdateService.Instance.UpdateError -= OnUpdateError;
                
                // Shutdown the app
                System.Windows.Application.Current.Shutdown();
            }
            else
            {
                e.Cancel = true;
            }
            return;
        }
        
        // Unsubscribe from events
        UpdateService.Instance.UpdateProgress -= OnUpdateProgress;
        UpdateService.Instance.UpdateError -= OnUpdateError;
        
        base.OnClosing(e);
    }
}
