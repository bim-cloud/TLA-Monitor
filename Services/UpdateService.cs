using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows;
using Microsoft.Win32;

namespace AutodeskIDMonitor.Services;

/// <summary>
/// Handles automatic updates and system lock detection
/// </summary>
public class UpdateService : IDisposable
{
    private static UpdateService? _instance;
    public static UpdateService Instance => _instance ??= new UpdateService();

    private readonly HttpClient _httpClient;
    private readonly string _serverUrl = "http://141.145.153.32:5000";
    private readonly string _updateFolder;
    private bool _updateInProgress = false;
    private bool _isSystemLocked = false;
    private System.Timers.Timer? _updateCheckTimer;

    // Events
    public event EventHandler<UpdateAvailableEventArgs>? UpdateAvailable;
    public event EventHandler<UpdateProgressEventArgs>? UpdateProgress;
    public event EventHandler<string>? UpdateError;
    public event EventHandler<bool>? SystemLockStateChanged;

    public bool IsSystemLocked => _isSystemLocked;

    #region Event Args Classes
    
    public class UpdateAvailableEventArgs : EventArgs
    {
        public string CurrentVersion { get; set; } = "";
        public string LatestVersion { get; set; } = "";
        public string ReleaseNotes { get; set; } = "";
        public DateTime ReleaseDate { get; set; }
        public long FileSizeBytes { get; set; }
        public bool IsMandatory { get; set; }
    }

    public class UpdateProgressEventArgs : EventArgs
    {
        public int ProgressPercent { get; set; }
        public string Status { get; set; } = "";
        public long BytesDownloaded { get; set; }
        public long TotalBytes { get; set; }
    }
    
    #endregion

    private UpdateService()
    {
        _httpClient = new HttpClient();
        _httpClient.Timeout = TimeSpan.FromMinutes(10);
        _updateFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TangentIDMonitor", "Updates");
        
        Directory.CreateDirectory(_updateFolder);
        
        // Setup system lock detection
        SetupLockDetection();
    }

    #region System Lock Detection

    /// <summary>
    /// Setup detection for Windows lock/unlock events
    /// </summary>
    private void SetupLockDetection()
    {
        try
        {
            // Subscribe to session switch events (lock/unlock)
            SystemEvents.SessionSwitch += OnSessionSwitch;
            
            // Also detect power events (sleep/wake)
            SystemEvents.PowerModeChanged += OnPowerModeChanged;
            
            Debug.WriteLine("Lock detection initialized");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Failed to setup lock detection: {ex.Message}");
        }
    }

    private void OnSessionSwitch(object sender, SessionSwitchEventArgs e)
    {
        switch (e.Reason)
        {
            case SessionSwitchReason.SessionLock:
            case SessionSwitchReason.ConsoleDisconnect:
            case SessionSwitchReason.RemoteDisconnect:
                _isSystemLocked = true;
                Debug.WriteLine("System LOCKED - pausing monitoring");
                SystemLockStateChanged?.Invoke(this, true);
                break;

            case SessionSwitchReason.SessionUnlock:
            case SessionSwitchReason.ConsoleConnect:
            case SessionSwitchReason.RemoteConnect:
                _isSystemLocked = false;
                Debug.WriteLine("System UNLOCKED - resuming monitoring");
                SystemLockStateChanged?.Invoke(this, false);
                break;
        }
    }

    private void OnPowerModeChanged(object sender, PowerModeChangedEventArgs e)
    {
        switch (e.Mode)
        {
            case PowerModes.Suspend:
                _isSystemLocked = true;
                Debug.WriteLine("System SUSPEND - pausing monitoring");
                SystemLockStateChanged?.Invoke(this, true);
                break;

            case PowerModes.Resume:
                _isSystemLocked = false;
                Debug.WriteLine("System RESUME - resuming monitoring");
                SystemLockStateChanged?.Invoke(this, false);
                break;
        }
    }

    #endregion

    #region Update Check

    /// <summary>
    /// Start periodic update checks
    /// </summary>
    public void StartPeriodicUpdateCheck(string currentVersion, TimeSpan interval)
    {
        _updateCheckTimer?.Stop();
        _updateCheckTimer?.Dispose();

        _updateCheckTimer = new System.Timers.Timer(interval.TotalMilliseconds);
        _updateCheckTimer.Elapsed += async (s, e) =>
        {
            if (!_isSystemLocked && !_updateInProgress)
            {
                await CheckForUpdateAsync(currentVersion);
            }
        };
        _updateCheckTimer.AutoReset = true;
        _updateCheckTimer.Start();

        // Also check immediately on startup (after 30 seconds delay)
        Task.Run(async () =>
        {
            await Task.Delay(30000);
            if (!_isSystemLocked)
            {
                await CheckForUpdateAsync(currentVersion);
            }
        });

        Debug.WriteLine($"Update check started - interval: {interval.TotalHours} hours");
    }

    /// <summary>
    /// Check if an update is available
    /// </summary>
    public async Task<UpdateAvailableEventArgs?> CheckForUpdateAsync(string currentVersion)
    {
        try
        {
            Debug.WriteLine($"Checking for updates... Current version: {currentVersion}");
            
            var response = await _httpClient.GetAsync(
                $"{_serverUrl}/api/update/check?version={currentVersion}");
            
            if (!response.IsSuccessStatusCode)
            {
                Debug.WriteLine($"Update check failed: {response.StatusCode}");
                return null;
            }

            var json = await response.Content.ReadAsStringAsync();
            var updateInfo = JsonSerializer.Deserialize<UpdateCheckResponse>(json, 
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (updateInfo == null || !updateInfo.UpdateAvailable)
            {
                Debug.WriteLine("No update available");
                return null;
            }

            Debug.WriteLine($"Update available: {updateInfo.LatestVersion}");

            var args = new UpdateAvailableEventArgs
            {
                CurrentVersion = currentVersion,
                LatestVersion = updateInfo.LatestVersion ?? "",
                ReleaseNotes = updateInfo.ReleaseNotes ?? "",
                ReleaseDate = updateInfo.ReleaseDate,
                FileSizeBytes = updateInfo.FileSizeBytes,
                IsMandatory = updateInfo.IsMandatory
            };

            // Check if we should show reminder
            if (ShouldShowUpdateNotification())
            {
                UpdateAvailable?.Invoke(this, args);
            }
            
            return args;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Update check error: {ex.Message}");
            return null;
        }
    }

    private bool ShouldShowUpdateNotification()
    {
        try
        {
            var configPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TangentIDMonitor", "update_reminder.txt");
            
            if (!File.Exists(configPath))
                return true;
            
            var remindTime = DateTime.Parse(File.ReadAllText(configPath));
            return DateTime.Now >= remindTime;
        }
        catch
        {
            return true;
        }
    }

    public void SetRemindLater(int hours = 24)
    {
        try
        {
            var configPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TangentIDMonitor", "update_reminder.txt");
            
            Directory.CreateDirectory(Path.GetDirectoryName(configPath)!);
            File.WriteAllText(configPath, DateTime.Now.AddHours(hours).ToString("o"));
        }
        catch { }
    }

    #endregion

    #region Download and Install

    /// <summary>
    /// Download and install the update (uninstalls old version first)
    /// </summary>
    public async Task<bool> DownloadAndInstallUpdateAsync(string latestVersion, bool silentInstall = true)
    {
        if (_updateInProgress)
            return false;

        _updateInProgress = true;

        try
        {
            // Step 1: Download the installer
            ReportProgress(0, "Downloading update...");
            
            var installerPath = Path.Combine(_updateFolder, $"TangentIDMonitor_Setup_v{latestVersion}.exe");
            
            if (File.Exists(installerPath))
                File.Delete(installerPath);

            var downloadUrl = $"{_serverUrl}/api/update/download?version={latestVersion}";
            
            using (var response = await _httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead))
            {
                response.EnsureSuccessStatusCode();
                
                var totalBytes = response.Content.Headers.ContentLength ?? 0;
                var buffer = new byte[8192];
                var bytesRead = 0L;

                using var contentStream = await response.Content.ReadAsStreamAsync();
                using var fileStream = new FileStream(installerPath, FileMode.Create, FileAccess.Write, FileShare.None);
                
                int read;
                while ((read = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                {
                    await fileStream.WriteAsync(buffer, 0, read);
                    bytesRead += read;

                    if (totalBytes > 0)
                    {
                        var percent = (int)((bytesRead * 100) / totalBytes);
                        ReportProgress(percent, 
                            $"Downloading... {bytesRead / 1024 / 1024:F1} MB / {totalBytes / 1024 / 1024:F1} MB", 
                            bytesRead, totalBytes);
                    }
                }
            }

            ReportProgress(100, "Download complete. Preparing installation...");

            // Verify download
            if (!File.Exists(installerPath) || new FileInfo(installerPath).Length < 100000)
            {
                UpdateError?.Invoke(this, "Download failed or file corrupted");
                return false;
            }

            await Task.Delay(500);
            ReportProgress(100, "Starting installation...");

            // Step 2: Launch installer (will uninstall old version automatically)
            LaunchInstallerWithUninstall(installerPath, silentInstall);
            
            return true;
        }
        catch (Exception ex)
        {
            UpdateError?.Invoke(this, $"Update failed: {ex.Message}");
            return false;
        }
        finally
        {
            _updateInProgress = false;
        }
    }

    /// <summary>
    /// Launch installer that uninstalls old version first
    /// </summary>
    private void LaunchInstallerWithUninstall(string installerPath, bool silentInstall)
    {
        try
        {
            // Create batch script that:
            // 1. Waits for this app to close
            // 2. Uninstalls old version silently
            // 3. Installs new version
            // 4. Starts the new app minimized
            // 5. Cleans up

            var batchPath = Path.Combine(_updateFolder, "update_script.bat");
            var uninstallPath = GetUninstallPath();

            var batchContent = $@"@echo off
setlocal

echo ========================================
echo Tangent ID Monitor Auto-Update
echo ========================================
echo.

echo [1/4] Waiting for application to close...
:waitloop
tasklist /FI ""IMAGENAME eq AutodeskIDMonitor.exe"" 2>NUL | find /I /N ""AutodeskIDMonitor.exe"">NUL
if ""%ERRORLEVEL%""==""0"" (
    timeout /t 2 /nobreak >NUL
    goto waitloop
)

echo [2/4] Removing old version...
{(string.IsNullOrEmpty(uninstallPath) ? "echo No previous installation found" : $"\"{uninstallPath}\" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART")}
timeout /t 3 /nobreak >NUL

echo [3/4] Installing new version...
""{installerPath}"" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS /TASKS=""desktopicon,startupicon""
timeout /t 2 /nobreak >NUL

echo [4/4] Starting ID Monitor...
start """" ""%LOCALAPPDATA%\TangentIDMonitor\AutodeskIDMonitor.exe"" --minimized

echo.
echo Update complete!
timeout /t 3 /nobreak >NUL

:: Cleanup
del ""{batchPath}""
";
            File.WriteAllText(batchPath, batchContent);

            // Start the batch file
            var startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c \"{batchPath}\"",
                UseShellExecute = true,
                CreateNoWindow = !System.Diagnostics.Debugger.IsAttached,
                WindowStyle = System.Diagnostics.Debugger.IsAttached ? ProcessWindowStyle.Normal : ProcessWindowStyle.Hidden
            };

            Process.Start(startInfo);

            // Close this application
            Application.Current.Dispatcher.Invoke(() =>
            {
                Application.Current.Shutdown();
            });
        }
        catch (Exception ex)
        {
            UpdateError?.Invoke(this, $"Failed to launch installer: {ex.Message}");
        }
    }

    /// <summary>
    /// Find the uninstaller path from registry
    /// </summary>
    private string GetUninstallPath()
    {
        try
        {
            // Check HKEY_CURRENT_USER first (per-user install)
            var hkcuPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (var key = Registry.CurrentUser.OpenSubKey(hkcuPath))
            {
                if (key != null)
                {
                    foreach (var subKeyName in key.GetSubKeyNames())
                    {
                        using var subKey = key.OpenSubKey(subKeyName);
                        var displayName = subKey?.GetValue("DisplayName")?.ToString();
                        
                        if (displayName != null && 
                            (displayName.Contains("Tangent ID Monitor") || 
                             displayName.Contains("Autodesk ID Monitor")))
                        {
                            var uninstallString = subKey?.GetValue("UninstallString")?.ToString();
                            if (!string.IsNullOrEmpty(uninstallString))
                            {
                                return uninstallString.Trim('"');
                            }
                        }
                    }
                }
            }

            // Fallback: Check HKEY_LOCAL_MACHINE (for old admin installs)
            var paths = new[]
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            foreach (var path in paths)
            {
                using var key = Registry.LocalMachine.OpenSubKey(path);
                if (key == null) continue;

                foreach (var subKeyName in key.GetSubKeyNames())
                {
                    using var subKey = key.OpenSubKey(subKeyName);
                    var displayName = subKey?.GetValue("DisplayName")?.ToString();
                    
                    if (displayName != null && 
                        (displayName.Contains("Tangent ID Monitor") || 
                         displayName.Contains("Autodesk ID Monitor")))
                    {
                        var uninstallString = subKey?.GetValue("UninstallString")?.ToString();
                        if (!string.IsNullOrEmpty(uninstallString))
                        {
                            // Clean up the path (remove quotes if present)
                            return uninstallString.Trim('"');
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error finding uninstaller: {ex.Message}");
        }

        return "";
    }

    private void ReportProgress(int percent, string status, long bytesDownloaded = 0, long totalBytes = 0)
    {
        UpdateProgress?.Invoke(this, new UpdateProgressEventArgs
        {
            ProgressPercent = percent,
            Status = status,
            BytesDownloaded = bytesDownloaded,
            TotalBytes = totalBytes
        });
    }

    #endregion

    #region Cleanup

    public void CleanupOldUpdates()
    {
        try
        {
            if (!Directory.Exists(_updateFolder)) return;

            foreach (var file in Directory.GetFiles(_updateFolder, "*.exe"))
            {
                try
                {
                    if (new FileInfo(file).CreationTime < DateTime.Now.AddDays(-7))
                        File.Delete(file);
                }
                catch { }
            }

            foreach (var file in Directory.GetFiles(_updateFolder, "*.bat"))
            {
                try { File.Delete(file); } catch { }
            }
        }
        catch { }
    }

    public void Dispose()
    {
        _updateCheckTimer?.Stop();
        _updateCheckTimer?.Dispose();
        
        try
        {
            SystemEvents.SessionSwitch -= OnSessionSwitch;
            SystemEvents.PowerModeChanged -= OnPowerModeChanged;
        }
        catch { }

        _httpClient.Dispose();
    }

    #endregion

    #region Response Classes

    private class UpdateCheckResponse
    {
        public bool UpdateAvailable { get; set; }
        public string? LatestVersion { get; set; }
        public string? ReleaseNotes { get; set; }
        public DateTime ReleaseDate { get; set; }
        public long FileSizeBytes { get; set; }
        public bool IsMandatory { get; set; }
    }

    #endregion
}
