using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using CommunityToolkit.Mvvm.Input;
using AutodeskIDMonitor.Models;
using AutodeskIDMonitor.Services;
using AutodeskIDMonitor.Views;

namespace AutodeskIDMonitor.ViewModels;

public partial class MainViewModel
{
    [RelayCommand]
    private void ToggleAdmin()
    {
        if (AdminService.Instance.IsAdminLoggedIn)
        {
            var result = MessageBox.Show("Logout from admin mode?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes) AdminService.Instance.Logout();
        }
        else
        {
            var loginWindow = new AdminLoginWindow { Owner = Application.Current.MainWindow };
            loginWindow.ShowDialog();
        }
    }

    [RelayCommand]
    private void ChangePassword()
    {
        var changePasswordWindow = new ChangePasswordWindow { Owner = Application.Current.MainWindow };
        changePasswordWindow.ShowDialog();
    }

    [RelayCommand]
    private void SaveSettings()
    {
        var config = ConfigService.Instance.Config;
        config.Country = Country;
        config.Office = Office;
        config.CloudLoggingEnabled = CloudLoggingEnabled;
        config.CloudApiUrl = CloudApiUrl;
        config.CloudApiKey = CloudApiKey;
        config.NetworkLoggingEnabled = NetworkLoggingEnabled;
        config.NetworkLogPath = NetworkLogPath;
        config.MonitoringIntervalSeconds = CheckIntervalSeconds;
        config.StartMinimized = StartMinimized;
        config.AutoStartEnabled = AutoStartEnabled;
        ConfigService.Instance.Save();
        
        _cloudService.UpdateApiKey();
        CloudServerInfo = $"Cloud Server: {CloudApiUrl}";
        _refreshTimer.Interval = TimeSpan.FromSeconds(CheckIntervalSeconds);
        _syncTimer.Interval = TimeSpan.FromSeconds(CheckIntervalSeconds);
        
        MessageBox.Show("Settings saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        AddDebug("Settings saved");
    }

    [RelayCommand]
    private void SaveProfile()
    {
        if (!AdminService.Instance.IsAdminLoggedIn)
        {
            MessageBox.Show("Admin access required.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (string.IsNullOrWhiteSpace(EditEmail))
        {
            MessageBox.Show("Email is required.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var existing = UserProfiles.FirstOrDefault(p => p.Email.Equals(EditEmail, StringComparison.OrdinalIgnoreCase));
        if (existing != null) { existing.Name = EditName; existing.Password = EditPassword; }
        else UserProfiles.Add(new UserProfile { Name = EditName, Email = EditEmail, Password = EditPassword });

        BuildUserEmailMapping();
        MessageBox.Show("Profile saved!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    [RelayCommand]
    private void NewProfile() { SelectedProfile = null; EditName = ""; EditEmail = ""; EditPassword = ""; }

    [RelayCommand]
    private void DeleteProfile()
    {
        if (!AdminService.Instance.IsAdminLoggedIn || SelectedProfile == null) return;
        var result = MessageBox.Show($"Delete {SelectedProfile.Name}?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result == MessageBoxResult.Yes)
        {
            UserProfiles.Remove(SelectedProfile);
            SelectedProfile = null; EditName = ""; EditEmail = ""; EditPassword = "";
            BuildUserEmailMapping();
        }
    }

    [RelayCommand]
    private void ClearHistory()
    {
        if (!AdminService.Instance.IsAdminLoggedIn)
        {
            MessageBox.Show("Admin access required.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        var result = MessageBox.Show("Clear all history?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (result == MessageBoxResult.Yes) { ActivityHistory.Clear(); AddDebug("History cleared"); }
    }

    [RelayCommand]
    private void ClearDebug() => DebugLog = "";

    [RelayCommand]
    private void OpenWebDashboard()
    {
        try
        {
            var url = string.IsNullOrEmpty(CloudApiUrl) ? "http://141.145.153.32:5000/" : CloudApiUrl.TrimEnd('/') + "/";
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to open dashboard: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void OpenDownloadPage()
    {
        try
        {
            var url = string.IsNullOrEmpty(CloudApiUrl) ? "http://141.145.153.32:5000/download" : CloudApiUrl.TrimEnd('/') + "/download";
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to open download page: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void ViewProjectTime()
    {
        try
        {
            // Get project times from the monitor service
            var projectTimes = _monitorService.GetProjectTimes();
            
            if (projectTimes == null || projectTimes.Count == 0)
            {
                MessageBox.Show("No project time data available for today.", "Project Time", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            
            // Build summary
            var summary = new System.Text.StringBuilder();
            summary.AppendLine("📊 Today's Project Time Summary");
            summary.AppendLine("═══════════════════════════════════");
            summary.AppendLine();
            
            double totalHours = 0;
            foreach (var project in projectTimes.OrderByDescending(p => p.Value.CurrentDuration.TotalMinutes))
            {
                var duration = project.Value.CurrentDuration;
                var hours = duration.TotalHours;
                totalHours += hours;
                summary.AppendLine($"📁 {project.Value.ProjectName}");
                summary.AppendLine($"   Duration: {hours:F1} hours ({duration.TotalMinutes:F0} minutes)");
                summary.AppendLine($"   Started: {project.Value.StartTime:HH:mm}");
                summary.AppendLine();
            }
            
            summary.AppendLine("═══════════════════════════════════");
            summary.AppendLine($"📊 Total Project Time: {totalHours:F1} hours");
            
            MessageBox.Show(summary.ToString(), "Project Time Summary", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error getting project time: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private async Task RefreshAsync() => await RefreshAllAsync();

    [RelayCommand]
    private async Task RefreshProjectsAsync() => await RefreshAllAsync();

    [RelayCommand]
    private async Task RefreshProfilesAsync() => await RefreshAllAsync();

    [RelayCommand]
    private async Task RefreshHistoryAsync() => await RefreshAllAsync();

    [RelayCommand]
    private void SaveSessionName()
    {
        if (!AdminService.Instance.IsAdminLoggedIn)
        {
            MessageBox.Show("Admin access required.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (SelectedSession == null || string.IsNullOrWhiteSpace(EditSessionName))
        {
            MessageBox.Show("Please select a user and enter a name.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        // Store original name if not already stored
        if (string.IsNullOrEmpty(SelectedSession.OriginalPersonName))
        {
            SelectedSession.OriginalPersonName = SelectedSession.PersonName;
        }
        
        // Create key for storing name override
        var key = $"{SelectedSession.MachineId}|{SelectedSession.WindowsUser}".ToLower();
        _nameOverrides[key] = EditSessionName;
        
        SelectedSession.PersonName = EditSessionName;
        AddDebug($"Changed name from '{SelectedSession.OriginalPersonName}' to '{EditSessionName}' (persisted)");
        MessageBox.Show("Name updated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    [RelayCommand]
    private void ResetSessionName()
    {
        if (!AdminService.Instance.IsAdminLoggedIn)
        {
            MessageBox.Show("Admin access required.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (SelectedSession == null)
        {
            MessageBox.Show("Please select a user.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        // Create key and remove from overrides
        var key = $"{SelectedSession.MachineId}|{SelectedSession.WindowsUser}".ToLower();
        _nameOverrides.Remove(key);
        
        // Get original name from stored dictionary
        if (_originalNames.TryGetValue(key, out var originalName))
        {
            SelectedSession.PersonName = originalName;
            SelectedSession.OriginalPersonName = originalName;
            EditSessionName = originalName;
            AddDebug($"Reset name to '{originalName}' (removed override)");
            MessageBox.Show("Name reset to original value!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        else if (!string.IsNullOrEmpty(SelectedSession.OriginalPersonName))
        {
            SelectedSession.PersonName = SelectedSession.OriginalPersonName;
            EditSessionName = SelectedSession.PersonName;
            AddDebug($"Reset name to '{SelectedSession.OriginalPersonName}' (removed override)");
            MessageBox.Show("Name reset to original value!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        else
        {
            MessageBox.Show("No original name stored to reset to.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    [RelayCommand]
    private void ExportHistory()
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "CSV Files|*.csv|All Files|*.*",
            FileName = $"IDMonitor_ActivityHistory_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
            Title = "Export Activity History"
        };

        if (dialog.ShowDialog() != true) return;

        try
        {
            var lines = new List<string> { "Timestamp,Event Type,Person Name,Autodesk Email,Previous Email,Machine ID" };
            foreach (var entry in ActivityHistory)
            {
                lines.Add($"\"{entry.Timestamp:yyyy-MM-dd HH:mm:ss}\",\"{entry.EventType}\",\"{entry.PersonName}\",\"{entry.AutodeskEmail}\",\"{entry.PreviousEmail}\",\"{entry.MachineId}\"");
            }
            File.WriteAllLines(dialog.FileName, lines);
            
            var result = MessageBox.Show($"Export successful!\n\nFile saved to:\n{dialog.FileName}\n\nWould you like to open the file?", 
                "Export Complete", MessageBoxButton.YesNo, MessageBoxImage.Information);
            
            if (result == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });

            AddDebug($"Exported activity history to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private async Task TestConnectionAsync()
    {
        try
        {
            var sessions = await _cloudService.GetSessionsAsync();
            MessageBox.Show($"Connection successful! Found {sessions.Count} active sessions.", "Connection Test", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Connection failed: {ex.Message}", "Connection Test", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void OpenDashboard()
    {
        try { if (!string.IsNullOrEmpty(CloudApiUrl)) Process.Start(new ProcessStartInfo(CloudApiUrl) { UseShellExecute = true }); }
        catch { }
    }

    [RelayCommand]
    private void BrowseLogPath()
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Log Files|*.log|All Files|*.*",
            FileName = "network.log",
            Title = "Select Log File Location"
        };
        if (dialog.ShowDialog() == true) NetworkLogPath = dialog.FileName;
    }

    [RelayCommand]
    private void ExportToExcel()
    {
        // Show export selection dialog
        var exportWindow = new Window
        {
            Title = "Export to Excel",
            Width = 350,
            Height = 250,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = Application.Current.MainWindow,
            ResizeMode = ResizeMode.NoResize
        };

        var stackPanel = new System.Windows.Controls.StackPanel { Margin = new Thickness(20) };
        
        var titleText = new System.Windows.Controls.TextBlock 
        { 
            Text = "Select data to export:", 
            FontSize = 14, 
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(0, 0, 0, 15)
        };
        stackPanel.Children.Add(titleText);

        var rbUserProfiles = new System.Windows.Controls.RadioButton { Content = "User Profiles", IsChecked = true, Margin = new Thickness(0, 5, 0, 5) };
        var rbCurrentStatus = new System.Windows.Controls.RadioButton { Content = "Current Status", Margin = new Thickness(0, 5, 0, 5) };
        var rbProjects = new System.Windows.Controls.RadioButton { Content = "Projects", Margin = new Thickness(0, 5, 0, 5) };
        
        stackPanel.Children.Add(rbUserProfiles);
        stackPanel.Children.Add(rbCurrentStatus);
        stackPanel.Children.Add(rbProjects);

        var buttonPanel = new System.Windows.Controls.StackPanel 
        { 
            Orientation = System.Windows.Controls.Orientation.Horizontal, 
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 20, 0, 0)
        };

        var exportButton = new System.Windows.Controls.Button 
        { 
            Content = "Export", 
            Width = 80, 
            Height = 30,
            Margin = new Thickness(0, 0, 10, 0),
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 180, 216)),
            Foreground = System.Windows.Media.Brushes.White
        };
        
        var cancelButton = new System.Windows.Controls.Button 
        { 
            Content = "Cancel", 
            Width = 80, 
            Height = 30
        };

        exportButton.Click += (s, e) =>
        {
            string exportType = "";
            if (rbUserProfiles.IsChecked == true) exportType = "UserProfiles";
            else if (rbCurrentStatus.IsChecked == true) exportType = "CurrentStatus";
            else if (rbProjects.IsChecked == true) exportType = "Projects";

            exportWindow.Close();
            DoExport(exportType);
        };

        cancelButton.Click += (s, e) => exportWindow.Close();

        buttonPanel.Children.Add(exportButton);
        buttonPanel.Children.Add(cancelButton);
        stackPanel.Children.Add(buttonPanel);

        exportWindow.Content = stackPanel;
        exportWindow.ShowDialog();
    }

    private void DoExport(string exportType)
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "CSV Files|*.csv|All Files|*.*",
            FileName = $"IDMonitor_{exportType}_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
            Title = "Export to CSV"
        };

        if (dialog.ShowDialog() != true) return;

        try
        {
            var lines = new List<string>();

            switch (exportType)
            {
                case "UserProfiles":
                    lines.Add("Name,Autodesk Email,Status");
                    foreach (var profile in UserProfiles)
                    {
                        lines.Add($"\"{profile.Name}\",\"{profile.Email}\",\"{profile.StatusText}\"");
                    }
                    break;

                case "CurrentStatus":
                    lines.Add("Person Name,Autodesk Email,Windows User,Machine ID,Status,ID Status,Last Activity");
                    foreach (var session in CurrentSessions)
                    {
                        lines.Add($"\"{session.PersonName}\",\"{session.AutodeskEmail}\",\"{session.WindowsUser}\",\"{session.MachineId}\",\"{(session.IsLoggedIn ? "Logged In" : "Logged Out")}\",\"{session.IdStatus}\",\"{session.LastActivity:yyyy-MM-dd HH:mm:ss}\"");
                    }
                    break;

                case "Projects":
                    lines.Add("User Name,Projects (Revit Version),Sessions,Co-Workers,Machine,Last Activity");
                    foreach (var project in Projects)
                    {
                        var projectsClean = project.ProjectsWithVersions.Replace("\n", " | ");
                        var coWorkersClean = project.CoWorkers.Replace("\n", " | ");
                        lines.Add($"\"{project.UserName}\",\"{projectsClean}\",\"{project.SessionCount}\",\"{coWorkersClean}\",\"{project.MachineName}\",\"{project.LastActivity:yyyy-MM-dd HH:mm:ss}\"");
                    }
                    break;
            }

            System.IO.File.WriteAllLines(dialog.FileName, lines);
            
            var result = MessageBox.Show($"Export successful!\n\nFile saved to:\n{dialog.FileName}\n\nWould you like to open the file?", 
                "Export Complete", MessageBoxButton.YesNo, MessageBoxImage.Information);
            
            if (result == MessageBoxResult.Yes)
            {
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
            }

            AddDebug($"Exported {exportType} to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            AddDebug($"Export error: {ex.Message}");
        }
    }

    private async Task RefreshAllAsync()
    {
        try
        {
            var sessions = await _cloudService.GetSessionsAsync();
            var cloudCurrentUser = sessions.FirstOrDefault(s => 
                s.MachineName.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase) &&
                s.WindowsUser.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase));

            if (!_isAutodeskLoggedIn && cloudCurrentUser != null && cloudCurrentUser.IsLoggedIn)
            {
                _isAutodeskLoggedIn = true;
                _currentAutodeskEmail = cloudCurrentUser.AutodeskEmail ?? "";
                UpdateLoginStatusDisplay();
                AddDebug($"Using cloud data: {_currentAutodeskEmail}");
            }

            Application.Current.Dispatcher.Invoke(() =>
            {
                CurrentSessions.Clear();
                Projects.Clear();
                var allSessions = new List<UserSession>();

                // Get LOCAL Revit data first (most accurate for current user)
                var localRevitProcesses = _monitorService.GetActiveRevitProcesses();
                var localProjectTimes = _monitorService.GetProjectTimes();
                
                // Build local open projects list
                var localOpenProjects = localRevitProcesses
                    .Where(p => !string.IsNullOrEmpty(p.ProjectName))
                    .Select(p => new RevitProjectInfo 
                    { 
                        ProjectName = p.ProjectName, 
                        RevitVersion = p.RevitVersion 
                    })
                    .ToList();
                
                // Get current project from local detection
                var localCurrentProject = localRevitProcesses
                    .FirstOrDefault(p => !string.IsNullOrEmpty(p.ProjectName))?.ProjectName ?? "";
                var localRevitVersion = localRevitProcesses.FirstOrDefault()?.RevitVersion ?? "";
                var localSessionCount = localRevitProcesses.Count;
                
                // Fallback to cloud data if local detection found nothing
                if (localOpenProjects.Count == 0 && cloudCurrentUser != null)
                {
                    localOpenProjects = cloudCurrentUser.OpenProjects ?? new List<RevitProjectInfo>();
                    localCurrentProject = cloudCurrentUser.CurrentProject ?? "";
                    localRevitVersion = cloudCurrentUser.RevitVersion ?? "";
                    localSessionCount = cloudCurrentUser.RevitSessionCount;
                }

                // Determine best email for current user
                string currentUserEmail = "(Not logged in)";
                if (_isAutodeskLoggedIn && !string.IsNullOrWhiteSpace(_currentAutodeskEmail))
                {
                    // Auto-correct the email format
                    currentUserEmail = LocalStorageExtensions.AutoCorrectAutodeskEmail(_currentAutodeskEmail, _currentWindowsDisplayName);
                }
                else
                {
                    // Fallback: Try to find email from UserProfiles by matching PersonName
                    var matchingProfile = UserProfiles.FirstOrDefault(p => 
                        !string.IsNullOrEmpty(p.Name) &&
                        p.Name.Equals(_currentWindowsDisplayName, StringComparison.OrdinalIgnoreCase));
                    if (matchingProfile != null && !string.IsNullOrWhiteSpace(matchingProfile.Email))
                    {
                        currentUserEmail = matchingProfile.Email;
                    }
                    else
                    {
                        // Try to generate email from display name
                        var generatedEmail = LocalStorageExtensions.AutoCorrectAutodeskEmail("", _currentWindowsDisplayName);
                        if (!string.IsNullOrEmpty(generatedEmail) && generatedEmail.Contains("@"))
                        {
                            currentUserEmail = generatedEmail;
                        }
                    }
                }

                // Add current user session with LOCAL data
                allSessions.Add(new UserSession
                {
                    PersonName = _currentWindowsDisplayName,
                    AutodeskEmail = currentUserEmail,
                    WindowsUser = _currentWindowsUser,
                    MachineId = _currentMachineName,
                    LastActivity = DateTime.Now,
                    IsLoggedIn = _isAutodeskLoggedIn,
                    IsCurrentUser = true,
                    RevitVersion = localRevitVersion,
                    CurrentProject = localCurrentProject,
                    RevitSessionCount = localSessionCount,
                    OpenProjects = localOpenProjects
                });

                // Add cloud sessions (with proper deduplication)
                var addedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                
                // Mark current user as added
                addedKeys.Add($"{_currentMachineName}|{_currentWindowsUser}");
                
                foreach (var session in sessions)
                {
                    // Create unique key for this session
                    var sessionKey = $"{session.MachineName}|{session.WindowsUser}";
                    
                    // Skip if already added (prevents duplicates)
                    if (addedKeys.Contains(sessionKey))
                        continue;
                    
                    addedKeys.Add(sessionKey);

                    // Determine best email - use cloud data or fallback to UserProfiles
                    string sessionEmail = "(Not logged in)";
                    var displayName = session.GetDisplayName();
                    
                    if (!string.IsNullOrWhiteSpace(session.AutodeskEmail))
                    {
                        // Auto-correct the email format
                        sessionEmail = LocalStorageExtensions.AutoCorrectAutodeskEmail(session.AutodeskEmail, displayName);
                    }
                    else if (session.IsLoggedIn)
                    {
                        // Fallback: Try to find email from UserProfiles by matching DisplayName
                        var matchingProfile = UserProfiles.FirstOrDefault(p => 
                            !string.IsNullOrEmpty(p.Name) &&
                            p.Name.Equals(displayName, StringComparison.OrdinalIgnoreCase));
                        if (matchingProfile != null && !string.IsNullOrWhiteSpace(matchingProfile.Email))
                        {
                            sessionEmail = matchingProfile.Email;
                        }
                        else
                        {
                            // Try to generate email from display name
                            var generatedEmail = LocalStorageExtensions.AutoCorrectAutodeskEmail("", displayName);
                            if (!string.IsNullOrEmpty(generatedEmail) && generatedEmail.Contains("@"))
                            {
                                sessionEmail = generatedEmail;
                            }
                        }
                    }

                    allSessions.Add(new UserSession
                    {
                        PersonName = session.GetDisplayName(),
                        AutodeskEmail = sessionEmail,
                        WindowsUser = session.WindowsUser,
                        MachineId = session.MachineName,
                        // Convert from server UTC time to UAE local time (UTC+4)
                        LastActivity = session.LastSeen.Kind == DateTimeKind.Utc 
                            ? session.LastSeen.ToLocalTime() 
                            : session.LastSeen.AddHours(4),
                        IsLoggedIn = session.IsLoggedIn,
                        IsCurrentUser = false,
                        ClientVersion = session.ClientVersion,
                        RevitVersion = session.RevitVersion ?? "",
                        CurrentProject = session.CurrentProject ?? "",
                        RevitSessionCount = session.RevitSessionCount,
                        OpenProjects = session.OpenProjects ?? new List<RevitProjectInfo>()
                    });
                }

                // Apply name overrides before building Projects and projectUserMap
                // Also store original names for reset functionality
                foreach (var session in allSessions)
                {
                    var key = $"{session.MachineId}|{session.WindowsUser}".ToLower();
                    
                    // Store original name from cloud (before any override)
                    if (!_originalNames.ContainsKey(key))
                    {
                        _originalNames[key] = session.PersonName;
                    }
                    
                    // Apply override if exists
                    if (_nameOverrides.TryGetValue(key, out var overrideName))
                    {
                        session.PersonName = overrideName;
                    }
                }

                // Build project-to-users mapping after name overrides (for co-workers calculation)
                var projectUserMap = new Dictionary<string, List<string>>();
                foreach (var s in allSessions.Where(s => s.IsLoggedIn))
                {
                    var projects = s.OpenProjects.Count > 0 
                        ? s.OpenProjects.Select(p => p.ProjectName).ToList()
                        : (!string.IsNullOrEmpty(s.CurrentProject) ? new List<string> { s.CurrentProject } : new List<string>());
                    
                    foreach (var proj in projects.Where(p => !string.IsNullOrEmpty(p)))
                    {
                        if (!projectUserMap.ContainsKey(proj))
                            projectUserMap[proj] = new List<string>();
                        if (!projectUserMap[proj].Contains(s.PersonName))
                            projectUserMap[proj].Add(s.PersonName);
                    }
                }

                // Build Projects tab - ONE ROW PER USER with all their projects
                var loggedInUsers = allSessions.Where(s => s.IsLoggedIn).OrderBy(s => s.PersonName).ToList();
                foreach (var user in loggedInUsers)
                {
                    // Build combined "Project (Revit Version)" format from OpenProjects
                    var projectsWithVersions = new List<string>();
                    
                    // Determine actual Revit count - use OpenProjects.Count if available, else RevitSessionCount
                    var actualRevitCount = user.OpenProjects.Count > 0 
                        ? user.OpenProjects.Count 
                        : (user.RevitSessionCount > 0 ? user.RevitSessionCount : (!string.IsNullOrEmpty(user.CurrentProject) ? 1 : 0));
                    
                    if (user.OpenProjects.Count > 0)
                    {
                        foreach (var proj in user.OpenProjects)
                        {
                            var version = !string.IsNullOrEmpty(proj.RevitVersion) ? proj.RevitVersion : "Unknown";
                            version = version.Replace("Autodesk ", "");
                            projectsWithVersions.Add($"{proj.ProjectName} ({version})");
                        }
                    }
                    else if (!string.IsNullOrEmpty(user.CurrentProject))
                    {
                        // Fallback to single project
                        var version = !string.IsNullOrEmpty(user.RevitVersion) ? user.RevitVersion : "Unknown";
                        version = version.Replace("Autodesk ", "");
                        projectsWithVersions.Add($"{user.CurrentProject} ({version})");
                    }

                    // Build co-workers string showing which project they share
                    var coWorkersByProject = new List<string>();
                    var userProjects = user.OpenProjects.Count > 0 
                        ? user.OpenProjects.Select(p => p.ProjectName).ToList()
                        : (!string.IsNullOrEmpty(user.CurrentProject) ? new List<string> { user.CurrentProject } : new List<string>());
                    
                    foreach (var proj in userProjects.Where(p => !string.IsNullOrEmpty(p)))
                    {
                        if (projectUserMap.TryGetValue(proj, out var usersOnProject))
                        {
                            var coworkers = usersOnProject
                                .Where(u => !u.Equals(user.PersonName, StringComparison.OrdinalIgnoreCase))
                                .OrderBy(n => n)
                                .ToList();
                            
                            if (coworkers.Count > 0)
                            {
                                coWorkersByProject.Add($"{proj}: {string.Join(", ", coworkers)}");
                            }
                        }
                    }

                    Projects.Add(new ProjectInfo
                    {
                        UserName = user.PersonName,
                        ProjectsWithVersions = projectsWithVersions.Count > 0 ? string.Join("\n", projectsWithVersions) : "-",
                        SessionCount = actualRevitCount,
                        MachineName = user.MachineId,
                        LastActivity = user.LastActivity,
                        CoWorkers = coWorkersByProject.Count > 0 ? string.Join("\n", coWorkersByProject) : "-"
                    });
                }

                // Calculate email usage and ID status for Current Sessions
                var emailUsageMap = CalculateEmailUsage(allSessions);
                
                // Calculate SharedIdInfo and set original names for each session
                // (Name overrides already applied above)
                foreach (var session in allSessions)
                {
                    // Store original name for reset functionality (before override was applied)
                    var key = $"{session.MachineId}|{session.WindowsUser}".ToLower();
                    if (_nameOverrides.ContainsKey(key))
                    {
                        // Original name is stored in cloud, we need to track it
                        // Find original name from cloud session
                        session.OriginalPersonName = _originalNames.GetValueOrDefault(key, session.PersonName);
                    }
                    else
                    {
                        session.OriginalPersonName = session.PersonName;
                        // Store it for future reference
                        _originalNames[key] = session.PersonName;
                    }
                    
                    session.Details = GetEmailUsageDetails(session.AutodeskEmail, emailUsageMap);
                    session.IdStatus = GetIdStatus(session);
                    
                    // Calculate SharedIdInfo - show other users with the same ID (names only, no machine IDs)
                    if (session.IsLoggedIn && !string.IsNullOrEmpty(session.AutodeskEmail) && session.AutodeskEmail != "(Not logged in)")
                    {
                        var sharedUsers = allSessions
                            .Where(s => s.IsLoggedIn && 
                                        s.AutodeskEmail.Equals(session.AutodeskEmail, StringComparison.OrdinalIgnoreCase) &&
                                        !s.PersonName.Equals(session.PersonName, StringComparison.OrdinalIgnoreCase))
                            .Select(s => s.PersonName)
                            .ToList();
                        
                        if (sharedUsers.Count > 0)
                        {
                            session.SharedIdInfo = $"Shared with: {string.Join(", ", sharedUsers)}";
                        }
                        else
                        {
                            session.SharedIdInfo = "-";
                        }
                    }
                    else
                    {
                        session.SharedIdInfo = "-";
                    }
                }
                
                // Sort sessions alphabetically by PersonName for stable order
                var sortedSessions = allSessions.OrderBy(s => s.PersonName, StringComparer.OrdinalIgnoreCase).ToList();
                foreach (var session in sortedSessions)
                {
                    CurrentSessions.Add(session);
                }

                // Update profile statuses - match by Autodesk email from active sessions
                var activeSessionEmails = allSessions
                    .Where(s => s.IsLoggedIn && !string.IsNullOrEmpty(s.AutodeskEmail) && s.AutodeskEmail != "(Not logged in)")
                    .Select(s => s.AutodeskEmail.ToLower().Trim())
                    .ToHashSet();

                // Debug: log active emails
                AddDebug($"Active session emails: {string.Join(", ", activeSessionEmails)}");

                foreach (var profile in UserProfiles)
                {
                    var profileEmail = profile.Email.ToLower().Trim();
                    profile.IsActive = activeSessionEmails.Contains(profileEmail);
                }
                
                // Force UI refresh for profiles
                OnPropertyChanged(nameof(UserProfiles));

                // Calculate total Revit sessions
                var totalRevitSessions = allSessions.Where(s => s.IsLoggedIn).Sum(s => Math.Max(0, s.RevitSessionCount));
                TotalRevitSessions = totalRevitSessions;
                TotalRevitSessionsText = $"Total Revit Sessions: {totalRevitSessions}";
                StatusBarText = _isAutodeskLoggedIn 
                    ? $"Current: {_currentAutodeskEmail}" 
                    : "Not logged in";
                LastRefreshText = $"Last refresh: {DateTime.Now:HH:mm:ss}";
            });

            AddDebug($"Refresh: {CurrentSessions.Count} sessions, {Projects.Count} projects");
        }
        catch (Exception ex) { AddDebug($"Refresh error: {ex.Message}"); }
    }

    private string GetIdStatus(UserSession session)
    {
        if (!session.IsLoggedIn || string.IsNullOrEmpty(session.AutodeskEmail) || session.AutodeskEmail == "(Not logged in)")
            return "-";

        var personNameLower = session.PersonName.ToLower().Trim();
        var currentEmail = session.AutodeskEmail.ToLower();
        string? expectedEmail = null;
        
        if (_userEmailMapping.TryGetValue(personNameLower, out var mapped))
            expectedEmail = mapped;
        else
        {
            foreach (var kvp in _userEmailMapping)
            {
                var nameParts = kvp.Key.Split(' ');
                if (nameParts.Any(p => personNameLower.Contains(p) || p.Contains(personNameLower)))
                { expectedEmail = kvp.Value; break; }
            }
        }

        if (expectedEmail == null)
        {
            var matchingProfile = UserProfiles.FirstOrDefault(p => p.Email.Equals(session.AutodeskEmail, StringComparison.OrdinalIgnoreCase));
            if (matchingProfile != null)
                return matchingProfile.Name.Equals(session.PersonName, StringComparison.OrdinalIgnoreCase) ? "Same ID" : "Different ID";
            return "-";
        }

        return currentEmail.Equals(expectedEmail, StringComparison.OrdinalIgnoreCase) ? "Same ID" : "Different ID";
    }

    private Dictionary<string, (int Count, List<string> Users)> CalculateEmailUsage(List<UserSession> sessions)
    {
        var result = new Dictionary<string, (int Count, List<string> Users)>(StringComparer.OrdinalIgnoreCase);
        var emailGroups = sessions
            .Where(s => s.IsLoggedIn && !string.IsNullOrEmpty(s.AutodeskEmail) && s.AutodeskEmail != "(Not logged in)")
            .GroupBy(s => s.AutodeskEmail.ToLower());

        foreach (var group in emailGroups)
        {
            var users = group.Select(s => s.PersonName).Distinct().ToList();
            result[group.Key] = (users.Count, users);
        }
        return result;
    }

    private string GetEmailUsageDetails(string? email, Dictionary<string, (int Count, List<string> Users)> usageMap)
    {
        if (string.IsNullOrEmpty(email) || email == "(Not logged in)") return "-";
        if (usageMap.TryGetValue(email.ToLower(), out var usage) && usage.Count > 1)
            return $"🔴 Shared by {usage.Count}: {string.Join(", ", usage.Users)}";
        return "-";
    }

    private async Task SyncToCloudAsync()
    {
        if (!CloudLoggingEnabled) return;
        try
        {
            // Get real-time activity data
            var activityState = _monitorService.GetCurrentActivityState().ToString();
            var idleSeconds = _monitorService.GetIdleSeconds();
            var revitProcesses = _monitorService.GetActiveRevitProcesses();
            var projectTimes = _monitorService.GetProjectTimes();
            
            // Get activity breakdown and meeting status for server
            var activityBreakdown = _monitorService.GetDailyActivityBreakdown();
            var meetings = _monitorService.GetActiveMeetings();
            var activeMeeting = meetings.FirstOrDefault(m => m.IsInActiveMeeting);
            var isInMeeting = activeMeeting != null;
            var meetingApp = activeMeeting?.AppName ?? "";
            
            // If MonitorService shows 0 Revit hours, try to get from LocalStorageService
            if (activityBreakdown.RevitMinutes < 1)
            {
                var todayRecord = _localStorage.GetUserDailyRecord(_currentMachineName, _currentWindowsUser, DateTime.Today);
                if (todayRecord != null && todayRecord.TotalHours > 0)
                {
                    activityBreakdown.RevitMinutes = todayRecord.TotalHours * 60;
                }
            }
            
            // Send enhanced activity status with breakdown
            var success = await _cloudService.SendActivityStatusAsync(
                _currentWindowsUser, _currentWindowsDisplayName, _currentMachineName,
                _currentAutodeskEmail, _isAutodeskLoggedIn, Country, Office, APP_VERSION,
                activityState, idleSeconds, revitProcesses, projectTimes,
                activityBreakdown, isInMeeting, meetingApp);

            Application.Current.Dispatcher.Invoke(() =>
            {
                IsConnected = success || _isAutodeskLoggedIn;
                IsServerConnected = success;
                CloudStatusText = success ? $"Cloud: ✓ ({Country})" : "Cloud: ✗";
                LastSyncTime = $"Last sync: {DateTime.Now:HH:mm:ss}";
            });
            
            // Update live dashboard data from cloud
            if (success)
            {
                var sessions = await _cloudService.GetSessionsAsync();
                
                // Update Activity History from cloud session changes
                UpdateActivityHistoryFromCloud(sessions);
                
                // Preserve selected project before updating
                var selectedProjectName = SelectedLiveProject?.ProjectName;
                
                UpdateLiveDashboard(sessions);
                
                // Restore selection if it still exists
                if (!string.IsNullOrEmpty(selectedProjectName))
                {
                    SelectedLiveProject = LiveProjects.FirstOrDefault(p => p.ProjectName == selectedProjectName);
                }
                
                // Track activity for CURRENT USER using LOCAL data (most accurate)
                var localProjects = revitProcesses
                    .Where(p => !string.IsNullOrEmpty(p.ProjectName))
                    .Select(p => new RevitProjectInfo { ProjectName = p.ProjectName, RevitVersion = p.RevitVersion ?? "" })
                    .ToList();
                
                if (localProjects.Count > 0)
                {
                    // Create a session info for current user with local project data
                    var currentUserSession = new CloudSessionInfo
                    {
                        WindowsUser = _currentWindowsUser,
                        WindowsDisplayName = _currentWindowsDisplayName,
                        MachineName = _currentMachineName,
                        IsLoggedIn = _isAutodeskLoggedIn,
                        RevitSessionCount = revitProcesses.Count,
                        RevitVersion = revitProcesses.FirstOrDefault()?.RevitVersion ?? "",
                        OpenProjects = localProjects,
                        CurrentProject = localProjects.FirstOrDefault()?.ProjectName ?? "",
                        LastSeen = DateTime.Now
                    };
                    _timeTracking.TrackActivity(currentUserSession);
                }
                
                // Track activity from OTHER cloud sessions (but skip current user - we already tracked with local data)
                foreach (var session in sessions.Where(s => 
                    !s.MachineName.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase) ||
                    !s.WindowsUser.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase)))
                {
                    _timeTracking.TrackActivity(session);
                }
            }
        }
        catch (Exception ex) 
        { 
            AddDebug($"Sync error: {ex.Message}");
            IsServerConnected = false;
        }
    }
    
    // Track previous cloud session states for activity history
    private Dictionary<string, CloudSessionInfo> _previousCloudSessions = new();
    
    private void UpdateActivityHistoryFromCloud(List<CloudSessionInfo> sessions)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            bool hasNewEntries = false;
            
            foreach (var session in sessions)
            {
                var key = $"{session.MachineName}|{session.WindowsUser}".ToLower();
                
                // Skip current user (tracked locally)
                if (session.MachineName.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase) &&
                    session.WindowsUser.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase))
                    continue;
                
                if (_previousCloudSessions.TryGetValue(key, out var previous))
                {
                    // Check for status changes
                    if (!previous.IsLoggedIn && session.IsLoggedIn)
                    {
                        // User logged in
                        var entry = new HistoryEntry
                        {
                            Timestamp = DateTime.Now,
                            EventType = "Login",
                            PersonName = session.GetDisplayName(),
                            AutodeskEmail = session.AutodeskEmail,
                            MachineId = session.MachineName
                        };
                        ActivityHistory.Insert(0, entry);
                        _localStorage.AddHistoryEntry(entry);
                        hasNewEntries = true;
                    }
                    else if (previous.IsLoggedIn && !session.IsLoggedIn)
                    {
                        // User logged out
                        var entry = new HistoryEntry
                        {
                            Timestamp = DateTime.Now,
                            EventType = "Logout",
                            PersonName = session.GetDisplayName(),
                            PreviousEmail = previous.AutodeskEmail,
                            MachineId = session.MachineName
                        };
                        ActivityHistory.Insert(0, entry);
                        _localStorage.AddHistoryEntry(entry);
                        hasNewEntries = true;
                    }
                    else if (previous.AutodeskEmail != session.AutodeskEmail && 
                             !string.IsNullOrEmpty(previous.AutodeskEmail) && 
                             !string.IsNullOrEmpty(session.AutodeskEmail))
                    {
                        // ID changed
                        var entry = new HistoryEntry
                        {
                            Timestamp = DateTime.Now,
                            EventType = "ID Change",
                            PersonName = session.GetDisplayName(),
                            PreviousEmail = previous.AutodeskEmail,
                            AutodeskEmail = session.AutodeskEmail,
                            MachineId = session.MachineName
                        };
                        ActivityHistory.Insert(0, entry);
                        _localStorage.AddHistoryEntry(entry);
                        hasNewEntries = true;
                    }
                }
                else if (_firstCloudSync)
                {
                    // First time seeing this user - add initial state entry if logged in
                    if (session.IsLoggedIn && !string.IsNullOrEmpty(session.AutodeskEmail))
                    {
                        var entry = new HistoryEntry
                        {
                            Timestamp = DateTime.Now,
                            EventType = "Online",
                            PersonName = session.GetDisplayName(),
                            AutodeskEmail = session.AutodeskEmail,
                            MachineId = session.MachineName
                        };
                        ActivityHistory.Insert(0, entry);
                        hasNewEntries = true;
                    }
                }
                
                // Update previous state
                _previousCloudSessions[key] = session;
            }
            
            // Mark first sync complete
            if (_firstCloudSync)
            {
                _firstCloudSync = false;
                if (hasNewEntries)
                {
                    // Save all current entries
                    _localStorage.SaveActivityHistory(ActivityHistory);
                }
            }
            
            // Limit history to 1000 entries
            while (ActivityHistory.Count > 1000)
            {
                ActivityHistory.RemoveAt(ActivityHistory.Count - 1);
            }
        });
    }
    
    private bool _firstCloudSync = true;

    private void UpdateLiveDashboard(List<CloudSessionInfo> sessions)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            // Get LOCAL Revit data (most accurate for current user)
            var localRevitProcesses = _monitorService.GetActiveRevitProcesses();
            var localOpenProjects = localRevitProcesses
                .Where(p => !string.IsNullOrEmpty(p.ProjectName))
                .Select(p => new RevitProjectInfo 
                { 
                    ProjectName = p.ProjectName, 
                    RevitVersion = p.RevitVersion ?? ""
                })
                .ToList();
            var localCurrentProject = localRevitProcesses
                .FirstOrDefault(p => !string.IsNullOrEmpty(p.ProjectName))?.ProjectName ?? "";
            var localRevitVersion = localRevitProcesses.FirstOrDefault()?.RevitVersion ?? "";
            var localRevitCount = localRevitProcesses.Count;
            
            // Update Live Users
            LiveUsers.Clear();
            foreach (var session in sessions.OrderBy(s => s.GetDisplayName()))
            {
                // Check if this is the current user - merge local data
                bool isCurrentUser = session.MachineName.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase) &&
                                    session.WindowsUser.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase);
                
                // Get project from OpenProjects or CurrentProject
                var currentProject = "";
                var revitVersion = session.RevitVersion?.Replace("Autodesk ", "") ?? "";
                var revitCount = session.RevitSessionCount;
                List<RevitProjectInfo> openProjects;
                
                if (isCurrentUser && localOpenProjects.Count > 0)
                {
                    // Use LOCAL data for current user (most accurate)
                    openProjects = localOpenProjects;
                    currentProject = string.Join(", ", localOpenProjects.Select(p => p.ProjectName));
                    revitVersion = localRevitVersion?.Replace("Autodesk ", "") ?? "";
                    revitCount = localRevitCount > 0 ? localRevitCount : session.RevitSessionCount;
                }
                else if (session.OpenProjects?.Any() == true)
                {
                    openProjects = session.OpenProjects;
                    currentProject = string.Join(", ", session.OpenProjects.Select(p => p.ProjectName));
                    var firstProj = session.OpenProjects.FirstOrDefault();
                    if (firstProj != null && !string.IsNullOrEmpty(firstProj.RevitVersion))
                    {
                        revitVersion = firstProj.RevitVersion.Replace("Autodesk ", "");
                    }
                }
                else if (!string.IsNullOrEmpty(session.CurrentProject))
                {
                    openProjects = new List<RevitProjectInfo> 
                    { 
                        new RevitProjectInfo { ProjectName = session.CurrentProject, RevitVersion = session.RevitVersion ?? "" } 
                    };
                    currentProject = session.CurrentProject;
                }
                else
                {
                    openProjects = new List<RevitProjectInfo>();
                }
                
                var liveUser = new LiveUserActivity
                {
                    UserId = session.WindowsUser,
                    UserName = session.GetDisplayName(),
                    MachineId = session.MachineName,
                    AutodeskEmail = session.AutodeskEmail,
                    IsRevitOpen = revitCount > 0,
                    RevitInstanceCount = revitCount,
                    CurrentProject = currentProject,
                    RevitVersion = revitVersion,
                    // Convert from server UTC time to local time (UAE is UTC+4)
                    LastActivity = session.LastSeen.Kind == DateTimeKind.Utc 
                        ? session.LastSeen.ToLocalTime() 
                        : session.LastSeen.AddHours(4) // Assume server is UTC, add 4 hours for UAE
                };
                
                // Set activity status with clear symbols
                if (!session.IsLoggedIn)
                {
                    liveUser.ActivityStatus = "Offline";
                    liveUser.ActivityStatusIcon = "⬤"; // Gray dot
                }
                else if (revitCount > 0)
                {
                    liveUser.ActivityStatus = "Active";
                    liveUser.ActivityStatusIcon = "✓"; // Green tick
                }
                else
                {
                    liveUser.ActivityStatus = "Idle";
                    liveUser.ActivityStatusIcon = "✗"; // Red cross
                }
                
                // Build active projects list
                liveUser.ActiveProjects = openProjects.Select(p => p.ProjectName).ToList();
                
                LiveUsers.Add(liveUser);
            }
            
            // Update Live Projects - collect from ALL users including current user's local data
            LiveProjects.Clear();
            var projectUserMap = new Dictionary<string, List<(CloudSessionInfo Session, string RevitVersion, List<RevitProjectInfo> Projects)>>(StringComparer.OrdinalIgnoreCase);
            
            foreach (var session in sessions.Where(s => s.IsLoggedIn))
            {
                // Check if this is the current user - use local data
                bool isCurrentUser = session.MachineName.Equals(_currentMachineName, StringComparison.OrdinalIgnoreCase) &&
                                    session.WindowsUser.Equals(_currentWindowsUser, StringComparison.OrdinalIgnoreCase);
                
                List<RevitProjectInfo> projectsToUse;
                string versionToUse;
                
                if (isCurrentUser && localOpenProjects.Count > 0)
                {
                    projectsToUse = localOpenProjects;
                    versionToUse = localRevitVersion?.Replace("Autodesk ", "") ?? "";
                }
                else if (session.OpenProjects?.Any() == true)
                {
                    projectsToUse = session.OpenProjects;
                    versionToUse = session.RevitVersion?.Replace("Autodesk ", "") ?? "";
                }
                else if (!string.IsNullOrEmpty(session.CurrentProject))
                {
                    projectsToUse = new List<RevitProjectInfo> 
                    { 
                        new RevitProjectInfo { ProjectName = session.CurrentProject, RevitVersion = session.RevitVersion ?? "" } 
                    };
                    versionToUse = session.RevitVersion?.Replace("Autodesk ", "") ?? "";
                }
                else
                {
                    continue; // No projects to track
                }
                
                foreach (var proj in projectsToUse.Where(p => !string.IsNullOrEmpty(p.ProjectName)))
                {
                    if (!projectUserMap.ContainsKey(proj.ProjectName))
                        projectUserMap[proj.ProjectName] = new List<(CloudSessionInfo, string, List<RevitProjectInfo>)>();
                    
                    var version = !string.IsNullOrEmpty(proj.RevitVersion) 
                        ? proj.RevitVersion.Replace("Autodesk ", "") 
                        : versionToUse;
                    
                    if (!projectUserMap[proj.ProjectName].Any(x => x.Session.WindowsUser == session.WindowsUser))
                        projectUserMap[proj.ProjectName].Add((session, version, projectsToUse));
                }
            }
            
            // Sort projects ALPHABETICALLY for stable, readable list
            foreach (var kvp in projectUserMap.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
            {
                var revitVersions = kvp.Value
                    .Select(x => x.RevitVersion)
                    .Where(v => !string.IsNullOrEmpty(v))
                    .Distinct()
                    .ToList();
                
                var project = new LiveProjectSummary
                {
                    ProjectName = kvp.Key,
                    ActiveUserCount = kvp.Value.Count,
                    ActiveUsers = string.Join(", ", kvp.Value.Select(x => x.Session.GetDisplayName())),
                    HasActiveUsers = true,
                    // Convert from server UTC time to UAE local time (UTC+4)
                    LastActivity = kvp.Value.Max(x => x.Session.LastSeen).AddHours(4),
                    UserNames = kvp.Value.Select(x => x.Session.GetDisplayName()).ToList(),
                    RevitVersion = string.Join(", ", revitVersions),
                    RevitVersions = revitVersions,
                    TotalHoursWorked = 0
                };
                LiveProjects.Add(project);
            }
            
            // Update Dashboard Stats
            DashboardStats.TotalUsers = sessions.Count;
            DashboardStats.ActiveUsers = LiveUsers.Count(u => u.IsRevitOpen);
            DashboardStats.IdleUsers = sessions.Count(s => s.IsLoggedIn) - DashboardStats.ActiveUsers;
            DashboardStats.OfflineUsers = sessions.Count(s => !s.IsLoggedIn);
            DashboardStats.ActiveProjects = LiveProjects.Count;
            DashboardStats.TotalRevitInstances = LiveUsers.Sum(u => u.RevitInstanceCount);
        });
    }

    // ===== TIME TRACKING COMMANDS =====

    [RelayCommand]
    private void SelectDate(DateTime date)
    {
        SelectedDate = date;
    }

    [RelayCommand]
    private void PreviousMonth()
    {
        CalendarMonth = CalendarMonth.AddMonths(-1);
    }

    [RelayCommand]
    private void NextMonth()
    {
        CalendarMonth = CalendarMonth.AddMonths(1);
    }

    [RelayCommand]
    private void ViewUserDetails(UserSession session)
    {
        if (session != null)
        {
            LoadUserDetails(session.MachineId, session.WindowsUser);
        }
    }

    [RelayCommand]
    private void ViewProjectUserDetails(ProjectInfo project)
    {
        if (project != null)
        {
            // Find the user session for this project entry
            var session = CurrentSessions.FirstOrDefault(s => s.PersonName == project.UserName && s.MachineId == project.MachineName);
            if (session != null)
            {
                LoadUserDetails(session.MachineId, session.WindowsUser);
            }
        }
    }

    [RelayCommand]
    private void ExportDailyReport()
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "CSV Files|*.csv|All Files|*.*",
            FileName = $"DailyReport_{SelectedDate:yyyy-MM-dd}.csv",
            Title = "Export Daily Report"
        };

        if (dialog.ShowDialog() != true) return;

        try
        {
            var records = _localStorage.GetDailyRecords(SelectedDate);
            _excelExport.ExportDailyRecords(records, dialog.FileName);
            
            var result = MessageBox.Show($"Export successful!\n\nFile saved to:\n{dialog.FileName}\n\nWould you like to open the file?", 
                "Export Complete", MessageBoxButton.YesNo, MessageBoxImage.Information);
            
            if (result == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });

            AddDebug($"Exported daily report to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void ExportProjectReport()
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "CSV Files|*.csv|All Files|*.*",
            FileName = $"ProjectReport_{SelectedDate:yyyy-MM-dd}.csv",
            Title = "Export Project Report"
        };

        if (dialog.ShowDialog() != true) return;

        try
        {
            var summaries = _timeTracking.GetProjectStats(SelectedDate);
            _excelExport.ExportProjectSummaries(summaries, SelectedDate, dialog.FileName);
            
            var result = MessageBox.Show($"Export successful!\n\nFile saved to:\n{dialog.FileName}\n\nWould you like to open the file?", 
                "Export Complete", MessageBoxButton.YesNo, MessageBoxImage.Information);
            
            if (result == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });

            AddDebug($"Exported project report to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void ExportUserReport()
    {
        if (SelectedUserProfile == null)
        {
            MessageBox.Show("Please select a user first.", "No User Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "CSV Files|*.csv|All Files|*.*",
            FileName = $"UserReport_{SelectedUserProfile.UserName.Replace(" ", "_")}_{DateTime.Now:yyyyMMdd}.csv",
            Title = "Export User Report"
        };

        if (dialog.ShowDialog() != true) return;

        try
        {
            var history = _localStorage.GetRecordsForDateRange(DateTime.Today.AddDays(-30), DateTime.Today)
                .Where(r => r.UserId == SelectedUserProfile.UserId && r.MachineId == SelectedUserProfile.MachineId)
                .ToList();
            
            _excelExport.ExportUserProfile(SelectedUserProfile, history, dialog.FileName);
            
            var result = MessageBox.Show($"Export successful!\n\nFile saved to:\n{dialog.FileName}\n\nWould you like to open the file?", 
                "Export Complete", MessageBoxButton.YesNo, MessageBoxImage.Information);
            
            if (result == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });

            AddDebug($"Exported user report to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    
    [RelayCommand]
    private void ExportUserActivity()
    {
        // Export current user's activity with hourly breakdown for selected date
        var userName = Environment.UserName;
        var userId = _currentWindowsUser;
        
        // If a specific user is selected in the grid, use that user instead
        if (SelectedDailyRecord != null)
        {
            userName = SelectedDailyRecord.UserName;
            userId = SelectedDailyRecord.UserId;
        }
        
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel Files|*.xlsx|All Files|*.*",
            FileName = $"UserActivity_{userName.Replace(" ", "_")}_{SelectedDate:yyyy-MM-dd}.xlsx",
            Title = "Export User Activity"
        };
        
        if (dialog.ShowDialog() != true) return;
        
        try
        {
            // Get activity breakdown
            DailyActivityBreakdown? breakdown = null;
            
            if (SelectedDate.Date == DateTime.Today && userId == _currentWindowsUser)
            {
                // Live data for current user today
                breakdown = _monitorService.GetDailyActivityBreakdown();
            }
            else
            {
                // Historical data
                var historicalData = _localStorage.GetUserDailyActivity(userId, SelectedDate);
                if (historicalData != null)
                {
                    breakdown = new DailyActivityBreakdown
                    {
                        Date = historicalData.Date,
                        RevitMinutes = historicalData.RevitMinutes,
                        MeetingMinutes = historicalData.MeetingMinutes,
                        TotalIdleMinutes = historicalData.IdleMinutes,
                        OtherMinutes = historicalData.OtherMinutes,
                        HourlyBreakdown = historicalData.HourlyBreakdown.Select(h => new HourlyActivity
                        {
                            Hour = h.Hour,
                            RevitMinutes = h.RevitMinutes,
                            MeetingMinutes = h.MeetingMinutes,
                            IdleMinutes = h.IdleMinutes,
                            OtherMinutes = h.OtherMinutes,
                            IsWorkHour = h.Hour >= 8 && h.Hour < 18
                        }).ToList()
                    };
                }
            }
            
            // Get daily record if available
            var record = DailyRecords.FirstOrDefault(r => r.UserId == userId);
            
            _excelExport.ExportUserActivity(userName, userId, SelectedDate, breakdown, record, dialog.FileName);
            
            var result = MessageBox.Show($"Export successful!\n\nFile saved to:\n{dialog.FileName}\n\nWould you like to open the file?", 
                "Export Complete", MessageBoxButton.YesNo, MessageBoxImage.Information);
            
            if (result == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
            
            AddDebug($"Exported user activity to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            AddDebug($"Export user activity error: {ex.Message}");
        }
    }

    [RelayCommand]
    private void ExportDateRangeReport()
    {
        // Simple date range picker using a dialog
        var startDate = SelectedDate.AddDays(-7);
        var endDate = SelectedDate;

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "CSV Files|*.csv|All Files|*.*",
            FileName = $"WeeklyReport_{startDate:yyyy-MM-dd}_to_{endDate:yyyy-MM-dd}.csv",
            Title = "Export Weekly Report"
        };

        if (dialog.ShowDialog() != true) return;

        try
        {
            var records = _localStorage.GetRecordsForDateRange(startDate, endDate);
            _excelExport.ExportAllUsersSummary(records, startDate, endDate, dialog.FileName);
            
            var result = MessageBox.Show($"Export successful!\n\nFile saved to:\n{dialog.FileName}\n\nWould you like to open the file?", 
                "Export Complete", MessageBoxButton.YesNo, MessageBoxImage.Information);
            
            if (result == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });

            AddDebug($"Exported weekly report to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void ExportCurrentStatus()
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "CSV Files|*.csv|All Files|*.*",
            FileName = $"CurrentStatus_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
            Title = "Export Current Status"
        };

        if (dialog.ShowDialog() != true) return;

        try
        {
            _excelExport.ExportCurrentStatus(CurrentSessions.ToList(), dialog.FileName);
            
            var result = MessageBox.Show($"Export successful!\n\nFile saved to:\n{dialog.FileName}\n\nWould you like to open the file?", 
                "Export Complete", MessageBoxButton.YesNo, MessageBoxImage.Information);
            
            if (result == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });

            AddDebug($"Exported current status to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void RefreshTimeTracking()
    {
        _ = LoadDailyRecordsAsync();
        LoadCalendarDays();
        UpdateWorkDayStatus();
    }
    
    [RelayCommand]
    private void ResetTodayTimeData()
    {
        if (!AdminService.Instance.IsAdminLoggedIn)
        {
            MessageBox.Show("Admin access required.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        var result = MessageBox.Show(
            "This will clear all time tracking data for today and start fresh.\n\n" +
            "Use this if hours are showing incorrectly.\n\n" +
            "Are you sure you want to continue?",
            "Reset Time Data",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);
        
        if (result == MessageBoxResult.Yes)
        {
            try
            {
                // Use the comprehensive cleanup method
                _localStorage.ClearTodayAndRestart();
                
                AddDebug("Reset today's time tracking data - starting fresh");
                
                // Refresh the display
                _ = LoadDailyRecordsAsync();
                UpdateWorkDayStatus();
                
                MessageBox.Show(
                    "Time tracking data for today has been reset.\n\n" +
                    "Time tracking will start fresh from now.",
                    "Reset Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Reset failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
    
    [RelayCommand]
    private void RecalculateSelectedDate()
    {
        if (!AdminService.Instance.IsAdminLoggedIn)
        {
            MessageBox.Show("Admin access required.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        try
        {
            // Recalculate data for the selected date
            _localStorage.RecalculateDailyData(SelectedDate);
            
            AddDebug($"Recalculated time data for {SelectedDate:yyyy-MM-dd}");
            
            // Refresh the display
            _ = LoadDailyRecordsAsync();
            
            MessageBox.Show(
                $"Time data for {SelectedDate:yyyy-MM-dd} has been recalculated.\n\n" +
                "Invalid sessions have been removed.",
                "Recalculation Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Recalculation failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    
    [RelayCommand]
    private async Task ExportDateRange()
    {
        if (!AdminService.Instance.IsAdminLoggedIn)
        {
            MessageBox.Show("Admin access required.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        // Show date range picker
        var result = MessageBox.Show(
            "Export report for:\n\n• Today (click Yes)\n• Last 7 days (click No)\n• Cancel (click Cancel)",
            "Select Date Range",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question);
        
        if (result == MessageBoxResult.Cancel) return;
        
        var endDate = DateTime.Today;
        var startDate = result == MessageBoxResult.Yes ? endDate : endDate.AddDays(-6);
        
        try
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                DefaultExt = ".csv",
                FileName = $"IDMonitor_Report_{startDate:yyyyMMdd}_{endDate:yyyyMMdd}.csv"
            };
            
            if (dialog.ShowDialog() == true)
            {
                AddDebug($"Exporting report to {dialog.FileName} for {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}");
                
                var csv = await _cloudService.ExportReportAsync(startDate, endDate, "all");
                
                if (!string.IsNullOrEmpty(csv))
                {
                    File.WriteAllText(dialog.FileName, csv);
                    MessageBox.Show($"Report exported to:\n{dialog.FileName}", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    await ExportLocalDataToCsv(dialog.FileName, startDate, endDate);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    
    [RelayCommand]
    private async Task ExportServerProjectReport()
    {
        if (!AdminService.Instance.IsAdminLoggedIn)
        {
            MessageBox.Show("Admin access required.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        try
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                DefaultExt = ".csv",
                FileName = $"IDMonitor_ServerProjectReport_{DateTime.Today:yyyyMMdd}.csv"
            };
            
            if (dialog.ShowDialog() == true)
            {
                var endDate = DateTime.Today;
                var startDate = endDate.AddDays(-6); // Last 7 days
                
                var csv = await _cloudService.ExportReportAsync(startDate, endDate, "project");
                
                if (!string.IsNullOrEmpty(csv))
                {
                    File.WriteAllText(dialog.FileName, csv);
                    MessageBox.Show($"Project report exported to:\n{dialog.FileName}", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Could not generate project report. Server may be unavailable.", "Export Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    
    [RelayCommand]
    private async Task DeleteHistoricalData()
    {
        if (!AdminService.Instance.IsAdminLoggedIn)
        {
            MessageBox.Show("Admin access required.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        var result = MessageBox.Show(
            $"⚠️ WARNING: This will permanently delete all time tracking data for {SelectedDate:yyyy-MM-dd}.\n\n" +
            "This action cannot be undone.\n\n" +
            "Are you sure you want to continue?",
            "Confirm Delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);
        
        if (result != MessageBoxResult.Yes) return;
        
        // Double confirmation
        var confirmResult = MessageBox.Show(
            $"Please confirm again:\n\n" +
            $"DELETE all data for {SelectedDate:yyyy-MM-dd}?",
            "Final Confirmation",
            MessageBoxButton.YesNo,
            MessageBoxImage.Stop);
        
        if (confirmResult != MessageBoxResult.Yes) return;
        
        try
        {
            // Delete from server
            var serverDeleted = await _cloudService.DeleteActivityDataAsync(SelectedDate);
            
            // Delete from local storage
            _localStorage.DeleteDailyData(SelectedDate);
            
            AddDebug($"Deleted data for {SelectedDate:yyyy-MM-dd}");
            
            // Refresh display
            await LoadDailyRecordsAsync();
            
            MessageBox.Show(
                $"Data for {SelectedDate:yyyy-MM-dd} has been deleted.\n\n" +
                $"Server: {(serverDeleted ? "✓ Deleted" : "✗ Not available")}\n" +
                "Local: ✓ Deleted",
                "Delete Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Delete failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    
    private async Task ExportLocalDataToCsv(string filePath, DateTime startDate, DateTime endDate)
    {
        var csv = new System.Text.StringBuilder();
        csv.AppendLine("Date,User,Machine,Revit Hours,Meeting Hours,Idle Hours,Other Hours,Overtime,Total Hours,Current Project");
        
        for (var date = startDate; date <= endDate; date = date.AddDays(1))
        {
            var records = _localStorage.GetDailyRecords(date);
            foreach (var record in records)
            {
                csv.AppendLine($"{record.Date:yyyy-MM-dd},\"{record.UserName}\",\"{record.MachineId}\",{record.TotalHours:F2},0,0,0,{record.OvertimeHours:F2},{record.TotalHours:F2},\"\"");
            }
        }
        
        await File.WriteAllTextAsync(filePath, csv.ToString());
        MessageBox.Show($"Report exported to:\n{filePath}", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
