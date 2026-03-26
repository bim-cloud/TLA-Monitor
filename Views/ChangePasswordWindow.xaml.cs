using System.Windows;
using AutodeskIDMonitor.Services;

namespace AutodeskIDMonitor.Views;

public partial class ChangePasswordWindow : Window
{
    public ChangePasswordWindow()
    {
        InitializeComponent();
        CurrentPasswordBox.Focus();
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        var currentPassword = CurrentPasswordBox.Password;
        var newPassword = NewPasswordBox.Password;
        var confirmPassword = ConfirmPasswordBox.Password;

        // Validation
        if (string.IsNullOrEmpty(currentPassword))
        {
            ShowError("Please enter your current password");
            return;
        }

        if (string.IsNullOrEmpty(newPassword))
        {
            ShowError("Please enter a new password");
            return;
        }

        if (newPassword.Length < 4)
        {
            ShowError("New password must be at least 4 characters");
            return;
        }

        if (newPassword != confirmPassword)
        {
            ShowError("New passwords do not match");
            return;
        }

        // Attempt to change password
        if (AdminService.Instance.ChangePassword(currentPassword, newPassword))
        {
            MessageBox.Show("Password changed successfully!", "Success", 
                MessageBoxButton.OK, MessageBoxImage.Information);
            DialogResult = true;
            Close();
        }
        else
        {
            ShowError("Current password is incorrect");
            CurrentPasswordBox.SelectAll();
            CurrentPasswordBox.Focus();
        }
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void ShowError(string message)
    {
        ErrorText.Text = message;
        ErrorText.Visibility = Visibility.Visible;
    }
}
