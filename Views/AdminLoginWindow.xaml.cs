using System.Windows;
using System.Windows.Input;
using AutodeskIDMonitor.Services;

namespace AutodeskIDMonitor.Views;

public partial class AdminLoginWindow : Window
{
    private int _attemptCount = 0;
    private const int MaxAttempts = 3;

    public AdminLoginWindow()
    {
        InitializeComponent();
        PasswordBox.Focus();
    }

    private void PasswordBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) AttemptLogin();
    }

    private void Login_Click(object sender, RoutedEventArgs e) => AttemptLogin();

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void AttemptLogin()
    {
        var password = PasswordBox.Password;
        if (string.IsNullOrEmpty(password)) { ShowError("Please enter a password"); return; }

        _attemptCount++;
        if (AdminService.Instance.Login(password))
        {
            DialogResult = true;
            Close();
        }
        else
        {
            ShowError("Invalid password");
            AttemptText.Text = $"Attempt {_attemptCount} of {MaxAttempts}";
            AttemptText.Visibility = Visibility.Visible;
            PasswordBox.SelectAll();
            PasswordBox.Focus();
            if (_attemptCount >= MaxAttempts)
            {
                MessageBox.Show("Too many failed attempts. Please try again later.",
                    "Login Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                DialogResult = false;
                Close();
            }
        }
    }

    private void ShowError(string message)
    {
        ErrorText.Text = message;
        ErrorText.Visibility = Visibility.Visible;
    }
}
