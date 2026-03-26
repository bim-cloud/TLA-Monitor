using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using AutodeskIDMonitor.Services;

namespace AutodeskIDMonitor.Views;

public partial class UserLoginWindow : Window
{
    private readonly bool _isFirstTime;

    public string LoggedInEmail { get; private set; } = "";
    public string LoggedInName  { get; private set; } = "";
    public bool   LoginSuccess  { get; private set; } = false;

    public UserLoginWindow(bool isFirstTime = false, string prefillEmail = "", string prefillName = "")
    {
        InitializeComponent();
        _isFirstTime = isFirstTime;

        if (isFirstTime)
        {
            NamePanel.Visibility = Visibility.Visible;
            if (!string.IsNullOrEmpty(prefillName)) NameBox.Text = prefillName;
        }

        if (!string.IsNullOrEmpty(prefillEmail)) EmailBox.Text = prefillEmail;

        // Focus first empty field
        if (isFirstTime && string.IsNullOrEmpty(NameBox.Text))
            Loaded += (_, _) => NameBox.Focus();
        else if (string.IsNullOrEmpty(EmailBox.Text))
            Loaded += (_, _) => EmailBox.Focus();
        else
            Loaded += (_, _) => PasswordBox.Focus();
    }

    private void EmailBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        ErrorText.Visibility = Visibility.Collapsed;
    }

    private void PasswordBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) AttemptLogin();
    }

    private void SignIn_Click(object sender, RoutedEventArgs e) => AttemptLogin();

    private void AttemptLogin()
    {
        ErrorText.Visibility = Visibility.Collapsed;
        ResetMessage.Visibility = Visibility.Collapsed;

        var email    = EmailBox.Text.Trim().ToLower();
        var password = PasswordBox.Password;
        var name     = _isFirstTime ? NameBox.Text.Trim() : "";

        // Validate
        if (_isFirstTime && string.IsNullOrEmpty(name))
        { ShowError("Please enter your full name."); NameBox.Focus(); return; }

        if (string.IsNullOrEmpty(email))
        { ShowError("Please enter your work email."); EmailBox.Focus(); return; }

        if (!email.Contains("@"))
        { ShowError("Please enter a valid email address."); EmailBox.Focus(); return; }

        if (string.IsNullOrEmpty(password))
        { ShowError("Please enter your password."); PasswordBox.Focus(); return; }

        // Auto-complete @tangentlandscape.com if only name part typed
        if (!email.Contains("@"))
            email = email + "@tangentlandscape.com";

        // Validate against stored user credentials
        if (!UserCredentialService.Instance.ValidateUser(email, password))
        {
            ShowError("Incorrect email or password. Please try again.");
            PasswordBox.SelectAll();
            PasswordBox.Focus();
            return;
        }

        LoggedInEmail = email;
        LoggedInName  = _isFirstTime ? name : UserCredentialService.Instance.GetDisplayName(email);
        LoginSuccess  = true;

        if (_isFirstTime)
            UserCredentialService.Instance.SetupUser(email, password, name);

        DialogResult = true;
        Close();
    }

    private void ForgotPassword_Click(object sender, RoutedEventArgs e)
    {
        var email = EmailBox.Text.Trim().ToLower();
        if (string.IsNullOrEmpty(email) || !email.Contains("@"))
        {
            ShowError("Enter your work email above, then click Forgot password.");
            EmailBox.Focus();
            return;
        }

        // Send reset email via mailto (opens Outlook)
        var subject = Uri.EscapeDataString("Autodesk ID Monitor – Password Reset Request");
        var body    = Uri.EscapeDataString(
            $"Hi IT Team,\n\nI need to reset my Autodesk ID Monitor password.\n\nEmail: {email}\n\nPlease send me a new temporary password.\n\nThank you.");
        var mailto  = $"mailto:it@tangentlandscape.com?subject={subject}&body={body}";

        try
        {
            Process.Start(new ProcessStartInfo(mailto) { UseShellExecute = true });
            ResetMessage.Text = $"✓ A reset request has been opened in Outlook for {email}. The IT team will send you a new password.";
            ResetMessage.Visibility = Visibility.Visible;
            ErrorText.Visibility    = Visibility.Collapsed;
        }
        catch
        {
            ShowError("Could not open Outlook. Please email it@tangentlandscape.com directly.");
        }
    }

    private void ShowError(string msg)
    {
        ErrorText.Text       = msg;
        ErrorText.Visibility = Visibility.Visible;
    }
}
