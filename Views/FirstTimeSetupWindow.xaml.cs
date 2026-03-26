using System.Windows;
using System.Windows.Controls;

namespace AutodeskIDMonitor.Views;

public partial class FirstTimeSetupWindow : Window
{
    public string EnteredName { get; private set; } = "";
    public string EnteredEmail { get; private set; } = "";
    public bool SetupCompleted { get; private set; } = false;
    
    public FirstTimeSetupWindow(string suggestedName = "", string suggestedEmail = "")
    {
        InitializeComponent();
        
        // Pre-fill if we have suggestions
        if (!string.IsNullOrEmpty(suggestedName) && suggestedName.ToLower() != "user")
        {
            NameTextBox.Text = suggestedName;
        }
        
        if (!string.IsNullOrEmpty(suggestedEmail))
        {
            EmailTextBox.Text = suggestedEmail;
        }
        
        UpdatePlaceholders();
    }
    
    private void NameTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdatePlaceholders();
        ValidateInput();
    }
    
    private void EmailTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdatePlaceholders();
    }
    
    private void UpdatePlaceholders()
    {
        PlaceholderText.Visibility = string.IsNullOrEmpty(NameTextBox.Text) 
            ? Visibility.Visible : Visibility.Collapsed;
        EmailPlaceholderText.Visibility = string.IsNullOrEmpty(EmailTextBox.Text) 
            ? Visibility.Visible : Visibility.Collapsed;
    }
    
    private void ValidateInput()
    {
        var name = NameTextBox.Text.Trim();
        
        if (string.IsNullOrEmpty(name))
        {
            ContinueButton.IsEnabled = false;
            ValidationMessage.Visibility = Visibility.Collapsed;
            return;
        }
        
        if (name.Length < 2)
        {
            ContinueButton.IsEnabled = false;
            ValidationMessage.Text = "Name must be at least 2 characters";
            ValidationMessage.Visibility = Visibility.Visible;
            return;
        }
        
        if (name.ToLower() == "user" || name.ToLower() == "admin" || name.ToLower() == "administrator")
        {
            ContinueButton.IsEnabled = false;
            ValidationMessage.Text = "Please enter your actual name";
            ValidationMessage.Visibility = Visibility.Visible;
            return;
        }
        
        ContinueButton.IsEnabled = true;
        ValidationMessage.Visibility = Visibility.Collapsed;
    }
    
    private void ContinueButton_Click(object sender, RoutedEventArgs e)
    {
        EnteredName = NameTextBox.Text.Trim();
        EnteredEmail = EmailTextBox.Text.Trim();
        
        // Auto-correct email format if needed
        if (!string.IsNullOrEmpty(EnteredEmail))
        {
            EnteredEmail = AutoCorrectEmail(EnteredEmail);
        }
        
        SetupCompleted = true;
        DialogResult = true;
        Close();
    }
    
    /// <summary>
    /// Auto-correct email format - convert partial names to full email
    /// </summary>
    private string AutoCorrectEmail(string input)
    {
        input = input.Trim().ToLower();
        
        // If already a valid email, return as-is
        if (input.Contains("@") && input.Contains("."))
        {
            return input;
        }
        
        // Try to construct email from name
        // Remove spaces and special characters
        var cleanName = System.Text.RegularExpressions.Regex.Replace(input, @"[^a-z0-9.]", "");
        
        // Common patterns:
        // "amnasalim" -> "amna.salim@tangentlandscape.com"
        // "john smith" -> "john.smith@tangentlandscape.com"
        
        // If it looks like a concatenated name (no dots), try to split
        if (!cleanName.Contains("."))
        {
            // Try common name splitting patterns
            // This is a simple heuristic - split at common name boundaries
            var formatted = TryFormatAsEmail(cleanName);
            if (!string.IsNullOrEmpty(formatted))
            {
                return formatted + "@tangentlandscape.com";
            }
        }
        
        // Default: add domain
        if (!cleanName.Contains("@"))
        {
            return cleanName + "@tangentlandscape.com";
        }
        
        return input;
    }
    
    private string TryFormatAsEmail(string name)
    {
        if (string.IsNullOrEmpty(name) || name.Length < 4)
            return name;
        
        // Common first names to help with splitting
        string[] commonFirstNames = { "amna", "anshu", "john", "jane", "mike", "sarah", "david", "lisa", 
            "ahmed", "mohammed", "fatima", "ali", "omar", "layla", "noor", "hassan", "hussein",
            "abinshan", "adhithyan", "afsal", "akshaya", "ananthu", "aparna", "athira", "jesto",
            "laura", "mohamed", "rajeev" };
        
        foreach (var firstName in commonFirstNames)
        {
            if (name.StartsWith(firstName) && name.Length > firstName.Length)
            {
                var lastName = name.Substring(firstName.Length);
                return $"{firstName}.{lastName}";
            }
        }
        
        return name;
    }
}
