using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace AutodeskIDMonitor.Services;

public class AdminService
{
    private const string DEFAULT_PASSWORD = "admin";
    private const string CONFIG_FILENAME = "admin_config.json";
    
    private static AdminService? _instance;
    private static readonly object _lock = new();
    
    private string _passwordHash = "";
    private bool _isAdminLoggedIn = false;
    private readonly string _configPath;

    public static AdminService Instance
    {
        get
        {
            if (_instance == null)
            {
                lock (_lock) { _instance ??= new AdminService(); }
            }
            return _instance;
        }
    }

    public event EventHandler<bool>? AdminStatusChanged;
    public bool IsAdminLoggedIn => _isAdminLoggedIn;

    private AdminService()
    {
        var appDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AutodeskIDMonitor");
        Directory.CreateDirectory(appDataPath);
        _configPath = Path.Combine(appDataPath, CONFIG_FILENAME);
    }

    public void Initialize() => LoadConfig();

    private void LoadConfig()
    {
        try
        {
            if (File.Exists(_configPath))
            {
                var json = File.ReadAllText(_configPath);
                var config = JsonSerializer.Deserialize<Models.AdminConfig>(json);
                if (config != null && !string.IsNullOrEmpty(config.PasswordHash))
                {
                    _passwordHash = config.PasswordHash;
                    return;
                }
            }
        }
        catch { }
        _passwordHash = ComputeHash(DEFAULT_PASSWORD);
        SaveConfig();
    }

    private void SaveConfig()
    {
        try
        {
            var config = new Models.AdminConfig { PasswordHash = _passwordHash };
            var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_configPath, json);
        }
        catch { }
    }

    public bool Login(string password)
    {
        if (string.IsNullOrEmpty(password)) return false;
        var hash = ComputeHash(password);
        if (hash.Equals(_passwordHash, StringComparison.OrdinalIgnoreCase))
        {
            _isAdminLoggedIn = true;
            AdminStatusChanged?.Invoke(this, true);
            return true;
        }
        return false;
    }

    public void Logout()
    {
        _isAdminLoggedIn = false;
        AdminStatusChanged?.Invoke(this, false);
    }

    public bool ChangePassword(string currentPassword, string newPassword)
    {
        if (string.IsNullOrEmpty(newPassword) || newPassword.Length < 4) return false;
        if (!ComputeHash(currentPassword).Equals(_passwordHash, StringComparison.OrdinalIgnoreCase)) return false;
        _passwordHash = ComputeHash(newPassword);
        SaveConfig();
        return true;
    }

    private static string ComputeHash(string input)
    {
        using var sha256 = SHA256.Create();
        var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(input));
        return Convert.ToHexString(bytes);
    }
}
