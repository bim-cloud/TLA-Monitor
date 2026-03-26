using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace AutodeskIDMonitor.Services;

/// <summary>
/// Manages per-user credentials (email + hashed password) stored locally.
/// On first run the password is seeded from the DefaultUsers list in the WPF client
/// so users can log in immediately without admin involvement.
/// Dubai (UTC+4) is the canonical timezone for all time display.
/// </summary>
public class UserCredentialService
{
    // ── Dubai timezone (canonical for all TLA offices) ──────────────────────
    public static readonly TimeZoneInfo DubaiTz =
        TimeZoneInfo.FindSystemTimeZoneById("Arabian Standard Time");   // Windows tz id
    
    /// <summary>Convert any DateTime to Dubai local time for display.</summary>
    public static DateTime ToDubaiTime(DateTime utcOrLocal)
    {
        var utc = utcOrLocal.Kind == DateTimeKind.Local
            ? utcOrLocal.ToUniversalTime()
            : DateTime.SpecifyKind(utcOrLocal, DateTimeKind.Utc);
        return TimeZoneInfo.ConvertTimeFromUtc(utc, DubaiTz);
    }

    /// <summary>Current Dubai local time.</summary>
    public static DateTime NowDubai => TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, DubaiTz);

    // ── Singleton ────────────────────────────────────────────────────────────
    private static UserCredentialService? _instance;
    private static readonly object _lock = new();
    public static UserCredentialService Instance
    {
        get { lock (_lock) { return _instance ??= new UserCredentialService(); } }
    }

    // ── Storage ──────────────────────────────────────────────────────────────
    private readonly string _credFile;
    private Dictionary<string, UserCred> _creds = new(StringComparer.OrdinalIgnoreCase);

    private UserCredentialService()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TangentIDMonitor", "Data");
        Directory.CreateDirectory(dir);
        _credFile = Path.Combine(dir, "user_creds.json");
        Load();
        SeedDefaults();
    }

    // ── Default credentials seeded from the admin DefaultUsers list ─────────
    private static readonly List<(string Email, string Password)> Defaults = new()
    {
        ("abhilash.rajesh@tangentlandscape.com",      "Abhilash@123"),
        ("adhithyan.biju@tangentlandscape.com",       "TLA@2025"),
        ("afsal.badharu@tangentlandscape.com",        "TLA2024@209"),
        ("akshaya.jayakrishnan@tangentlandscape.com", "TLA2023_209"),
        ("amna.salim@tangentlandscape.com",           "TLA2023_209"),
        ("ananthu.unnikrishnan@tangentlandscape.com", "TLA20252101"),
        ("anshu.jalaludeen@tangentlandscape.com",     "TLA2024@213"),
        ("aparna.lakshmi@tangentlandscape.com",       "TLA2024@213"),
        ("athira.sivdas@tangentlandscape.com",        "TLA2025_209"),
        ("elbin.paulose@tangentlandscape.com",        "Tangent#2026"),
        ("jesto.joy@tangentlandscape.com",            "TLA2023@209"),
        ("jibin.issac@tangentlandscape.com",          "TLA2024_101"),
        ("jithin.pavithran@tangentlandscape.com",     "Tla2025@208"),
        ("jovanie.apa@tangentlandscape.com",          "Tangent@2024"),
        ("laura.cruz@tangentlandscape.com",           "TLA@2025"),
        ("lincy.kirubaharan@tangentlandscape.com",    "Tangent@123"),
        ("maznaz.firoz@tangentlandscape.com",         "TLA2025@101"),
        ("min.zaw@tangentlandscape.com",              "Tla.MIN@12345"),
        ("mohamed.asif@tangentlandscape.com",         "TLA2024@216"),
        ("narsha.abdura@tangentlandscape.com",        "Tangent@2026"),
        ("noufal.palliparambil@tangentlandscape.com", "TLA2024@209"),
        ("rahul.jain@tangentlandscape.com",           "Tangent@2025"),
        ("rashid.abdullah@tangentlandscape.com",      "Tangent@2025"),
        ("rivin.wilson@tangentlandscape.com",         "Tla2025@208"),
        ("sabarish.malayath@tangentlandscape.com",    "Tla123@2025"),
        ("safas.umar@tangentlandscape.com",           "TLA2024@215"),
        ("shasti.dharan@tangentlandscape.com",        "TLA2021@209"),
        ("syed.tahir@tangentlandscape.com",           "Tangent@2025"),
        ("toby.jose@tangentlandscape.com",            "Tangent@123"),
    };

    private void SeedDefaults()
    {
        bool changed = false;
        foreach (var (email, pwd) in Defaults)
        {
            var key = email.ToLower();
            if (!_creds.ContainsKey(key))
            {
                _creds[key] = new UserCred
                {
                    Email        = key,
                    DisplayName  = FormatName(key),
                    PasswordHash = Hash(pwd),
                    IsSetup      = false
                };
                changed = true;
            }
        }
        if (changed) Save();
    }

    // ── Public API ────────────────────────────────────────────────────────────
    public bool ValidateUser(string email, string password)
    {
        var key = email.ToLower();
        if (!_creds.TryGetValue(key, out var cred)) return false;
        return cred.PasswordHash.Equals(Hash(password), StringComparison.OrdinalIgnoreCase);
    }

    public string GetDisplayName(string email)
    {
        var key = email.ToLower();
        return _creds.TryGetValue(key, out var c) ? c.DisplayName : FormatName(key);
    }

    public bool IsSetup(string email) =>
        _creds.TryGetValue(email.ToLower(), out var c) && c.IsSetup;

    /// <summary>Called after successful first-time login to mark setup complete.</summary>
    public void SetupUser(string email, string password, string displayName)
    {
        var key = email.ToLower();
        if (!_creds.ContainsKey(key))
            _creds[key] = new UserCred { Email = key };

        _creds[key].DisplayName  = displayName;
        _creds[key].PasswordHash = Hash(password);
        _creds[key].IsSetup      = true;
        Save();
    }

    public bool ChangePassword(string email, string oldPwd, string newPwd)
    {
        var key = email.ToLower();
        if (!_creds.TryGetValue(key, out var cred)) return false;
        if (!cred.PasswordHash.Equals(Hash(oldPwd), StringComparison.OrdinalIgnoreCase)) return false;
        cred.PasswordHash = Hash(newPwd);
        Save();
        return true;
    }

    // ── Helpers ──────────────────────────────────────────────────────────────
    private static string Hash(string input)
    {
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(input)));
    }

    private static string FormatName(string email)
    {
        // "anshu.jalaludeen@tangentlandscape.com" → "Anshu Jalaludeen"
        var local = email.Split('@')[0];
        return string.Join(" ", local.Split('.').Select(p =>
            string.IsNullOrEmpty(p) ? p : char.ToUpper(p[0]) + p[1..]));
    }

    private void Load()
    {
        try
        {
            if (!File.Exists(_credFile)) return;
            var list = JsonSerializer.Deserialize<List<UserCred>>(File.ReadAllText(_credFile));
            if (list != null)
                _creds = list.ToDictionary(c => c.Email.ToLower(), StringComparer.OrdinalIgnoreCase);
        }
        catch { }
    }

    private void Save()
    {
        try
        {
            var json = JsonSerializer.Serialize(_creds.Values.ToList(),
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_credFile, json);
        }
        catch { }
    }

    private class UserCred
    {
        public string Email        { get; set; } = "";
        public string DisplayName  { get; set; } = "";
        public string PasswordHash { get; set; } = "";
        public bool   IsSetup      { get; set; } = false;
    }
}
