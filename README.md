# Autodesk ID Monitor v2.0.0 - Build Ready

## Quick Build

### Option 1: Visual Studio 2022
1. Open `AutodeskIDMonitor.sln`
2. Select `Release` configuration
3. Build → Build Solution (Ctrl+Shift+B)
4. Output: `bin\Release\net8.0-windows\`

### Option 2: Command Line
```cmd
dotnet restore
dotnet build -c Release
```

### Option 3: Publish Self-Contained EXE
```cmd
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```
Or run: `Publish_SelfContained.bat`

---

## Requirements
- Windows 10/11 (64-bit)
- .NET 8.0 SDK
- Visual Studio 2022 (recommended)

---

## Configuration

After first run, go to Settings tab and configure:

| Setting | Value |
|---------|-------|
| Cloud API URL | `http://141.145.153.32:5000` |
| API Key | `Tangent@2026` |
| Country | `UAE` |
| Office | `Dubai` |

Config file location: `%LOCALAPPDATA%\AutodeskIDMonitor\config.json`

---

## Features

- ✅ Real-time Autodesk login monitoring
- ✅ Revit project tracking with version detection
- ✅ Meeting detection (Teams, Zoom, Webex)
- ✅ Idle time tracking
- ✅ Activity breakdown charts
- ✅ Admin dashboard with date selection
- ✅ Excel export
- ✅ Auto-update support
- ✅ System tray with minimize

---

## Server

Server URL: http://141.145.153.32:5000
Server Version: Flask v2.0.0

Test connection:
```
http://141.145.153.32:5000/api/health
```

---

## Troubleshooting

### App won't start
Run `CleanAndBuild.bat` then rebuild.

### Connection issues
1. Verify server URL includes `http://`
2. Check API key matches server
3. Test: `curl http://141.145.153.32:5000/api/health`

### Autodesk not detected
- Ensure Autodesk software is installed
- Check `%LOCALAPPDATA%\Autodesk\.autodesk\loginstate.json` exists

---

## Version History

- v2.0.0 - Flask server integration, improved activity tracking
- v1.8.0 - Auto-update, system tray
- v1.5.0 - Activity charts, meeting detection
