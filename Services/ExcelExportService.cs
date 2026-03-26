using System.IO;
using System.Linq;
using System.Text;
using AutodeskIDMonitor.Models;
using ClosedXML.Excel;

namespace AutodeskIDMonitor.Services;

public class ExcelExportService
{
    // Tangent brand colors
    private const string TANGENT_CYAN = "#00B4D8";
    private const string TANGENT_DARK = "#333333";
    private const string HEADER_BG = "#E3F2FD";
    private const string OVERTIME_COLOR = "#E65100";
    private const string SUCCESS_COLOR = "#4CAF50";

    /// <summary>
    /// Export daily records to professional Excel format
    /// </summary>
    public string ExportDailyRecords(List<DailyWorkRecord> records, string filePath)
    {
        // Ensure xlsx extension
        filePath = Path.ChangeExtension(filePath, ".xlsx");
        
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Daily Report");
        
        // Header with company name
        ws.Cell("A1").Value = "TANGENT LANDSCAPE ARCHITECTURE";
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 16;
        ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Range("A1:K1").Merge();
        
        ws.Cell("A2").Value = $"Daily Time Tracking Report - Generated: {DateTime.Now:yyyy-MM-dd HH:mm}";
        ws.Cell("A2").Style.Font.Italic = true;
        ws.Range("A2:K2").Merge();
        
        // Headers row
        int headerRow = 4;
        var headers = new[] { "Date", "User Name", "User ID", "Machine ID", "First Activity", "Last Activity", 
                             "Total Hours", "Regular Hours", "Overtime", "Projects Worked", "Project Count" };
        
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(headerRow, i + 1);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
        
        // Data rows
        int row = headerRow + 1;
        foreach (var record in records.OrderBy(r => r.Date).ThenBy(r => r.UserName))
        {
            ws.Cell(row, 1).Value = record.Date;
            ws.Cell(row, 1).Style.DateFormat.Format = "yyyy-MM-dd";
            ws.Cell(row, 2).Value = record.UserName;
            ws.Cell(row, 3).Value = record.UserId;
            ws.Cell(row, 4).Value = record.MachineId;
            ws.Cell(row, 5).Value = record.FirstActivity;
            ws.Cell(row, 5).Style.DateFormat.Format = "HH:mm:ss";
            ws.Cell(row, 6).Value = record.LastActivity;
            ws.Cell(row, 6).Style.DateFormat.Format = "HH:mm:ss";
            ws.Cell(row, 7).Value = record.TotalHours;
            ws.Cell(row, 7).Style.NumberFormat.Format = "0.0";
            ws.Cell(row, 7).Style.Font.Bold = true;
            ws.Cell(row, 8).Value = record.RegularHours;
            ws.Cell(row, 8).Style.NumberFormat.Format = "0.0";
            ws.Cell(row, 9).Value = record.OvertimeHours;
            ws.Cell(row, 9).Style.NumberFormat.Format = "0.0";
            if (record.OvertimeHours > 0)
                ws.Cell(row, 9).Style.Font.FontColor = XLColor.FromHtml(OVERTIME_COLOR);
            ws.Cell(row, 10).Value = record.ProjectsWorked;
            ws.Cell(row, 11).Value = record.ProjectCount;
            
            // Alternate row colors
            if (row % 2 == 0)
                ws.Range(row, 1, row, 11).Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
            
            ws.Range(row, 1, row, 11).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            row++;
        }
        
        // Summary section
        row += 2;
        ws.Cell(row, 1).Value = "SUMMARY";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Font.FontSize = 12;
        ws.Cell(row, 1).Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
        
        row++;
        ws.Cell(row, 1).Value = "Total Users:";
        ws.Cell(row, 2).Value = records.Select(r => r.UserName).Distinct().Count();
        row++;
        ws.Cell(row, 1).Value = "Total Hours:";
        ws.Cell(row, 2).Value = records.Sum(r => r.TotalHours);
        ws.Cell(row, 2).Style.NumberFormat.Format = "0.0";
        ws.Cell(row, 2).Style.Font.Bold = true;
        row++;
        ws.Cell(row, 1).Value = "Total Overtime:";
        ws.Cell(row, 2).Value = records.Sum(r => r.OvertimeHours);
        ws.Cell(row, 2).Style.NumberFormat.Format = "0.0";
        ws.Cell(row, 2).Style.Font.FontColor = XLColor.FromHtml(OVERTIME_COLOR);
        
        ws.Columns().AdjustToContents();
        ws.Column(10).Width = 40;
        
        workbook.SaveAs(filePath);
        return filePath;
    }

    /// <summary>
    /// Export project summaries to professional Excel format
    /// </summary>
    public string ExportProjectSummaries(List<ProjectWorkSummary> summaries, DateTime date, string filePath)
    {
        filePath = Path.ChangeExtension(filePath, ".xlsx");
        
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Project Report");
        
        // Header
        ws.Cell("A1").Value = "TANGENT LANDSCAPE ARCHITECTURE";
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 16;
        ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Range("A1:E1").Merge();
        
        ws.Cell("A2").Value = $"Project Activity Report - {date:yyyy-MM-dd}";
        ws.Cell("A2").Style.Font.Italic = true;
        ws.Range("A2:E2").Merge();
        
        // Project Summary section
        int row = 4;
        ws.Cell(row, 1).Value = "PROJECT SUMMARY";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Font.FontSize = 12;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
        ws.Range(row, 1, row, 5).Merge();
        
        row++;
        var projectHeaders = new[] { "Project Name", "Total Hours", "Users", "Who Worked", "Last Activity" };
        for (int i = 0; i < projectHeaders.Length; i++)
        {
            var cell = ws.Cell(row, i + 1);
            cell.Value = projectHeaders[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
        
        row++;
        foreach (var summary in summaries.OrderByDescending(s => s.TotalHours))
        {
            ws.Cell(row, 1).Value = summary.ProjectName;
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 2).Value = summary.TotalHours;
            ws.Cell(row, 2).Style.NumberFormat.Format = "0.0";
            ws.Cell(row, 2).Style.Font.Bold = true;
            ws.Cell(row, 2).Style.Font.FontColor = XLColor.FromHtml(SUCCESS_COLOR);
            ws.Cell(row, 3).Value = summary.UserCount;
            ws.Cell(row, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(row, 4).Value = summary.UsersWorking;
            ws.Cell(row, 5).Value = summary.LastActivity;
            ws.Cell(row, 5).Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
            
            ws.Range(row, 1, row, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            row++;
        }
        
        // User Details by Project
        row += 2;
        ws.Cell(row, 1).Value = "USER DETAILS BY PROJECT";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Font.FontSize = 12;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
        ws.Range(row, 1, row, 5).Merge();
        
        row++;
        var userHeaders = new[] { "Project", "User Name", "User ID", "Hours", "Last Activity" };
        for (int i = 0; i < userHeaders.Length; i++)
        {
            var cell = ws.Cell(row, i + 1);
            cell.Value = userHeaders[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
        
        row++;
        foreach (var summary in summaries.OrderBy(s => s.ProjectName))
        {
            foreach (var user in summary.UserTimes.OrderByDescending(u => u.Hours))
            {
                ws.Cell(row, 1).Value = summary.ProjectName;
                ws.Cell(row, 2).Value = user.UserName;
                ws.Cell(row, 3).Value = user.UserId;
                ws.Cell(row, 4).Value = user.Hours;
                ws.Cell(row, 4).Style.NumberFormat.Format = "0.0";
                ws.Cell(row, 5).Value = user.LastActivity;
                ws.Cell(row, 5).Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
                
                ws.Range(row, 1, row, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                row++;
            }
        }
        
        ws.Columns().AdjustToContents();
        ws.Column(1).Width = 35;
        ws.Column(4).Width = 35;
        
        workbook.SaveAs(filePath);
        return filePath;
    }

    /// <summary>
    /// Export user profile with history to professional Excel format
    /// </summary>
    public string ExportUserProfile(UserDetailedProfile profile, List<DailyWorkRecord> history, string filePath)
    {
        filePath = Path.ChangeExtension(filePath, ".xlsx");
        
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("User Profile");
        
        ws.Cell("A1").Value = "TANGENT LANDSCAPE ARCHITECTURE";
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 16;
        ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Range("A1:G1").Merge();
        
        ws.Cell("A2").Value = $"User Profile Report - {profile.UserName}";
        ws.Cell("A2").Style.Font.Italic = true;
        ws.Range("A2:G2").Merge();
        
        int row = 4;
        ws.Cell(row, 1).Value = "USER INFORMATION";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
        ws.Range(row, 1, row, 4).Merge();
        
        row++;
        var infoLabels = new[] { "User Name", "User ID", "Machine ID", "Autodesk Email", 
                                "Status", "Last Online", "Today Hours", "Today Overtime", 
                                "Week Hours", "Month Hours" };
        var infoValues = new object[] { 
            profile.UserName, profile.UserId, profile.MachineId, profile.AutodeskEmail,
            profile.IsCurrentlyOnline ? "🟢 Online" : "⬤ Offline", 
            profile.LastOnline.ToString("yyyy-MM-dd HH:mm"),
            profile.TodayHours, profile.TodayOvertime, profile.WeekHours, profile.MonthHours
        };
        
        for (int i = 0; i < infoLabels.Length; i++)
        {
            ws.Cell(row, 1).Value = infoLabels[i];
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
            ws.Cell(row, 2).Value = infoValues[i]?.ToString() ?? "";
            if (infoLabels[i].Contains("Hours") && infoValues[i] is double d)
            {
                ws.Cell(row, 2).Style.NumberFormat.Format = "0.0";
                if (infoLabels[i].Contains("Overtime") && d > 0)
                    ws.Cell(row, 2).Style.Font.FontColor = XLColor.FromHtml(OVERTIME_COLOR);
            }
            row++;
        }
        
        row += 2;
        ws.Cell(row, 1).Value = "DAILY HISTORY";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
        ws.Range(row, 1, row, 7).Merge();
        
        row++;
        var historyHeaders = new[] { "Date", "Total Hours", "Regular Hours", "Overtime", "Projects", "First Activity", "Last Activity" };
        for (int i = 0; i < historyHeaders.Length; i++)
        {
            var cell = ws.Cell(row, i + 1);
            cell.Value = historyHeaders[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
        
        row++;
        foreach (var record in history.OrderByDescending(r => r.Date))
        {
            ws.Cell(row, 1).Value = record.Date;
            ws.Cell(row, 1).Style.DateFormat.Format = "yyyy-MM-dd";
            ws.Cell(row, 2).Value = record.TotalHours;
            ws.Cell(row, 2).Style.NumberFormat.Format = "0.0";
            ws.Cell(row, 2).Style.Font.Bold = true;
            ws.Cell(row, 3).Value = record.RegularHours;
            ws.Cell(row, 3).Style.NumberFormat.Format = "0.0";
            ws.Cell(row, 4).Value = record.OvertimeHours;
            ws.Cell(row, 4).Style.NumberFormat.Format = "0.0";
            if (record.OvertimeHours > 0)
                ws.Cell(row, 4).Style.Font.FontColor = XLColor.FromHtml(OVERTIME_COLOR);
            ws.Cell(row, 5).Value = record.ProjectsWorked;
            ws.Cell(row, 6).Value = record.FirstActivity;
            ws.Cell(row, 6).Style.DateFormat.Format = "HH:mm";
            ws.Cell(row, 7).Value = record.LastActivity;
            ws.Cell(row, 7).Style.DateFormat.Format = "HH:mm";
            
            ws.Range(row, 1, row, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            row++;
        }
        
        ws.Columns().AdjustToContents();
        ws.Column(5).Width = 40;
        
        workbook.SaveAs(filePath);
        return filePath;
    }

    /// <summary>
    /// Export all users summary for a date range to professional Excel format
    /// </summary>
    public string ExportAllUsersSummary(List<DailyWorkRecord> records, DateTime startDate, DateTime endDate, string filePath)
    {
        filePath = Path.ChangeExtension(filePath, ".xlsx");
        
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Weekly Summary");
        
        ws.Cell("A1").Value = "TANGENT LANDSCAPE ARCHITECTURE";
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 16;
        ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Range("A1:H1").Merge();
        
        ws.Cell("A2").Value = $"Activity Summary Report - {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}";
        ws.Cell("A2").Style.Font.Italic = true;
        ws.Range("A2:H2").Merge();
        
        int row = 4;
        ws.Cell(row, 1).Value = "SUMMARY BY USER";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
        ws.Range(row, 1, row, 8).Merge();
        
        row++;
        var summaryHeaders = new[] { "User Name", "User ID", "Machine ID", "Days Worked", "Total Hours", "Regular Hours", "Overtime", "Avg Hours/Day" };
        for (int i = 0; i < summaryHeaders.Length; i++)
        {
            var cell = ws.Cell(row, i + 1);
            cell.Value = summaryHeaders[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
        
        row++;
        var userGroups = records.GroupBy(r => new { r.UserId, r.MachineId, r.UserName });
        double grandTotalHours = 0;
        double grandTotalOvertime = 0;
        
        foreach (var group in userGroups.OrderBy(g => g.Key.UserName))
        {
            var totalHours = group.Sum(r => r.TotalHours);
            var regularHours = group.Sum(r => r.RegularHours);
            var overtimeHours = group.Sum(r => r.OvertimeHours);
            var days = group.Count();
            grandTotalHours += totalHours;
            grandTotalOvertime += overtimeHours;
            
            ws.Cell(row, 1).Value = group.Key.UserName;
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 2).Value = group.Key.UserId;
            ws.Cell(row, 3).Value = group.Key.MachineId;
            ws.Cell(row, 4).Value = days;
            ws.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(row, 5).Value = totalHours;
            ws.Cell(row, 5).Style.NumberFormat.Format = "0.0";
            ws.Cell(row, 5).Style.Font.Bold = true;
            ws.Cell(row, 5).Style.Font.FontColor = XLColor.FromHtml(SUCCESS_COLOR);
            ws.Cell(row, 6).Value = regularHours;
            ws.Cell(row, 6).Style.NumberFormat.Format = "0.0";
            ws.Cell(row, 7).Value = overtimeHours;
            ws.Cell(row, 7).Style.NumberFormat.Format = "0.0";
            if (overtimeHours > 0)
                ws.Cell(row, 7).Style.Font.FontColor = XLColor.FromHtml(OVERTIME_COLOR);
            ws.Cell(row, 8).Value = days > 0 ? totalHours / days : 0;
            ws.Cell(row, 8).Style.NumberFormat.Format = "0.0";
            
            ws.Range(row, 1, row, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            row++;
        }
        
        // Grand total
        row++;
        ws.Cell(row, 1).Value = "GRAND TOTAL";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Range(row, 1, row, 4).Merge();
        ws.Cell(row, 5).Value = grandTotalHours;
        ws.Cell(row, 5).Style.NumberFormat.Format = "0.0";
        ws.Cell(row, 5).Style.Font.Bold = true;
        ws.Cell(row, 5).Style.Font.FontSize = 14;
        ws.Cell(row, 5).Style.Font.FontColor = XLColor.FromHtml(SUCCESS_COLOR);
        ws.Cell(row, 7).Value = grandTotalOvertime;
        ws.Cell(row, 7).Style.NumberFormat.Format = "0.0";
        ws.Cell(row, 7).Style.Font.Bold = true;
        if (grandTotalOvertime > 0)
            ws.Cell(row, 7).Style.Font.FontColor = XLColor.FromHtml(OVERTIME_COLOR);
        ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
        ws.Range(row, 1, row, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        
        // Daily details
        row += 3;
        ws.Cell(row, 1).Value = "DAILY DETAILS";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
        ws.Range(row, 1, row, 7).Merge();
        
        row++;
        var detailHeaders = new[] { "Date", "User Name", "User ID", "Machine ID", "Total Hours", "Overtime", "Projects" };
        for (int i = 0; i < detailHeaders.Length; i++)
        {
            var cell = ws.Cell(row, i + 1);
            cell.Value = detailHeaders[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
        
        row++;
        foreach (var record in records.OrderBy(r => r.Date).ThenBy(r => r.UserName))
        {
            ws.Cell(row, 1).Value = record.Date;
            ws.Cell(row, 1).Style.DateFormat.Format = "yyyy-MM-dd";
            ws.Cell(row, 2).Value = record.UserName;
            ws.Cell(row, 3).Value = record.UserId;
            ws.Cell(row, 4).Value = record.MachineId;
            ws.Cell(row, 5).Value = record.TotalHours;
            ws.Cell(row, 5).Style.NumberFormat.Format = "0.0";
            ws.Cell(row, 6).Value = record.OvertimeHours;
            ws.Cell(row, 6).Style.NumberFormat.Format = "0.0";
            if (record.OvertimeHours > 0)
                ws.Cell(row, 6).Style.Font.FontColor = XLColor.FromHtml(OVERTIME_COLOR);
            ws.Cell(row, 7).Value = record.ProjectsWorked;
            
            ws.Range(row, 1, row, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            row++;
        }
        
        ws.Columns().AdjustToContents();
        ws.Column(7).Width = 40;
        
        workbook.SaveAs(filePath);
        return filePath;
    }

    /// <summary>
    /// Export current session status to professional Excel format
    /// </summary>
    public string ExportCurrentStatus(List<UserSession> sessions, string filePath)
    {
        filePath = Path.ChangeExtension(filePath, ".xlsx");
        
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Current Status");
        
        ws.Cell("A1").Value = "TANGENT LANDSCAPE ARCHITECTURE";
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 16;
        ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Range("A1:H1").Merge();
        
        ws.Cell("A2").Value = $"Current Status Report - {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
        ws.Cell("A2").Style.Font.Italic = true;
        ws.Range("A2:H2").Merge();
        
        int row = 4;
        ws.Cell(row, 1).Value = "SUMMARY";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
        ws.Range(row, 1, row, 4).Merge();
        
        row++;
        ws.Cell(row, 1).Value = "Total Users:";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 2).Value = sessions.Count;
        row++;
        ws.Cell(row, 1).Value = "Online:";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 2).Value = sessions.Count(s => s.IsLoggedIn);
        ws.Cell(row, 2).Style.Font.FontColor = XLColor.FromHtml(SUCCESS_COLOR);
        row++;
        ws.Cell(row, 1).Value = "Active (Revit):";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 2).Value = sessions.Count(s => s.RevitSessionCount > 0);
        ws.Cell(row, 2).Style.Font.FontColor = XLColor.FromHtml(SUCCESS_COLOR);
        ws.Cell(row, 2).Style.Font.Bold = true;
        
        row += 2;
        ws.Cell(row, 1).Value = "USER STATUS";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
        ws.Range(row, 1, row, 8).Merge();
        
        row++;
        var headers = new[] { "User Name", "Autodesk Email", "Status", "ID Status", "Machine ID", "Revit", "Current Project", "Last Activity" };
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(row, i + 1);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
        
        row++;
        foreach (var session in sessions.OrderBy(s => s.PersonName))
        {
            ws.Cell(row, 1).Value = session.PersonName;
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 2).Value = session.AutodeskEmail;
            ws.Cell(row, 3).Value = session.IsLoggedIn ? "Online" : "Offline";
            ws.Cell(row, 3).Style.Font.FontColor = session.IsLoggedIn 
                ? XLColor.FromHtml(SUCCESS_COLOR) 
                : XLColor.Gray;
            ws.Cell(row, 4).Value = session.IdStatus;
            ws.Cell(row, 5).Value = session.MachineId;
            ws.Cell(row, 6).Value = session.RevitSessionCount;
            ws.Cell(row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            if (session.RevitSessionCount > 0)
            {
                ws.Cell(row, 6).Style.Font.Bold = true;
                ws.Cell(row, 6).Style.Font.FontColor = XLColor.FromHtml(SUCCESS_COLOR);
            }
            ws.Cell(row, 7).Value = session.CurrentProject;
            ws.Cell(row, 7).Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
            ws.Cell(row, 8).Value = session.LastActivity;
            ws.Cell(row, 8).Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
            
            ws.Range(row, 1, row, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            row++;
        }
        
        ws.Columns().AdjustToContents();
        ws.Column(7).Width = 35;
        
        workbook.SaveAs(filePath);
        return filePath;
    }
    
    /// <summary>
    /// Export a specific user's detailed activity for a date
    /// Includes hourly breakdown with all activity types
    /// </summary>
    public string ExportUserActivity(string userName, string userId, DateTime date, 
        DailyActivityBreakdown? breakdown, DailyWorkRecord? record, string filePath)
    {
        filePath = Path.ChangeExtension(filePath, ".xlsx");
        
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("User Activity");
        
        // Header
        ws.Cell("A1").Value = "TANGENT LANDSCAPE ARCHITECTURE";
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 16;
        ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Range("A1:H1").Merge();
        
        ws.Cell("A2").Value = $"User Activity Report - Generated: {DateTime.Now:yyyy-MM-dd HH:mm}";
        ws.Cell("A2").Style.Font.Italic = true;
        ws.Range("A2:H2").Merge();
        
        // User info section
        int row = 4;
        ws.Cell(row, 1).Value = "User Name:";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 2).Value = userName;
        row++;
        
        ws.Cell(row, 1).Value = "User ID:";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 2).Value = userId;
        row++;
        
        ws.Cell(row, 1).Value = "Date:";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 2).Value = date;
        ws.Cell(row, 2).Style.DateFormat.Format = "yyyy-MM-dd";
        row++;
        
        if (record != null)
        {
            ws.Cell(row, 1).Value = "Current Project:";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 2).Value = record.ProjectsWorked;
            row++;
        }
        
        row++; // Empty row
        
        // Summary section
        ws.Cell(row, 1).Value = "ACTIVITY SUMMARY";
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Font.FontSize = 12;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
        ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
        ws.Range(row, 1, row, 4).Merge();
        row++;
        
        if (breakdown != null)
        {
            var summaryHeaders = new[] { "Activity Type", "Total Minutes", "Total Hours", "Percentage" };
            for (int i = 0; i < summaryHeaders.Length; i++)
            {
                ws.Cell(row, i + 1).Value = summaryHeaders[i];
                ws.Cell(row, i + 1).Style.Font.Bold = true;
                ws.Cell(row, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
            }
            row++;
            
            var totalMinutes = breakdown.RevitMinutes + breakdown.MeetingMinutes + 
                              breakdown.TotalIdleMinutes + breakdown.OtherMinutes;
            
            // Revit
            ws.Cell(row, 1).Value = "Revit (Active Work)";
            ws.Cell(row, 2).Value = Math.Round(breakdown.RevitMinutes, 1);
            ws.Cell(row, 3).Value = Math.Round(breakdown.RevitMinutes / 60.0, 2);
            ws.Cell(row, 4).Value = totalMinutes > 0 ? Math.Round(breakdown.RevitMinutes / totalMinutes * 100, 1) : 0;
            ws.Cell(row, 4).Style.NumberFormat.Format = "0.0\"%\"";
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
            row++;
            
            // Meetings
            ws.Cell(row, 1).Value = "Meetings (Teams/Zoom)";
            ws.Cell(row, 2).Value = Math.Round(breakdown.MeetingMinutes, 1);
            ws.Cell(row, 3).Value = Math.Round(breakdown.MeetingMinutes / 60.0, 2);
            ws.Cell(row, 4).Value = totalMinutes > 0 ? Math.Round(breakdown.MeetingMinutes / totalMinutes * 100, 1) : 0;
            ws.Cell(row, 4).Style.NumberFormat.Format = "0.0\"%\"";
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.Purple;
            row++;
            
            // Idle
            ws.Cell(row, 1).Value = "Idle Time";
            ws.Cell(row, 2).Value = Math.Round(breakdown.TotalIdleMinutes, 1);
            ws.Cell(row, 3).Value = Math.Round(breakdown.TotalIdleMinutes / 60.0, 2);
            ws.Cell(row, 4).Value = totalMinutes > 0 ? Math.Round(breakdown.TotalIdleMinutes / totalMinutes * 100, 1) : 0;
            ws.Cell(row, 4).Style.NumberFormat.Format = "0.0\"%\"";
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.Gray;
            row++;
            
            // Other
            ws.Cell(row, 1).Value = "Other Applications";
            ws.Cell(row, 2).Value = Math.Round(breakdown.OtherMinutes, 1);
            ws.Cell(row, 3).Value = Math.Round(breakdown.OtherMinutes / 60.0, 2);
            ws.Cell(row, 4).Value = totalMinutes > 0 ? Math.Round(breakdown.OtherMinutes / totalMinutes * 100, 1) : 0;
            ws.Cell(row, 4).Style.NumberFormat.Format = "0.0\"%\"";
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.Orange;
            row++;
            
            // Total
            ws.Cell(row, 1).Value = "TOTAL";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 2).Value = Math.Round(totalMinutes, 1);
            ws.Cell(row, 2).Style.Font.Bold = true;
            ws.Cell(row, 3).Value = Math.Round(totalMinutes / 60.0, 2);
            ws.Cell(row, 3).Style.Font.Bold = true;
            ws.Cell(row, 4).Value = 100;
            ws.Cell(row, 4).Style.NumberFormat.Format = "0\"%\"";
            ws.Cell(row, 4).Style.Font.Bold = true;
            row += 2;
            
            // Hourly breakdown section
            ws.Cell(row, 1).Value = "HOURLY ACTIVITY BREAKDOWN";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 1).Style.Font.FontSize = 12;
            ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(TANGENT_CYAN);
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
            ws.Range(row, 1, row, 6).Merge();
            row++;
            
            var hourlyHeaders = new[] { "Hour", "Revit (min)", "Meeting (min)", "Idle (min)", "Other (min)", "Total (min)" };
            for (int i = 0; i < hourlyHeaders.Length; i++)
            {
                ws.Cell(row, i + 1).Value = hourlyHeaders[i];
                ws.Cell(row, i + 1).Style.Font.Bold = true;
                ws.Cell(row, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml(HEADER_BG);
                ws.Cell(row, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            row++;
            
            foreach (var hourly in breakdown.HourlyBreakdown.OrderBy(h => h.Hour))
            {
                var hourLabel = hourly.Hour == 0 ? "12 AM" : 
                               hourly.Hour < 12 ? $"{hourly.Hour} AM" : 
                               hourly.Hour == 12 ? "12 PM" : 
                               $"{hourly.Hour - 12} PM";
                
                ws.Cell(row, 1).Value = hourLabel;
                ws.Cell(row, 2).Value = Math.Round(hourly.RevitMinutes, 1);
                ws.Cell(row, 3).Value = Math.Round(hourly.MeetingMinutes, 1);
                ws.Cell(row, 4).Value = Math.Round(hourly.IdleMinutes, 1);
                ws.Cell(row, 5).Value = Math.Round(hourly.OtherMinutes, 1);
                ws.Cell(row, 6).Value = Math.Round(hourly.RevitMinutes + hourly.MeetingMinutes + 
                                                    hourly.IdleMinutes + hourly.OtherMinutes, 1);
                
                // Highlight work hours
                if (hourly.Hour >= 8 && hourly.Hour < 18)
                {
                    ws.Range(row, 1, row, 6).Style.Fill.BackgroundColor = XLColor.FromArgb(245, 245, 245);
                }
                
                // Color code activity cells
                if (hourly.RevitMinutes > 0)
                    ws.Cell(row, 2).Style.Font.FontColor = XLColor.FromHtml(TANGENT_CYAN);
                if (hourly.MeetingMinutes > 0)
                    ws.Cell(row, 3).Style.Font.FontColor = XLColor.Purple;
                
                row++;
            }
        }
        
        ws.Columns().AdjustToContents();
        workbook.SaveAs(filePath);
        return filePath;
    }
}
