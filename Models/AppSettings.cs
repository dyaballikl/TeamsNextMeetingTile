using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace TeamsNextMeetingTile.Models;

public class AppSettings
{
    // Relative offset from the top-left of the Microsoft Teams window
    public double OffsetX { get; set; } = 52;
    public double OffsetY { get; set; } = 122;
    public double TileWidth { get; set; } = 230;

    // Fallback standalone screen coordinates when not docked to Teams
    public double StandaloneLeft { get; set; } = 100;
    public double StandaloneTop { get; set; } = 120;

    public bool DockToTeamsWindow { get; set; } = true;
    public bool AutoHideWhenTeamsCovered { get; set; } = true;
    public bool IsCollapsed { get; set; } = false;
    public bool LockPosition { get; set; } = false;
    public bool IncludeInProgressMeetings { get; set; } = false; // If false, prioritizes the next meeting that hasn't started yet (or started < 3 mins ago)
    public int GraceMinutesAfterStart { get; set; } = 3;

    // Calendar Sync Sources
    public bool EnableM365WebViewSync { get; set; } = true;
    public string M365CalendarUrl { get; set; } = "https://outlook.office.com/calendar/view/workweek";
    public string IcsFeedUrl { get; set; } = string.Empty;
    public bool EnableClassicOutlookCom { get; set; } = true;

    public List<MeetingItem> ManualMeetings { get; set; } = new();
    public List<MeetingItem> CachedMeetings { get; set; } = new();

    public static string DataFolder =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TeamsNextMeetingTile");

    public static string SettingsFilePath =>
        Path.Combine(DataFolder, "settings.json");

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(SettingsFilePath))
            {
                string json = File.ReadAllText(SettingsFilePath);
                var loaded = JsonSerializer.Deserialize<AppSettings>(json);
                if (loaded != null)
                {
                    return loaded;
                }
            }
        }
        catch
        {
            // Ignore corrupt settings and use defaults
        }
        return new AppSettings();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(DataFolder);
            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsFilePath, json);
        }
        catch
        {
            // Ignore write errors
        }
    }
}
