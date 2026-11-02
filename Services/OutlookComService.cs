using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using TeamsNextMeetingTile.Models;

namespace TeamsNextMeetingTile.Services;

public static class OutlookComService
{
    public static List<MeetingItem> TryReadUpcomingMeetings(DateTime windowStart, DateTime windowEnd, bool requireOutlookRunning = true)
    {
        var results = new List<MeetingItem>();

        try
        {
            if (requireOutlookRunning && Process.GetProcessesByName("OUTLOOK").Length == 0)
            {
                return results;
            }

            Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType == null)
                return results;

            dynamic? app = Activator.CreateInstance(outlookType);
            if (app == null)
                return results;

            dynamic ns = app.GetNamespace("MAPI");
            // 9 = olFolderCalendar
            dynamic calendarFolder = ns.GetDefaultFolder(9);
            dynamic items = calendarFolder.Items;
            items.IncludeRecurrences = true;
            items.Sort("[Start]");

            string filter = $"[End] >= '{windowStart.ToString("g", CultureInfo.CurrentCulture)}' AND [Start] <= '{windowEnd.ToString("g", CultureInfo.CurrentCulture)}'";
            dynamic restricted = items.Restrict(filter);

            int safetyLimit = 60;
            int count = 0;

            foreach (dynamic item in restricted)
            {
                if (++count > safetyLimit)
                    break;

                try
                {
                    bool isAllDay = false;
                    try { isAllDay = (bool)item.AllDayEvent; } catch { }
                    if (isAllDay)
                        continue;

                    // 0 = olFree
                    int busyStatus = 2;
                    try { busyStatus = (int)item.BusyStatus; } catch { }
                    if (busyStatus == 0)
                        continue;

                    DateTime start = (DateTime)item.Start;
                    DateTime end = (DateTime)item.End;
                    if (end < windowStart || start > windowEnd)
                        continue;

                    string subject = (string)(item.Subject ?? "Meeting");
                    string location = string.Empty;
                    string body = string.Empty;
                    try { location = (string)(item.Location ?? string.Empty); } catch { }
                    try { body = (string)(item.Body ?? string.Empty); } catch { }

                    string joinUrl = IcsCalendarParser.ExtractTeamsJoinUrl(location, body);

                    results.Add(new MeetingItem
                    {
                        Id = $"COM_{subject}_{start:O}",
                        Title = subject.Trim(),
                        StartTime = start,
                        EndTime = end,
                        Location = location.Trim(),
                        JoinUrl = joinUrl,
                        Source = "Outlook"
                    });
                }
                catch
                {
                    // Skip unreadable item
                }
            }
        }
        catch
        {
            // Ignore COM errors when Classic Outlook is not active
        }

        return results;
    }
}
