using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TeamsNextMeetingTile.Models;

namespace TeamsNextMeetingTile.Services;

public static class IcsCalendarParser
{
    private static readonly HttpClient Http = new()
    {
        Timeout = TimeSpan.FromSeconds(20)
    };

    private static readonly Regex TeamsUrlRegex = new(
        @"https://(?:teams\.microsoft\.com/(?:l/meetup-join|meet)/[^\s""<>\\]+|teams\.live\.com/meet/[^\s""<>\\]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static async Task<List<MeetingItem>> FetchAndParseAsync(string icsUrl, DateTime windowStart, DateTime windowEnd)
    {
        if (string.IsNullOrWhiteSpace(icsUrl))
            return new List<MeetingItem>();

        string normalizedUrl = icsUrl.Trim();
        if (normalizedUrl.StartsWith("webcal://", StringComparison.OrdinalIgnoreCase))
        {
            normalizedUrl = "https://" + normalizedUrl.Substring("webcal://".Length);
        }

        string rawIcs = await Http.GetStringAsync(normalizedUrl).ConfigureAwait(false);
        return ParseIcsContent(rawIcs, windowStart, windowEnd);
    }

    public static List<MeetingItem> ParseIcsContent(string rawIcs, DateTime windowStart, DateTime windowEnd)
    {
        var results = new List<MeetingItem>();
        if (string.IsNullOrWhiteSpace(rawIcs))
            return results;

        // Unfold RFC 5545 folded lines
        var unfoldedLines = new List<string>();
        string[] rawLines = rawIcs.Replace("\r\n", "\n").Split('\n');
        foreach (string line in rawLines)
        {
            if ((line.StartsWith(" ") || line.StartsWith("\t")) && unfoldedLines.Count > 0)
            {
                unfoldedLines[^1] += line.Substring(1);
            }
            else
            {
                unfoldedLines.Add(line);
            }
        }

        var rawEvents = new List<RawVEvent>();
        RawVEvent? current = null;

        foreach (string line in unfoldedLines)
        {
            string trimmed = line.Trim();
            if (trimmed.Equals("BEGIN:VEVENT", StringComparison.OrdinalIgnoreCase))
            {
                current = new RawVEvent();
                continue;
            }
            if (trimmed.Equals("END:VEVENT", StringComparison.OrdinalIgnoreCase))
            {
                if (current != null)
                {
                    rawEvents.Add(current);
                    current = null;
                }
                continue;
            }
            if (current == null)
                continue;

            int colonIdx = line.IndexOf(':');
            if (colonIdx <= 0)
                continue;

            string keyPart = line.Substring(0, colonIdx);
            string valPart = line.Substring(colonIdx + 1);

            string[] keySegments = keyPart.Split(';');
            string propName = keySegments[0].Trim().ToUpperInvariant();
            string? tzId = null;
            bool isDateOnly = false;

            for (int i = 1; i < keySegments.Length; i++)
            {
                string seg = keySegments[i].Trim();
                if (seg.StartsWith("TZID=", StringComparison.OrdinalIgnoreCase))
                {
                    tzId = seg.Substring(5).Trim('"');
                }
                else if (seg.Equals("VALUE=DATE", StringComparison.OrdinalIgnoreCase))
                {
                    isDateOnly = true;
                }
            }

            switch (propName)
            {
                case "UID":
                    current.Uid = valPart.Trim();
                    break;
                case "SUMMARY":
                    current.Summary = UnescapeIcsText(valPart);
                    break;
                case "LOCATION":
                    current.Location = UnescapeIcsText(valPart);
                    break;
                case "DESCRIPTION":
                    current.Description = UnescapeIcsText(valPart);
                    break;
                case "STATUS":
                    current.Status = valPart.Trim();
                    break;
                case "TRANSP":
                    current.Transp = valPart.Trim();
                    break;
                case "X-MICROSOFT-CDO-BUSYSTATUS":
                    current.BusyStatus = valPart.Trim();
                    break;
                case "DTSTART":
                    current.IsAllDay = isDateOnly || valPart.Trim().Length == 8;
                    current.StartTzId = tzId;
                    if (TryParseRawIcsWallClock(valPart.Trim(), out DateTime rawStart, out bool startUtc))
                    {
                        current.DtStartWallClock = rawStart;
                        current.IsStartUtc = startUtc;
                        current.DtStartLocal = ConvertWallClockToLocal(rawStart, tzId, startUtc);
                    }
                    break;
                case "DTEND":
                    current.EndTzId = tzId;
                    if (TryParseRawIcsWallClock(valPart.Trim(), out DateTime rawEnd, out bool endUtc))
                    {
                        current.DtEndWallClock = rawEnd;
                        current.IsEndUtc = endUtc;
                        current.DtEndLocal = ConvertWallClockToLocal(rawEnd, tzId, endUtc);
                    }
                    break;
                case "RECURRENCE-ID":
                    if (TryParseRawIcsWallClock(valPart.Trim(), out DateTime rawRec, out bool recUtc))
                    {
                        current.RecurrenceIdWallClock = rawRec;
                        current.RecurrenceIdLocal = ConvertWallClockToLocal(rawRec, tzId, recUtc);
                    }
                    break;
                case "RRULE":
                    current.RRule = valPart.Trim();
                    break;
                case "EXDATE":
                    foreach (string exPart in valPart.Split(',', StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (TryParseRawIcsWallClock(exPart.Trim(), out DateTime rawEx, out bool exUtc))
                        {
                            current.ExDatesLocal.Add(ConvertWallClockToLocal(rawEx, tzId, exUtc));
                        }
                    }
                    break;
            }
        }

        // Collect overridden recurrence instances by UID + Local time AND UID + WallClock date
        var overriddenLocalKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var overriddenWallKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var ev in rawEvents)
        {
            if (!string.IsNullOrEmpty(ev.Uid))
            {
                if (ev.RecurrenceIdLocal.HasValue)
                    overriddenLocalKeys.Add($"{ev.Uid}|{ev.RecurrenceIdLocal.Value:yyyyMMddHHmm}");
                if (ev.RecurrenceIdWallClock.HasValue)
                    overriddenWallKeys.Add($"{ev.Uid}|{ev.RecurrenceIdWallClock.Value:yyyyMMddHHmm}");
            }
        }

        foreach (var ev in rawEvents)
        {
            if (ev.DtStartWallClock == default)
                continue;

            if (ev.Status.Equals("CANCELLED", StringComparison.OrdinalIgnoreCase) ||
                ev.BusyStatus.Equals("FREE", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (ev.IsAllDay)
                continue;

            TimeSpan duration = ev.DtEndWallClock != default && ev.DtEndWallClock > ev.DtStartWallClock
                ? ev.DtEndWallClock - ev.DtStartWallClock
                : TimeSpan.FromMinutes(30);

            string joinUrl = ExtractTeamsJoinUrl(ev.Location, ev.Description);
            string title = string.IsNullOrWhiteSpace(ev.Summary) ? "Busy / Meeting" : ev.Summary;

            if (string.IsNullOrWhiteSpace(ev.RRule) || ev.RecurrenceIdLocal.HasValue)
            {
                DateTime startLocal = ev.DtStartLocal;
                DateTime endLocal = ev.DtEndLocal != default && ev.DtEndLocal > startLocal
                    ? ev.DtEndLocal
                    : startLocal + duration;

                if (endLocal >= windowStart && startLocal <= windowEnd)
                {
                    results.Add(new MeetingItem
                    {
                        Id = $"{ev.Uid}_{startLocal:O}",
                        Title = title,
                        StartTime = startLocal,
                        EndTime = endLocal,
                        Location = ev.Location,
                        JoinUrl = joinUrl,
                        Source = "ICS"
                    });
                }
            }
            else
            {
                var occurrences = ExpandRRule(ev, duration, windowStart, windowEnd, overriddenLocalKeys, overriddenWallKeys);
                foreach (var occStartLocal in occurrences)
                {
                    results.Add(new MeetingItem
                    {
                        Id = $"{ev.Uid}_{occStartLocal:O}",
                        Title = title,
                        StartTime = occStartLocal,
                        EndTime = occStartLocal + duration,
                        Location = ev.Location,
                        JoinUrl = joinUrl,
                        Source = "ICS"
                    });
                }
            }
        }

        return results
            .OrderBy(m => m.StartTime)
            .ToList();
    }

    private static List<DateTime> ExpandRRule(
        RawVEvent ev,
        TimeSpan duration,
        DateTime windowStart,
        DateTime windowEnd,
        HashSet<string> overriddenLocalKeys,
        HashSet<string> overriddenWallKeys)
    {
        var list = new List<DateTime>();
        var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string seg in ev.RRule.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            int eq = seg.IndexOf('=');
            if (eq > 0)
            {
                parts[seg.Substring(0, eq).Trim()] = seg.Substring(eq + 1).Trim();
            }
        }

        if (!parts.TryGetValue("FREQ", out string? freq))
            return list;

        int interval = 1;
        if (parts.TryGetValue("INTERVAL", out string? intStr) && int.TryParse(intStr, out int parsedInt) && parsedInt > 0)
            interval = parsedInt;

        int? maxCount = null;
        if (parts.TryGetValue("COUNT", out string? cntStr) && int.TryParse(cntStr, out int parsedCnt) && parsedCnt > 0)
            maxCount = parsedCnt;

        DateTime? untilLocal = null;
        if (parts.TryGetValue("UNTIL", out string? untilStr) &&
            TryParseRawIcsWallClock(untilStr, out DateTime rawUntil, out bool untilUtc))
        {
            untilLocal = ConvertWallClockToLocal(rawUntil, ev.StartTzId, untilUtc);
        }

        HashSet<DayOfWeek>? byDays = null;
        if (parts.TryGetValue("BYDAY", out string? byDayStr))
        {
            byDays = new HashSet<DayOfWeek>();
            foreach (string d in byDayStr.Split(',', StringSplitOptions.RemoveEmptyEntries))
            {
                string code = d.Length >= 2 ? d.Substring(d.Length - 2).ToUpperInvariant() : d.ToUpperInvariant();
                switch (code)
                {
                    case "MO": byDays.Add(DayOfWeek.Monday); break;
                    case "TU": byDays.Add(DayOfWeek.Tuesday); break;
                    case "WE": byDays.Add(DayOfWeek.Wednesday); break;
                    case "TH": byDays.Add(DayOfWeek.Thursday); break;
                    case "FR": byDays.Add(DayOfWeek.Friday); break;
                    case "SA": byDays.Add(DayOfWeek.Saturday); break;
                    case "SU": byDays.Add(DayOfWeek.Sunday); break;
                }
            }
        }

        int generatedCount = 0;
        // Expand in the event's own wall-clock timezone so DST transitions on future dates are accurately converted!
        DateTime seedWall = ev.DtStartWallClock;
        int maxDaysSearch = 800;

        for (int dayOffset = 0; dayOffset <= maxDaysSearch; dayOffset++)
        {
            DateTime candidateWall = seedWall.Date.AddDays(dayOffset).Add(seedWall.TimeOfDay);
            if (candidateWall < seedWall)
                continue;

            bool matchesFreq = false;
            if (freq.Equals("DAILY", StringComparison.OrdinalIgnoreCase))
            {
                int daysSinceSeed = (int)(candidateWall.Date - seedWall.Date).TotalDays;
                if (daysSinceSeed % interval == 0)
                {
                    matchesFreq = byDays == null || byDays.Contains(candidateWall.DayOfWeek);
                }
            }
            else if (freq.Equals("WEEKLY", StringComparison.OrdinalIgnoreCase))
            {
                int daysSinceSeed = (int)(candidateWall.Date - seedWall.Date).TotalDays;
                int weeksSinceSeed = daysSinceSeed / 7;
                if (weeksSinceSeed % interval == 0)
                {
                    if (byDays != null && byDays.Count > 0)
                        matchesFreq = byDays.Contains(candidateWall.DayOfWeek);
                    else
                        matchesFreq = candidateWall.DayOfWeek == seedWall.DayOfWeek;
                }
            }
            else if (freq.Equals("MONTHLY", StringComparison.OrdinalIgnoreCase))
            {
                int monthsSinceSeed = (candidateWall.Year - seedWall.Year) * 12 + (candidateWall.Month - seedWall.Month);
                if (monthsSinceSeed >= 0 && monthsSinceSeed % interval == 0 && candidateWall.Day == seedWall.Day)
                {
                    matchesFreq = true;
                }
            }

            if (!matchesFreq)
                continue;

            // Convert this specific occurrence date/time from the event's TZID to the user's local timezone
            DateTime candidateLocal = ConvertWallClockToLocal(candidateWall, ev.StartTzId, ev.IsStartUtc);

            if (untilLocal.HasValue && candidateLocal > untilLocal.Value)
                break;

            if (candidateLocal > windowEnd)
                break;

            generatedCount++;
            if (maxCount.HasValue && generatedCount > maxCount.Value)
                break;

            bool isExcluded = ev.ExDatesLocal.Any(ex => Math.Abs((ex - candidateLocal).TotalMinutes) < 5);
            if (isExcluded)
                continue;

            if (!string.IsNullOrEmpty(ev.Uid))
            {
                if (overriddenLocalKeys.Contains($"{ev.Uid}|{candidateLocal:yyyyMMddHHmm}") ||
                    overriddenWallKeys.Contains($"{ev.Uid}|{candidateWall:yyyyMMddHHmm}"))
                {
                    continue;
                }
            }

            if (candidateLocal + duration >= windowStart && candidateLocal <= windowEnd)
            {
                list.Add(candidateLocal);
            }
        }

        return list;
    }

    public static string ExtractTeamsJoinUrl(params string?[] fields)
    {
        foreach (string? field in fields)
        {
            if (string.IsNullOrWhiteSpace(field))
                continue;

            Match m = TeamsUrlRegex.Match(field);
            if (m.Success)
            {
                return m.Value.TrimEnd('>', ')', '.', ',', ';', '"', '\'');
            }
        }
        return string.Empty;
    }

    private static bool TryParseRawIcsWallClock(string raw, out DateTime wallClock, out bool isUtc)
    {
        wallClock = default;
        isUtc = false;
        if (string.IsNullOrWhiteSpace(raw))
            return false;

        isUtc = raw.EndsWith("Z", StringComparison.OrdinalIgnoreCase);
        string clean = raw.TrimEnd('Z', 'z');

        string[] formats =
        {
            "yyyyMMdd'T'HHmmss",
            "yyyyMMdd'T'HHmm",
            "yyyyMMdd"
        };

        if (!DateTime.TryParseExact(clean, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsed))
            return false;

        wallClock = DateTime.SpecifyKind(parsed, DateTimeKind.Unspecified);
        return true;
    }

    public static DateTime ConvertWallClockToLocal(DateTime wallClock, string? tzId, bool isUtc)
    {
        if (isUtc)
        {
            return DateTime.SpecifyKind(wallClock, DateTimeKind.Utc).ToLocalTime();
        }

        if (!string.IsNullOrWhiteSpace(tzId))
        {
            try
            {
                TimeZoneInfo tzi = TimeZoneInfo.FindSystemTimeZoneById(tzId.Trim());
                DateTime uns = DateTime.SpecifyKind(wallClock, DateTimeKind.Unspecified);
                if (tzi.IsInvalidTime(uns))
                {
                    uns = uns.AddHours(1);
                }
                return TimeZoneInfo.ConvertTime(uns, tzi, TimeZoneInfo.Local);
            }
            catch
            {
                // Fallback if unknown TZID
            }
        }

        return DateTime.SpecifyKind(wallClock, DateTimeKind.Local);
    }

    private static string UnescapeIcsText(string input)
    {
        if (string.IsNullOrEmpty(input))
            return string.Empty;

        return input
            .Replace("\\n", "\n")
            .Replace("\\N", "\n")
            .Replace("\\,", ",")
            .Replace("\\;", ";")
            .Replace("\\\\", "\\")
            .Trim();
    }

    private class RawVEvent
    {
        public string Uid { get; set; } = string.Empty;
        public string Summary { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Transp { get; set; } = string.Empty;
        public string BusyStatus { get; set; } = string.Empty;
        public bool IsAllDay { get; set; }
        public string? StartTzId { get; set; }
        public string? EndTzId { get; set; }
        public bool IsStartUtc { get; set; }
        public bool IsEndUtc { get; set; }
        public DateTime DtStartWallClock { get; set; }
        public DateTime DtEndWallClock { get; set; }
        public DateTime DtStartLocal { get; set; }
        public DateTime DtEndLocal { get; set; }
        public DateTime? RecurrenceIdWallClock { get; set; }
        public DateTime? RecurrenceIdLocal { get; set; }
        public string RRule { get; set; } = string.Empty;
        public List<DateTime> ExDatesLocal { get; } = new();
    }
}
