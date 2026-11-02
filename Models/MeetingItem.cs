using System;
using System.Globalization;

namespace TeamsNextMeetingTile.Models;

public class MeetingItem
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Title { get; set; } = string.Empty;
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
    public string Location { get; set; } = string.Empty;
    public string JoinUrl { get; set; } = string.Empty;
    public string Source { get; set; } = "M365";
    public bool IsAllDay { get; set; }

    public bool HasTeamsLink =>
        !string.IsNullOrWhiteSpace(JoinUrl) &&
        (JoinUrl.Contains("teams.microsoft.com", StringComparison.OrdinalIgnoreCase) ||
         JoinUrl.Contains("teams.live.com", StringComparison.OrdinalIgnoreCase) ||
         JoinUrl.StartsWith("msteams:", StringComparison.OrdinalIgnoreCase) ||
         JoinUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase));

    public string GetFormattedStartTime(DateTime now)
    {
        string timePart = StartTime.ToString("h:mm tt", CultureInfo.InvariantCulture);
        if (StartTime.Date == now.Date)
        {
            return timePart;
        }
        if (StartTime.Date == now.Date.AddDays(1))
        {
            return $"Tmrw {timePart}";
        }
        if ((StartTime.Date - now.Date).TotalDays < 7 && StartTime.Date > now.Date)
        {
            return $"{StartTime.ToString("ddd", CultureInfo.InvariantCulture)} {timePart}";
        }
        return $"{StartTime.ToString("dd MMM", CultureInfo.InvariantCulture)} {timePart}";
    }

    public string GetFormattedTimeRange(DateTime now)
    {
        string start = GetFormattedStartTime(now);
        string end = EndTime.ToString("h:mm tt", CultureInfo.InvariantCulture);
        return $"{start} – {end}";
    }

    /// <summary>
    /// Formats the remaining time until StartTime dynamically:
    /// - Over 2d 0h 0m: "#d"
    /// - 1d to 2d 0h 0m: "#d #h #m"
    /// - Under 1d (>= 1h): "#h #m"
    /// - Under 1h: "#m"
    /// </summary>
    public string GetDhmCountdown(DateTime now)
    {
        TimeSpan diff = StartTime - now;
        if (diff.TotalSeconds <= 0)
        {
            if (now <= EndTime)
            {
                int minsIn = Math.Max(0, (int)(now - StartTime).TotalMinutes);
                return minsIn == 0 ? "Now" : $"Started {minsIn}m ago";
            }
            return "Ended";
        }

        int totalMinutes = Math.Max(1, (int)Math.Ceiling(diff.TotalSeconds / 60.0));
        int days = totalMinutes / (24 * 60);
        int hours = (totalMinutes % (24 * 60)) / 60;
        int minutes = totalMinutes % 60;

        // Over 2d 0h 0m -> just show #d
        if (totalMinutes > 2 * 24 * 60)
        {
            return $"{days}d";
        }

        // 1d up to 2d 0h 0m -> show #d #h #m
        if (days >= 1)
        {
            return $"{days}d {hours}h {minutes}m";
        }

        // Under 1d, at least 1h -> show #h #m
        if (hours >= 1)
        {
            return $"{hours}h {minutes}m";
        }

        // Under 1h -> just show #m
        return $"{minutes}m";
    }
}
