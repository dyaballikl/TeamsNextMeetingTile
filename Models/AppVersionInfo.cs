using System;
using System.Linq;
using System.Reflection;

namespace TeamsNextMeetingTile.Models;

public static class AppVersionInfo
{
    public static string VersionString
    {
        get
        {
            var asm = Assembly.GetExecutingAssembly();
            string ver = asm.GetName().Version?.ToString(3) ?? "1.2.0";
            return $"v{ver}";
        }
    }

    public static string PublishDateString
    {
        get
        {
            var asm = Assembly.GetExecutingAssembly();
            var meta = asm.GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "PublishTimestamp");
            if (!string.IsNullOrWhiteSpace(meta?.Value))
            {
                return meta.Value;
            }
            return "2026-11-02 AEST";
        }
    }

    public static string FullDisplay => $"{VersionString}  •  Published: {PublishDateString}";
}
