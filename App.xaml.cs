using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using TeamsNextMeetingTile.Models;

namespace TeamsNextMeetingTile;

public partial class App : System.Windows.Application
{
    private static Mutex? _singleInstanceMutex;

    protected override void OnStartup(StartupEventArgs e)
    {
        if (e.Args.Length > 0 && e.Args[0] == "--self-test")
        {
            var settings = AppSettings.Load();
            var list = System.Threading.Tasks.Task.Run(() =>
                Services.IcsCalendarParser.FetchAndParseAsync(
                    settings.IcsFeedUrl,
                    DateTime.Now.Date,
                    DateTime.Now.Date.AddDays(2))).GetAwaiter().GetResult();

            var lines = new System.Collections.Generic.List<string>();
            foreach (var item in list)
            {
                lines.Add($"{item.StartTime:yyyy-MM-dd h:mm tt} -> {item.EndTime:h:mm tt} | {item.Title}");
            }
            File.WriteAllLines(Path.Combine(AppSettings.DataFolder, "ics_test_output.txt"), lines);
            Environment.Exit(0);
            return;
        }

        // If a previous instance is already running, replace it cleanly so re-launching always refreshes the tile
        int currentPid = Environment.ProcessId;
        foreach (Process proc in Process.GetProcessesByName("TeamsNextMeetingTile"))
        {
            try
            {
                if (proc.Id != currentPid)
                {
                    proc.Kill();
                    proc.WaitForExit(2000);
                }
            }
            catch { }
            finally
            {
                proc.Dispose();
            }
        }

        _singleInstanceMutex = new Mutex(true, "TeamsNextMeetingTile_SingleInstance_Mutex", out _);

        DispatcherUnhandledException += (_, args) =>
        {
            args.Handled = true;
        };

        base.OnStartup(e);
    }
}
