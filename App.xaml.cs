using System;
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

        _singleInstanceMutex = new Mutex(true, "TeamsNextMeetingTile_SingleInstance_Mutex", out bool createdNew);
        if (!createdNew)
        {
            System.Windows.MessageBox.Show("Teams Next Meeting Tile is already running (check your Teams sidebar or System Tray).",
                "Teams Next Meeting Tile", MessageBoxButton.OK, MessageBoxImage.Information);
            Shutdown();
            return;
        }

        DispatcherUnhandledException += (_, args) =>
        {
            args.Handled = true;
        };

        base.OnStartup(e);
    }
}
