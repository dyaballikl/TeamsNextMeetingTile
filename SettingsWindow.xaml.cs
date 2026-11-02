using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using TeamsNextMeetingTile.Models;

namespace TeamsNextMeetingTile;

public partial class SettingsWindow : Window
{
    private readonly MainWindow _mainWindow;
    private readonly AppSettings _settings;
    private bool _isLoadingUi = true;

    public IntPtr Hwnd => new WindowInteropHelper(this).Handle;

    public SettingsWindow(MainWindow mainWindow, AppSettings settings)
    {
        InitializeComponent();
        _mainWindow = mainWindow;
        _settings = settings;

        Title = $"Teams Next Meeting Tile ({AppVersionInfo.VersionString}) — Settings & Calendar Sync";
        TxtVersionFooter.Text = AppVersionInfo.FullDisplay;

        LoadSettingsIntoUi();
        TxtNewStart.Text = DateTime.Now.AddMinutes(45).ToString("yyyy-MM-dd HH:mm");
        RefreshMeetingsList();
    }

    private void LoadSettingsIntoUi()
    {
        _isLoadingUi = true;
        ChkDockToTeams.IsChecked = _settings.DockToTeamsWindow;
        ChkAutoHideCovered.IsChecked = _settings.AutoHideWhenTeamsCovered;
        ChkLockPosition.IsChecked = _settings.LockPosition;
        TxtTileWidth.Text = ((int)_settings.TileWidth).ToString(CultureInfo.InvariantCulture);

        ChkEnableM365.IsChecked = _settings.EnableM365WebViewSync;
        TxtIcsUrl.Text = _settings.IcsFeedUrl ?? string.Empty;
        ChkEnableClassicCom.IsChecked = _settings.EnableClassicOutlookCom;
        TxtM365Status.Text = $"Status: {_mainWindow.GetM365StatusText()}";
        _isLoadingUi = false;
    }

    public void RefreshMeetingsList()
    {
        TxtM365Status.Text = $"Status: {_mainWindow.GetM365StatusText()}";
        DateTime now = DateTime.Now;
        var all = _mainWindow.GetAllUpcomingMeetings();

        TxtMeetingCount.Text = $"{all.Count} upcoming meeting(s)";

        var rows = all.Select(m => new MeetingRowViewModel
        {
            Id = m.Id,
            Title = m.Title,
            StartDisplay = m.GetFormattedStartTime(now),
            DhmDisplay = m.GetDhmCountdown(now),
            Source = m.Source,
            DeleteVisibility = m.Source == "Manual" ? Visibility.Visible : Visibility.Collapsed
        }).ToList();

        LstMeetings.ItemsSource = rows;
    }

    private void Setting_Changed(object sender, RoutedEventArgs e)
    {
        if (_isLoadingUi)
            return;

        _settings.DockToTeamsWindow = ChkDockToTeams.IsChecked == true;
        _settings.AutoHideWhenTeamsCovered = ChkAutoHideCovered.IsChecked == true;
        _settings.LockPosition = ChkLockPosition.IsChecked == true;
        _settings.EnableM365WebViewSync = ChkEnableM365.IsChecked == true;
        _settings.EnableClassicOutlookCom = ChkEnableClassicCom.IsChecked == true;
        _settings.Save();

        _mainWindow.ApplySettingsImmediately();
    }

    private void TxtTileWidth_LostFocus(object sender, RoutedEventArgs e)
    {
        if (double.TryParse(TxtTileWidth.Text, out double w))
        {
            _settings.TileWidth = Math.Clamp(w, 175, 420);
            TxtTileWidth.Text = ((int)_settings.TileWidth).ToString(CultureInfo.InvariantCulture);
            _settings.Save();
            _mainWindow.ApplySettingsImmediately();
        }
    }

    private void BtnPresetQuickViews_Click(object sender, RoutedEventArgs e)
    {
        // Docks right above 'Quick views' inside the Teams Chat sidebar
        _settings.OffsetX = 52;
        _settings.OffsetY = 122;
        _settings.TileWidth = 230;
        TxtTileWidth.Text = "230";
        _settings.Save();
        _mainWindow.ApplySettingsImmediately();
    }

    private void BtnPresetChatHeader_Click(object sender, RoutedEventArgs e)
    {
        // Docks in the top area of the Chat sidebar header
        _settings.OffsetX = 52;
        _settings.OffsetY = 46;
        _settings.TileWidth = 230;
        TxtTileWidth.Text = "230";
        _settings.Save();
        _mainWindow.ApplySettingsImmediately();
    }

    private void BtnPresetTitleBar_Click(object sender, RoutedEventArgs e)
    {
        // Docks in the empty space of the Teams top window title bar left of the search box
        _settings.OffsetX = 88;
        _settings.OffsetY = 5;
        _settings.TileWidth = 225;
        TxtTileWidth.Text = "225";
        _settings.Save();
        _mainWindow.ApplySettingsImmediately();
    }

    private void BtnOpenM365Window_Click(object sender, RoutedEventArgs e)
    {
        _mainWindow.ShowM365CalendarSyncWindow();
    }

    private async void BtnSyncNow_Click(object sender, RoutedEventArgs e)
    {
        TxtM365Status.Text = "Status: Syncing calendars...";
        await _mainWindow.SyncAllCalendarsNowAsync();
        RefreshMeetingsList();
    }

    private async void BtnSaveIcs_Click(object sender, RoutedEventArgs e)
    {
        _settings.IcsFeedUrl = TxtIcsUrl.Text.Trim();
        _settings.Save();
        TxtM365Status.Text = "Status: Fetching ICS feed...";
        await _mainWindow.SyncAllCalendarsNowAsync();
        RefreshMeetingsList();
    }

    private void BtnAddManual_Click(object sender, RoutedEventArgs e)
    {
        string title = TxtNewTitle.Text.Trim();
        if (string.IsNullOrWhiteSpace(title))
            title = "Pinned Meeting";

        if (!DateTime.TryParse(TxtNewStart.Text.Trim(), out DateTime startDt))
        {
            MessageBox.Show("Please enter a valid start date/time (e.g. 2026-10-30 14:30).", "Invalid Time", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _settings.ManualMeetings.Add(new MeetingItem
        {
            Id = "MAN_" + Guid.NewGuid().ToString("N"),
            Title = title,
            StartTime = startDt,
            EndTime = startDt.AddMinutes(30),
            JoinUrl = TxtNewJoinUrl.Text.Trim(),
            Source = "Manual"
        });
        _settings.Save();
        _mainWindow.RebuildMergedMeetings();
        RefreshMeetingsList();
    }

    private void BtnDeleteManual_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string id)
        {
            _settings.ManualMeetings.RemoveAll(m => m.Id == id);
            _settings.Save();
            _mainWindow.RebuildMergedMeetings();
            RefreshMeetingsList();
        }
    }

    private void BtnQuitApp_Click(object sender, RoutedEventArgs e)
    {
        _mainWindow.ShutdownApplication();
    }

    private void BtnClose_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private class MeetingRowViewModel
    {
        public string Id { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string StartDisplay { get; set; } = string.Empty;
        public string DhmDisplay { get; set; } = string.Empty;
        public string Source { get; set; } = string.Empty;
        public Visibility DeleteVisibility { get; set; }
    }
}
