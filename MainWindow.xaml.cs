using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using TeamsNextMeetingTile.Models;
using TeamsNextMeetingTile.Services;
using Forms = System.Windows.Forms;

namespace TeamsNextMeetingTile;

public partial class MainWindow : Window
{
    private readonly AppSettings _settings;
    private readonly DispatcherTimer _countdownTimer;
    private readonly DispatcherTimer _dockTrackingTimer;
    private readonly DispatcherTimer _calendarPollTimer;

    private M365WebViewSyncWindow? _m365SyncWindow;
    private SettingsWindow? _settingsWindow;
    private Forms.NotifyIcon? _trayIcon;

    private List<MeetingItem> _m365Meetings = new();
    private List<MeetingItem> _icsMeetings = new();
    private List<MeetingItem> _comMeetings = new();
    private List<MeetingItem> _mergedUpcoming = new();

    private int _selectedMeetingIndex = 0;
    private bool _isDragging = false;
    private WinRect _lastTeamsBounds;
    private double _lastTeamsDpiScale = 1.0;
    private bool _hasLastTeamsBounds = false;

    private string _m365StatusMessage = "Starting M365 sync...";
    private bool _m365NeedsSignIn = false;
    private string _primaryActionMode = "None"; // "Join", "SignIn", "Settings"
    private string _activeJoinUrl = string.Empty;

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOACTIVATE = 0x0010;

    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_TOOLWINDOW = 0x00000080;
    private const int WS_EX_TRANSPARENT = 0x00000020;
    private const int WS_EX_NOACTIVATE = 0x08000000;

    public MainWindow()
    {
        InitializeComponent();

        _settings = AppSettings.Load();
        if (_settings.CachedMeetings.Count > 0)
        {
            _m365Meetings = _settings.CachedMeetings;
        }

        ApplySettingsImmediately();
        SetupTrayIcon();

        _countdownTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1)
        };
        _countdownTimer.Tick += (_, _) => UpdateTileDisplay();

        _dockTrackingTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(50)
        };
        _dockTrackingTimer.Tick += (_, _) => UpdateDockingPosition();

        _calendarPollTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(60)
        };
        _calendarPollTimer.Tick += async (_, _) => await SyncAllCalendarsNowAsync();

        Loaded += MainWindow_Loaded;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        // Hide from Alt+Tab switcher so it behaves like an embedded Teams tile
        IntPtr hwnd = new WindowInteropHelper(this).Handle;
        int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
        SetWindowLong(hwnd, GWL_EXSTYLE, exStyle | WS_EX_TOOLWINDOW);

        MenuVersionHeader.Header = AppVersionInfo.FullDisplay;

        RebuildMergedMeetings();
        UpdateTileDisplay();

        _countdownTimer.Start();
        _dockTrackingTimer.Start();
        _calendarPollTimer.Start();

        if (_settings.EnableM365WebViewSync)
        {
            EnsureM365SyncWindow();
            if (_m365SyncWindow != null)
            {
                await _m365SyncWindow.InitializeInBackgroundAsync();
            }
        }

        await SyncAllCalendarsNowAsync();
    }

    private void EnsureM365SyncWindow()
    {
        if (_m365SyncWindow != null)
            return;

        _m365SyncWindow = new M365WebViewSyncWindow(_settings.M365CalendarUrl);
        _m365SyncWindow.MeetingsUpdated += OnM365MeetingsUpdated;
        _m365SyncWindow.StatusChanged += (status, needsLogin) =>
        {
            _m365StatusMessage = status;
            _m365NeedsSignIn = needsLogin;
            UpdateTileDisplay();
            _settingsWindow?.RefreshMeetingsList();
        };
    }

    private void OnM365MeetingsUpdated(List<MeetingItem> items)
    {
        _m365Meetings = items;
        _settings.CachedMeetings = items;
        _settings.Save();
        RebuildMergedMeetings();
        UpdateTileDisplay();
        _settingsWindow?.RefreshMeetingsList();
    }

    public async Task SyncAllCalendarsNowAsync()
    {
        DateTime winStart = DateTime.Now.AddHours(-1);
        DateTime winEnd = DateTime.Now.AddDays(14);

        if (_settings.EnableM365WebViewSync)
        {
            EnsureM365SyncWindow();
            if (_m365SyncWindow != null)
            {
                await _m365SyncWindow.InitializeInBackgroundAsync();
                _m365SyncWindow.TriggerManualRefresh();
            }
        }

        if (!string.IsNullOrWhiteSpace(_settings.IcsFeedUrl))
        {
            try
            {
                _icsMeetings = await IcsCalendarParser.FetchAndParseAsync(_settings.IcsFeedUrl, winStart, winEnd);
            }
            catch (Exception ex)
            {
                _m365StatusMessage = $"ICS error: {ex.Message}";
            }
        }
        else
        {
            _icsMeetings.Clear();
        }

        if (_settings.EnableClassicOutlookCom)
        {
            _comMeetings = await Task.Run(() => OutlookComService.TryReadUpcomingMeetings(winStart, winEnd, requireOutlookRunning: true));
        }
        else
        {
            _comMeetings.Clear();
        }

        RebuildMergedMeetings();
        UpdateTileDisplay();
    }

    public void RebuildMergedMeetings()
    {
        DateTime now = DateTime.Now;
        DateTime graceCutoff = now.AddMinutes(-Math.Max(1, _settings.GraceMinutesAfterStart));

        var combined = new List<MeetingItem>();
        combined.AddRange(_settings.ManualMeetings);
        combined.AddRange(_icsMeetings);
        combined.AddRange(_comMeetings);

        // Only include M365 WebView items that don't duplicate an authoritative ICS/COM meeting on the same day
        foreach (var m365Item in _m365Meetings)
        {
            bool alreadyInAuthoritative = combined.Any(auth =>
                auth.Title.Trim().Equals(m365Item.Title.Trim(), StringComparison.OrdinalIgnoreCase) &&
                auth.StartTime.Date == m365Item.StartTime.Date &&
                Math.Abs((auth.StartTime - m365Item.StartTime).TotalHours) <= 2.0);

            if (!alreadyInAuthoritative)
            {
                combined.Add(m365Item);
            }
        }

        // Deduplicate by normalized title + start minute
        var deduped = combined
            .Where(m => m.EndTime > now && (m.StartTime >= graceCutoff || _settings.IncludeInProgressMeetings))
            .GroupBy(m => $"{m.Title.Trim().ToLowerInvariant()}|{m.StartTime:yyyy-MM-dd HH:mm}")
            .Select(g =>
            {
                // Prefer item that has a Teams JoinUrl
                var withUrl = g.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.JoinUrl));
                return withUrl ?? g.First();
            })
            .OrderBy(m => m.StartTime)
            .ToList();

        // If no future/grace meetings exist, check if there is an in-progress meeting right now
        if (deduped.Count == 0)
        {
            deduped = combined
                .Where(m => m.StartTime <= now && m.EndTime > now)
                .GroupBy(m => $"{m.Title.Trim().ToLowerInvariant()}|{m.StartTime:yyyy-MM-dd HH:mm}")
                .Select(g => g.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.JoinUrl)) ?? g.First())
                .OrderBy(m => m.StartTime)
                .ToList();
        }

        _mergedUpcoming = deduped;
        if (_selectedMeetingIndex >= _mergedUpcoming.Count)
        {
            _selectedMeetingIndex = 0;
        }
    }

    private void UpdateTileDisplay()
    {
        DateTime now = DateTime.Now;

        // Re-filter if the currently displayed meeting has passed the grace window
        if (_mergedUpcoming.Count > 0)
        {
            DateTime graceCutoff = now.AddMinutes(-Math.Max(1, _settings.GraceMinutesAfterStart));
            if (_mergedUpcoming.Any(m => m.EndTime <= now || (m.StartTime < graceCutoff && !_settings.IncludeInProgressMeetings && _mergedUpcoming.Count > 1)))
            {
                RebuildMergedMeetings();
            }
        }

        if (_mergedUpcoming.Count == 0)
        {
            RenderNoMeetingsState();
            return;
        }

        if (_selectedMeetingIndex < 0 || _selectedMeetingIndex >= _mergedUpcoming.Count)
            _selectedMeetingIndex = 0;

        MeetingItem current = _mergedUpcoming[_selectedMeetingIndex];
        TimeSpan untilStart = current.StartTime - now;
        string dhmText = current.GetDhmCountdown(now);
        string startText = current.GetFormattedStartTime(now);

        // Header status & count
        string indexSuffix = _mergedUpcoming.Count > 1 ? $" ({_selectedMeetingIndex + 1}/{_mergedUpcoming.Count})" : string.Empty;
        if (untilStart.TotalSeconds <= 0)
        {
            TxtHeaderStatus.Text = "NOW" + indexSuffix;
        }
        else if (untilStart.TotalMinutes <= 15)
        {
            TxtHeaderStatus.Text = "SOON" + indexSuffix;
        }
        else
        {
            TxtHeaderStatus.Text = "NEXT" + indexSuffix;
        }

        TxtDhmCountdown.Text = dhmText;
        TxtMeetingTitle.Text = current.Title;
        TxtMeetingTitle.ToolTip = $"{current.Title}\n{current.GetFormattedTimeRange(now)}\nCountdown: {dhmText}\nSource: {current.Source}";
        TxtStartTime.Text = $"Start: {startText}";
        TxtStartTime.ToolTip = $"Starts: {current.StartTime:dddd, dd MMM yyyy h:mm tt}\nEnds: {current.EndTime:h:mm tt}";

        // Collapsed slim bar
        TxtCollapsedStart.Text = startText;
        TxtCollapsedDhm.Text = dhmText;
        TxtCollapsedTitle.Text = current.Title;

        // Cycle buttons visibility
        CyclePanel.Visibility = _mergedUpcoming.Count > 1 ? Visibility.Visible : Visibility.Collapsed;

        // Urgency colors
        UpdateUrgencyColors(untilStart);

        // Join button
        if (current.HasTeamsLink)
        {
            _primaryActionMode = "Join";
            _activeJoinUrl = current.JoinUrl;
            BtnPrimaryAction.Content = "Join";
            BtnPrimaryAction.Visibility = Visibility.Visible;
        }
        else
        {
            _primaryActionMode = "None";
            _activeJoinUrl = string.Empty;
            BtnPrimaryAction.Visibility = Visibility.Collapsed;
        }

        if (_trayIcon != null)
        {
            string trayText = $"{current.Title} ({startText} | {dhmText})";
            _trayIcon.Text = trayText.Length > 63 ? trayText.Substring(0, 63) : trayText;
        }
    }

    private void RenderNoMeetingsState()
    {
        CyclePanel.Visibility = Visibility.Collapsed;

        if (_m365NeedsSignIn && _settings.EnableM365WebViewSync)
        {
            TxtHeaderStatus.Text = "SYNC";
            TxtDhmCountdown.Text = "--d --h --m";
            TxtMeetingTitle.Text = "Sign in once to sync M365 Calendar";
            TxtStartTime.Text = "Click 'Sign In' →";
            TxtCollapsedStart.Text = "M365";
            TxtCollapsedDhm.Text = "Sign In";
            TxtCollapsedTitle.Text = "Connect Calendar";

            _primaryActionMode = "SignIn";
            BtnPrimaryAction.Content = "Sign In";
            BtnPrimaryAction.Visibility = Visibility.Visible;
            UpdateUrgencyColors(TimeSpan.FromMinutes(10));
            return;
        }

        TxtHeaderStatus.Text = "NEXT";
        TxtDhmCountdown.Text = "0d 0h 0m";
        TxtMeetingTitle.Text = "No upcoming meetings found";
        TxtStartTime.Text = "Start: None scheduled";
        TxtCollapsedStart.Text = "Free";
        TxtCollapsedDhm.Text = "0d 0h 0m";
        TxtCollapsedTitle.Text = "No upcoming meetings";

        _primaryActionMode = "Settings";
        BtnPrimaryAction.Content = "Sync";
        BtnPrimaryAction.Visibility = Visibility.Visible;
        UpdateUrgencyColors(TimeSpan.FromHours(2));
    }

    private void UpdateUrgencyColors(TimeSpan untilStart)
    {
        if (untilStart.TotalMinutes <= 5)
        {
            // Imminent / In progress (Coral Red)
            var accent = new SolidColorBrush(System.Windows.Media.Color.FromRgb(227, 93, 106));
            AccentStripe.Background = accent;
            StatusDot.Fill = accent;
            CountdownBadge.BorderBrush = accent;
            CountdownBadge.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(68, 34, 38));
            TxtDhmCountdown.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 210, 214));
        }
        else if (untilStart.TotalMinutes <= 15)
        {
            // Starting soon (Amber Gold)
            var accent = new SolidColorBrush(System.Windows.Media.Color.FromRgb(242, 200, 17));
            AccentStripe.Background = accent;
            StatusDot.Fill = accent;
            CountdownBadge.BorderBrush = accent;
            CountdownBadge.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(64, 54, 24));
            TxtDhmCountdown.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 238, 160));
        }
        else
        {
            // Normal (Teams Brand Purple)
            var accent = new SolidColorBrush(System.Windows.Media.Color.FromRgb(91, 95, 199));
            var dotBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(127, 133, 245));
            AccentStripe.Background = accent;
            StatusDot.Fill = dotBrush;
            CountdownBadge.BorderBrush = accent;
            CountdownBadge.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 52, 74));
            TxtDhmCountdown.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(232, 234, 255));
        }
    }

    private void UpdateDockingPosition()
    {
        if (_isDragging)
            return;

        MenuDockToggle.Header = _settings.DockToTeamsWindow
            ? "✓ Dock to Microsoft Teams Window"
            : "Dock to Microsoft Teams Window";

        if (!_settings.DockToTeamsWindow)
        {
            ShowOverlayVisible();
            return;
        }

        IntPtr ownHwnd = new WindowInteropHelper(this).Handle;
        IntPtr settingsHwnd = _settingsWindow != null && _settingsWindow.IsLoaded ? _settingsWindow.Hwnd : IntPtr.Zero;
        IntPtr syncHwnd = _m365SyncWindow != null && _m365SyncWindow.IsLoaded ? _m365SyncWindow.Hwnd : IntPtr.Zero;

        TeamsWindowState state = TeamsWindowTracker.GetTeamsState(
            ownHwnd,
            settingsHwnd,
            syncHwnd,
            _settings.OffsetX,
            _settings.OffsetY,
            ActualWidth > 0 ? ActualWidth : _settings.TileWidth,
            ActualHeight > 0 ? ActualHeight : 72);

        if (!state.IsTeamsRunning || state.IsTeamsMinimized || state.TeamsHandle == IntPtr.Zero)
        {
            if (_settings.AutoHideWhenTeamsCovered)
            {
                HideOverlayInvisible();
            }
            else
            {
                ShowOverlayVisible();
            }
            return;
        }

        _lastTeamsBounds = state.TeamsBounds;
        _lastTeamsDpiScale = state.DpiScale;
        _hasLastTeamsBounds = true;

        if (state.IsTileAreaOccluded && _settings.AutoHideWhenTeamsCovered)
        {
            HideOverlayInvisible();
            return;
        }

        ShowOverlayVisible();

        // Convert physical screen pixels of Teams window + scaled offset to WPF DIPs
        double targetPhysicalX = state.TeamsBounds.Left + (_settings.OffsetX * state.DpiScale);
        double targetPhysicalY = state.TeamsBounds.Top + (_settings.OffsetY * state.DpiScale);

        PresentationSource? source = PresentationSource.FromVisual(this);
        if (source?.CompositionTarget != null)
        {
            System.Windows.Point dipPt = source.CompositionTarget.TransformFromDevice.Transform(
                new System.Windows.Point(targetPhysicalX, targetPhysicalY));

            if (Math.Abs(Left - dipPt.X) > 0.5)
                Left = dipPt.X;
            if (Math.Abs(Top - dipPt.Y) > 0.5)
                Top = dipPt.Y;
        }
        else
        {
            Left = targetPhysicalX;
            Top = targetPhysicalY;
        }
    }

    private void ShowOverlayVisible()
    {
        IntPtr hwnd = new WindowInteropHelper(this).Handle;
        if (Opacity < 1.0)
        {
            Opacity = 1.0;
            IsHitTestVisible = true;
            if (hwnd != IntPtr.Zero)
            {
                int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
                if ((exStyle & WS_EX_TRANSPARENT) != 0)
                {
                    SetWindowLong(hwnd, GWL_EXSTYLE, exStyle & ~WS_EX_TRANSPARENT);
                }
            }
        }

        if (hwnd != IntPtr.Zero)
        {
            // Keep topmost Z-order active even after overnight Windows lock screen / DWM display wake
            SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        }
    }

    private void HideOverlayInvisible()
    {
        if (Opacity > 0.0)
        {
            Opacity = 0.0;
            IsHitTestVisible = false;
            IntPtr hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd != IntPtr.Zero)
            {
                int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
                if ((exStyle & WS_EX_TRANSPARENT) == 0)
                {
                    SetWindowLong(hwnd, GWL_EXSTYLE, exStyle | WS_EX_TRANSPARENT);
                }
            }
        }
    }

    public void ApplySettingsImmediately()
    {
        Width = Math.Clamp(_settings.TileWidth, 175, 420);
        ExpandedPanel.Visibility = _settings.IsCollapsed ? Visibility.Collapsed : Visibility.Visible;
        CollapsedPanel.Visibility = _settings.IsCollapsed ? Visibility.Visible : Visibility.Collapsed;
        CardBorder.Cursor = _settings.LockPosition ? System.Windows.Input.Cursors.Arrow : System.Windows.Input.Cursors.SizeAll;

        RebuildMergedMeetings();
        UpdateTileDisplay();
        UpdateDockingPosition();
    }

    private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_settings.LockPosition)
            return;

        if (e.ButtonState == MouseButtonState.Pressed)
        {
            _isDragging = true;
            try
            {
                DragMove();
            }
            catch { }
            finally
            {
                _isDragging = false;
                SaveDraggedOffset();
            }
        }
    }

    private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_isDragging)
        {
            _isDragging = false;
            SaveDraggedOffset();
        }
    }

    private void SaveDraggedOffset()
    {
        if (_settings.DockToTeamsWindow && _hasLastTeamsBounds)
        {
            PresentationSource? source = PresentationSource.FromVisual(this);
            if (source?.CompositionTarget != null)
            {
                System.Windows.Point physPt = source.CompositionTarget.TransformToDevice.Transform(new System.Windows.Point(Left, Top));
                double scale = _lastTeamsDpiScale > 0 ? _lastTeamsDpiScale : 1.0;
                _settings.OffsetX = Math.Round((physPt.X - _lastTeamsBounds.Left) / scale);
                _settings.OffsetY = Math.Round((physPt.Y - _lastTeamsBounds.Top) / scale);
            }
        }
        else
        {
            _settings.StandaloneLeft = Left;
            _settings.StandaloneTop = Top;
        }
        _settings.Save();
    }

    private void BtnCollapse_Click(object sender, RoutedEventArgs e)
    {
        _settings.IsCollapsed = !_settings.IsCollapsed;
        _settings.Save();
        ApplySettingsImmediately();
    }

    private void BtnSettings_Click(object sender, RoutedEventArgs e)
    {
        OpenSettingsWindow();
    }

    private void BtnPrevMeeting_Click(object sender, RoutedEventArgs e)
    {
        if (_mergedUpcoming.Count <= 1)
            return;

        _selectedMeetingIndex = (_selectedMeetingIndex - 1 + _mergedUpcoming.Count) % _mergedUpcoming.Count;
        UpdateTileDisplay();
    }

    private void BtnNextMeeting_Click(object sender, RoutedEventArgs e)
    {
        if (_mergedUpcoming.Count <= 1)
            return;

        _selectedMeetingIndex = (_selectedMeetingIndex + 1) % _mergedUpcoming.Count;
        UpdateTileDisplay();
    }

    private void BtnPrimaryAction_Click(object sender, RoutedEventArgs e)
    {
        if (_primaryActionMode == "Join" && !string.IsNullOrWhiteSpace(_activeJoinUrl))
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = _activeJoinUrl,
                    UseShellExecute = true
                });
            }
            catch { }
        }
        else if (_primaryActionMode == "SignIn")
        {
            ShowM365CalendarSyncWindow();
        }
        else
        {
            OpenSettingsWindow();
        }
    }

    public void OpenSettingsWindow()
    {
        if (_settingsWindow == null || !_settingsWindow.IsLoaded)
        {
            _settingsWindow = new SettingsWindow(this, _settings);
            _settingsWindow.Closed += (_, _) => _settingsWindow = null;
            _settingsWindow.Show();
        }
        else
        {
            _settingsWindow.RefreshMeetingsList();
            _settingsWindow.Activate();
        }
    }

    public void ShowM365CalendarSyncWindow()
    {
        EnsureM365SyncWindow();
        _m365SyncWindow?.ShowInteractiveWindow();
    }

    public string GetM365StatusText() => _m365StatusMessage;

    public List<MeetingItem> GetAllUpcomingMeetings() => _mergedUpcoming.ToList();

    private void SetupTrayIcon()
    {
        _trayIcon = new Forms.NotifyIcon
        {
            Icon = CreateTrayBadgeIcon(),
            Visible = true,
            Text = "Teams Next Meeting Tile"
        };

        var menu = new Forms.ContextMenuStrip();
        var versionItem = new Forms.ToolStripMenuItem(AppVersionInfo.FullDisplay) { Enabled = false };
        menu.Items.Add(versionItem);
        menu.Items.Add(new Forms.ToolStripSeparator());
        menu.Items.Add("Settings & Upcoming Meetings...", null, (_, _) => Dispatcher.Invoke(OpenSettingsWindow));
        menu.Items.Add("Open M365 Calendar Window (Sign In)", null, (_, _) => Dispatcher.Invoke(ShowM365CalendarSyncWindow));
        menu.Items.Add("Sync Calendar Now", null, async (_, _) => await Dispatcher.InvokeAsync(SyncAllCalendarsNowAsync));
        menu.Items.Add(new Forms.ToolStripSeparator());
        menu.Items.Add("Snap Above 'Quick views' in Teams", null, (_, _) => Dispatcher.Invoke(() =>
        {
            _settings.DockToTeamsWindow = true;
            _settings.OffsetX = 52;
            _settings.OffsetY = 122;
            _settings.Save();
            ApplySettingsImmediately();
        }));
        menu.Items.Add("Quit", null, (_, _) => Dispatcher.Invoke(ShutdownApplication));

        _trayIcon.ContextMenuStrip = menu;
        _trayIcon.DoubleClick += (_, _) => Dispatcher.Invoke(OpenSettingsWindow);
    }

    private static System.Drawing.Icon CreateTrayBadgeIcon()
    {
        using var bmp = new Bitmap(32, 32);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(System.Drawing.Color.Transparent);
            using var bgBrush = new SolidBrush(System.Drawing.Color.FromArgb(91, 95, 199));
            g.FillEllipse(bgBrush, 2, 2, 28, 28);
            using var pen = new System.Drawing.Pen(System.Drawing.Color.White, 2.5f);
            g.DrawEllipse(pen, 7, 7, 18, 18);
            g.DrawLine(pen, 16, 11, 16, 16);
            g.DrawLine(pen, 16, 16, 20, 18);
        }
        return System.Drawing.Icon.FromHandle(bmp.GetHicon());
    }

    private void MenuSettings_Click(object sender, RoutedEventArgs e) => OpenSettingsWindow();
    private void MenuOpenM365_Click(object sender, RoutedEventArgs e) => ShowM365CalendarSyncWindow();
    private async void MenuSyncNow_Click(object sender, RoutedEventArgs e) => await SyncAllCalendarsNowAsync();

    private void MenuSnapDefault_Click(object sender, RoutedEventArgs e)
    {
        _settings.DockToTeamsWindow = true;
        _settings.OffsetX = 52;
        _settings.OffsetY = 122;
        _settings.Save();
        ApplySettingsImmediately();
    }

    private void MenuDockToggle_Click(object sender, RoutedEventArgs e)
    {
        _settings.DockToTeamsWindow = !_settings.DockToTeamsWindow;
        _settings.Save();
        ApplySettingsImmediately();
    }

    private void MenuExit_Click(object sender, RoutedEventArgs e) => ShutdownApplication();

    public void ShutdownApplication()
    {
        _countdownTimer.Stop();
        _dockTrackingTimer.Stop();
        _calendarPollTimer.Stop();

        if (_trayIcon != null)
        {
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _trayIcon = null;
        }

        _m365SyncWindow?.ForceClose();
        _settingsWindow?.Close();
        Application.Current.Shutdown();
    }
}
