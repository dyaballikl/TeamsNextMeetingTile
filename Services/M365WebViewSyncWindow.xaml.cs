using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using Microsoft.Web.WebView2.Core;
using TeamsNextMeetingTile.Models;

namespace TeamsNextMeetingTile.Services;

public partial class M365WebViewSyncWindow : Window
{
    public event Action<List<MeetingItem>>? MeetingsUpdated;
    public event Action<string, bool>? StatusChanged; // (statusMessage, requiresInteractiveLogin)

    private readonly DispatcherTimer _domPollTimer;
    private readonly DispatcherTimer _pageReloadTimer;
    private readonly Dictionary<string, MeetingItem> _discoveredMeetings = new(StringComparer.OrdinalIgnoreCase);
    private bool _isInitialized;
    private bool _allowClose;
    private string _targetUrl;

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_TOOLWINDOW = 0x00000080;
    private const int WS_EX_TRANSPARENT = 0x00000020;
    private const int WS_EX_NOACTIVATE = 0x08000000;

    public IntPtr Hwnd => new WindowInteropHelper(this).Handle;
    public bool IsInBackground { get; private set; } = true;
    public bool NeedsInteractiveSignIn { get; private set; }
    public string LastStatusText { get; private set; } = "Initializing M365 sync...";

    public M365WebViewSyncWindow(string calendarUrl)
    {
        InitializeComponent();
        LocationChanged += (_, _) =>
        {
            // Prevent Windows display sleep/wake from moving the hidden background window back onto the visible monitor
            if (IsInBackground && (Left > -10000 || Top > -10000))
            {
                Left = -20000;
                Top = -20000;
            }
        };
        _targetUrl = string.IsNullOrWhiteSpace(calendarUrl)
            ? "https://outlook.office.com/calendar/view/workweek"
            : calendarUrl;

        _domPollTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(12)
        };
        _domPollTimer.Tick += async (_, _) => await PollDomAndTokensAsync();

        _pageReloadTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMinutes(10)
        };
        _pageReloadTimer.Tick += (_, _) =>
        {
            if (_isInitialized && !NeedsInteractiveSignIn && CalendarWebView.CoreWebView2 != null)
            {
                try { CalendarWebView.CoreWebView2.Reload(); } catch { }
            }
        };
    }

    public async Task InitializeInBackgroundAsync()
    {
        if (_isInitialized)
            return;

        try
        {
            // Ensure the window handle is created off-screen so WebView2 renders the full viewport
            Show();
            SendToBackground();

            string userDataFolder = Path.Combine(AppSettings.DataFolder, "WebView2_M365Profile");
            Directory.CreateDirectory(userDataFolder);

            var options = new CoreWebView2EnvironmentOptions
            {
                AllowSingleSignOnUsingOSPrimaryAccount = true
            };

            CoreWebView2Environment env = await CoreWebView2Environment.CreateAsync(null, userDataFolder, options);
            await CalendarWebView.EnsureCoreWebView2Async(env);

            CalendarWebView.CoreWebView2.Settings.AreDevToolsEnabled = false;
            CalendarWebView.CoreWebView2.Settings.IsStatusBarEnabled = false;

            await CalendarWebView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(GetInjectedInterceptorScript());

            CalendarWebView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
            CalendarWebView.CoreWebView2.NavigationCompleted += CoreWebView2_NavigationCompleted;
            CalendarWebView.CoreWebView2.SourceChanged += CoreWebView2_SourceChanged;

            _isInitialized = true;
            UpdateStatus("Connecting to Outlook Calendar...", false);
            CalendarWebView.CoreWebView2.Navigate(_targetUrl);

            _domPollTimer.Start();
            _pageReloadTimer.Start();
        }
        catch (Exception ex)
        {
            UpdateStatus($"WebView2 init error: {ex.Message}", false);
        }
    }

    public void ShowInteractiveWindow()
    {
        IsInBackground = false;
        ShowInTaskbar = true;
        ShowActivated = true;
        Opacity = 1.0;

        IntPtr hwnd = Hwnd;
        if (hwnd != IntPtr.Zero)
        {
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            exStyle &= ~(WS_EX_TOOLWINDOW | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE);
            SetWindowLong(hwnd, GWL_EXSTYLE, exStyle);
        }

        double screenW = SystemParameters.WorkArea.Width;
        double screenH = SystemParameters.WorkArea.Height;
        Left = Math.Max(40, (screenW - Width) / 2);
        Top = Math.Max(40, (screenH - Height) / 2);

        if (!IsVisible)
            Show();
        if (WindowState == WindowState.Minimized)
            WindowState = WindowState.Normal;

        Activate();
        Focus();
    }

    public void SendToBackground()
    {
        IsInBackground = true;
        ShowInTaskbar = false;
        ShowActivated = false;
        WindowState = WindowState.Normal;
        Left = -20000;
        Top = -20000;

        IntPtr hwnd = Hwnd;
        if (hwnd != IntPtr.Zero)
        {
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            exStyle |= (WS_EX_TOOLWINDOW | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE);
            SetWindowLong(hwnd, GWL_EXSTYLE, exStyle);
        }
    }

    public void ForceClose()
    {
        _allowClose = true;
        _domPollTimer.Stop();
        _pageReloadTimer.Stop();
        Close();
    }

    public void TriggerManualRefresh()
    {
        if (!_isInitialized || CalendarWebView.CoreWebView2 == null)
            return;

        try
        {
            _ = PollDomAndTokensAsync();
            CalendarWebView.CoreWebView2.Reload();
        }
        catch { }
    }

    private void CoreWebView2_SourceChanged(object? sender, CoreWebView2SourceChangedEventArgs e)
    {
        CheckCurrentUrlForAuthPrompt();
    }

    private async void CoreWebView2_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        CheckCurrentUrlForAuthPrompt();
        if (!NeedsInteractiveSignIn)
        {
            await Task.Delay(2500);
            await PollDomAndTokensAsync();
        }
    }

    private void CheckCurrentUrlForAuthPrompt()
    {
        if (CalendarWebView.CoreWebView2 == null)
            return;

        string src = CalendarWebView.CoreWebView2.Source ?? string.Empty;
        if (src.Contains("login.microsoftonline.com", StringComparison.OrdinalIgnoreCase) ||
            src.Contains("login.live.com", StringComparison.OrdinalIgnoreCase) ||
            src.Contains("autologon.microsoftazuread-sso.com", StringComparison.OrdinalIgnoreCase))
        {
            UpdateStatus("Sign-in required in M365 Calendar window", true);
        }
        else if (src.Contains("outlook.office", StringComparison.OrdinalIgnoreCase) ||
                 src.Contains("outlook.live.com", StringComparison.OrdinalIgnoreCase))
        {
            UpdateStatus("Connected to M365 Outlook Calendar", false);
        }
    }

    private async Task PollDomAndTokensAsync()
    {
        if (!_isInitialized || CalendarWebView.CoreWebView2 == null)
            return;

        CheckCurrentUrlForAuthPrompt();
        if (NeedsInteractiveSignIn)
            return;

        try
        {
            await CalendarWebView.CoreWebView2.ExecuteScriptAsync("window.__agPollCalendar && window.__agPollCalendar();");
        }
        catch
        {
            // Ignore transient script errors during navigation
        }
    }

    private void CoreWebView2_WebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            string json = e.WebMessageAsJson;
            if (string.IsNullOrWhiteSpace(json))
                return;

            using JsonDocument doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;

            if (root.ValueKind == JsonValueKind.String)
            {
                string? innerStr = root.GetString();
                if (!string.IsNullOrWhiteSpace(innerStr))
                {
                    using JsonDocument innerDoc = JsonDocument.Parse(innerStr);
                    ProcessIncomingPayload(innerDoc.RootElement);
                }
            }
            else
            {
                ProcessIncomingPayload(root);
            }
        }
        catch
        {
            // Ignore malformed message
        }
    }

    private void ProcessIncomingPayload(JsonElement payload)
    {
        if (payload.ValueKind != JsonValueKind.Object)
            return;

        if (!payload.TryGetProperty("type", out JsonElement typeEl))
            return;

        string type = typeEl.GetString() ?? string.Empty;
        if (type == "auth_needed")
        {
            UpdateStatus("Click 'Sign In' to complete M365 login", true);
            return;
        }

        if (type == "events" && payload.TryGetProperty("items", out JsonElement itemsEl) && itemsEl.ValueKind == JsonValueKind.Array)
        {
            bool changed = false;
            DateTime cutoffPast = DateTime.Now.AddHours(-2);
            DateTime cutoffFuture = DateTime.Now.AddDays(14);

            foreach (JsonElement el in itemsEl.EnumerateArray())
            {
                string title = GetStringProp(el, "title").Trim();
                string startStr = GetStringProp(el, "start").Trim();
                string endStr = GetStringProp(el, "end").Trim();
                string startTz = GetStringProp(el, "startTz").Trim();
                string endTz = GetStringProp(el, "endTz").Trim();
                string location = GetStringProp(el, "location").Trim();
                string joinUrl = GetStringProp(el, "joinUrl").Trim();
                bool isAllDay = GetBoolProp(el, "isAllDay");
                bool isCancelled = GetBoolProp(el, "isCancelled");

                if (string.IsNullOrWhiteSpace(title) || isCancelled || isAllDay)
                    continue;

                if (!TryParseFlexibleDateTime(startStr, startTz, out DateTime startLocal))
                    continue;

                if (!TryParseFlexibleDateTime(endStr, string.IsNullOrEmpty(endTz) ? startTz : endTz, out DateTime endLocal) || endLocal <= startLocal)
                {
                    endLocal = startLocal.AddMinutes(30);
                }

                if (endLocal < cutoffPast || startLocal > cutoffFuture)
                    continue;

                if (string.IsNullOrWhiteSpace(joinUrl))
                {
                    joinUrl = IcsCalendarParser.ExtractTeamsJoinUrl(location, GetStringProp(el, "body"));
                }

                string key = $"{title.ToLowerInvariant()}|{startLocal:yyyy-MM-dd HH:mm}";
                _discoveredMeetings[key] = new MeetingItem
                {
                    Id = $"M365_{key}",
                    Title = title,
                    StartTime = startLocal,
                    EndTime = endLocal,
                    Location = location,
                    JoinUrl = joinUrl,
                    Source = "M365"
                };
                changed = true;
            }

            // Prune expired events older than 2 hours
            var expiredKeys = _discoveredMeetings
                .Where(kv => kv.Value.EndTime < cutoffPast)
                .Select(kv => kv.Key)
                .ToList();
            foreach (string k in expiredKeys)
            {
                _discoveredMeetings.Remove(k);
                changed = true;
            }

            if (changed || _discoveredMeetings.Count > 0)
            {
                var ordered = _discoveredMeetings.Values
                    .OrderBy(m => m.StartTime)
                    .ToList();
                UpdateStatus($"Synced {ordered.Count} upcoming meeting(s) from M365", false);
                MeetingsUpdated?.Invoke(ordered);
            }
        }
    }

    private static bool TryParseFlexibleDateTime(string input, string? tzId, out DateTime localDt)
    {
        localDt = default;
        if (string.IsNullOrWhiteSpace(input))
            return false;

        // If an explicit Windows/IANA timezone ID was provided alongside an un-offset timestamp
        if (!string.IsNullOrWhiteSpace(tzId) &&
            !tzId.Equals("UTC", StringComparison.OrdinalIgnoreCase) &&
            !input.EndsWith("Z", StringComparison.OrdinalIgnoreCase) &&
            DateTime.TryParse(input, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime rawWall))
        {
            localDt = IcsCalendarParser.ConvertWallClockToLocal(rawWall, tzId, false);
            return true;
        }

        if (DateTimeOffset.TryParse(input, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTimeOffset dto))
        {
            localDt = dto.LocalDateTime;
            return true;
        }

        if (DateTime.TryParse(input, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out DateTime dt))
        {
            localDt = dt;
            return true;
        }

        return false;
    }

    private static string GetStringProp(JsonElement el, string name)
    {
        if (el.TryGetProperty(name, out JsonElement p) && p.ValueKind == JsonValueKind.String)
            return p.GetString() ?? string.Empty;
        return string.Empty;
    }

    private static bool GetBoolProp(JsonElement el, string name)
    {
        if (el.TryGetProperty(name, out JsonElement p))
        {
            if (p.ValueKind == JsonValueKind.True) return true;
            if (p.ValueKind == JsonValueKind.False) return false;
        }
        return false;
    }

    private void UpdateStatus(string message, bool needsLogin)
    {
        LastStatusText = message;
        NeedsInteractiveSignIn = needsLogin;
        TxtSyncBanner.Text = needsLogin
            ? "Please sign into your Microsoft 365 account below. Once your calendar loads, click 'Run in Background'."
            : $"Status: {message} — You can click 'Run in Background (Hide)' anytime.";
        StatusChanged?.Invoke(message, needsLogin);
    }

    private void BtnRefreshPage_Click(object sender, RoutedEventArgs e)
    {
        TriggerManualRefresh();
    }

    private void BtnHideBackground_Click(object sender, RoutedEventArgs e)
    {
        SendToBackground();
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (!_allowClose)
        {
            e.Cancel = true;
            SendToBackground();
        }
    }

    private static string GetInjectedInterceptorScript()
    {
        return @"
(function() {
    if (window.__agCalendarHookInstalled) return;
    window.__agCalendarHookInstalled = true;

    function sendEvents(items) {
        if (!items || !items.length) return;
        try {
            window.chrome.webview.postMessage({ type: 'events', items: items });
        } catch (e) {}
    }

    function extractTeamsUrl(text) {
        if (!text || typeof text !== 'string') return '';
        var m = text.match(/https:\/\/(?:teams\.microsoft\.com\/l\/meetup-join\/[^\s""<>\\]+|teams\.live\.com\/meet\/[^\s""<>\\]+)/i);
        return m ? m[0] : '';
    }

    function walkJsonForCalendarItems(obj, found, depth) {
        if (!obj || typeof obj !== 'object' || depth > 9) return;
        if (Array.isArray(obj)) {
            for (var i = 0; i < obj.length; i++) {
                walkJsonForCalendarItems(obj[i], found, depth + 1);
            }
            return;
        }

        var subject = obj.Subject || obj.subject || obj.Title || obj.title;
        var rawStart = obj.Start || obj.start || obj.StartTime || obj.startTime;
        var rawEnd = obj.End || obj.end || obj.EndTime || obj.endTime;

        var startStr = typeof rawStart === 'string' ? rawStart : (rawStart && (rawStart.DateTime || rawStart.dateTime));
        var endStr = typeof rawEnd === 'string' ? rawEnd : (rawEnd && (rawEnd.DateTime || rawEnd.dateTime));
        var startTz = rawStart && (rawStart.TimeZone || rawStart.timeZone);
        var endTz = rawEnd && (rawEnd.TimeZone || rawEnd.timeZone);

        if (startStr && startTz === 'UTC' && !startStr.endsWith('Z')) startStr += 'Z';
        if (endStr && endTz === 'UTC' && !endStr.endsWith('Z')) endStr += 'Z';

        if (typeof subject === 'string' && typeof startStr === 'string' && startStr.length >= 10) {
            var loc = '';
            if (typeof obj.Location === 'string') loc = obj.Location;
            else if (obj.Location && obj.Location.DisplayName) loc = obj.Location.DisplayName;
            else if (obj.location && obj.location.displayName) loc = obj.location.displayName;

            var joinUrl = obj.OnlineMeetingJoinUrl || obj.onlineMeetingUrl || obj.JoinUrl ||
                          (obj.onlineMeeting && obj.onlineMeeting.joinUrl) ||
                          (obj.OnlineMeeting && obj.OnlineMeeting.JoinUrl) || '';
            if (!joinUrl) {
                joinUrl = extractTeamsUrl(JSON.stringify(obj));
            }

            var isAllDay = Boolean(obj.IsAllDayEvent || obj.isAllDay || obj.IsAllDay);
            var isCancelled = Boolean(obj.IsCancelled || obj.isCancelled);
            var busyType = obj.FreeBusyType || obj.showAs || '';
            if (busyType === 'Free' || busyType === 'free') isCancelled = true;

            found.push({
                title: subject,
                start: startStr,
                end: endStr || '',
                startTz: startTz || '',
                endTz: endTz || '',
                location: loc || '',
                joinUrl: joinUrl || '',
                isAllDay: isAllDay,
                isCancelled: isCancelled
            });
        }

        var keys = Object.keys(obj);
        for (var k = 0; k < keys.length; k++) {
            var val = obj[keys[k]];
            if (val && typeof val === 'object') {
                walkJsonForCalendarItems(val, found, depth + 1);
            }
        }
    }

    function inspectJsonText(text) {
        if (!text || text.length < 20) return;
        if (text.indexOf('Subject') === -1 && text.indexOf('subject') === -1 && text.indexOf('CalendarItem') === -1) return;
        try {
            var parsed = JSON.parse(text);
            var found = [];
            walkJsonForCalendarItems(parsed, found, 0);
            if (found.length > 0) sendEvents(found);
        } catch (e) {}
    }

    // Hook window.fetch
    var origFetch = window.fetch;
    if (origFetch) {
        window.fetch = async function() {
            var response = await origFetch.apply(this, arguments);
            try {
                var clone = response.clone();
                clone.text().then(inspectJsonText).catch(function(){});
            } catch (e) {}
            return response;
        };
    }

    // Hook XMLHttpRequest
    var origOpen = XMLHttpRequest.prototype.open;
    var origSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function() {
        return origOpen.apply(this, arguments);
    };
    XMLHttpRequest.prototype.send = function() {
        this.addEventListener('load', function() {
            try {
                if (!this.responseType || this.responseType === 'text') {
                    inspectJsonText(this.responseText);
                } else if (this.responseType === 'json' && this.response) {
                    var found = [];
                    walkJsonForCalendarItems(this.response, found, 0);
                    if (found.length > 0) sendEvents(found);
                }
            } catch (e) {}
        });
        return origSend.apply(this, arguments);
    };

    // Parse aria-labels from Outlook Web Calendar DOM
    function parseDomCalendarEvents() {
        var nodes = document.querySelectorAll('[aria-label]');
        var results = [];
        var now = new Date();

        // Matches patterns like:
        // 'Defect Triage, 10:30 AM to 11:00 AM, Friday, October 30, 2026...'
        // or 'Defect Triage, 10:30 am to 11:00 am, Friday, 30 October 2026...'
        var timeRangeRegex = /(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)?)\s*(?:to|–|-)\s*(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)?)/i;

        for (var i = 0; i < nodes.length; i++) {
            var label = nodes[i].getAttribute('aria-label') || '';
            if (label.length < 12 || label.length > 600) continue;
            if (label.toLowerCase().indexOf('all day') !== -1) continue;
            if (label.toLowerCase().indexOf('canceled:') === 0 || label.toLowerCase().indexOf('cancelled:') === 0) continue;

            var tm = label.match(timeRangeRegex);
            if (!tm) continue;

            var beforeTime = label.substring(0, tm.index).replace(/,\s*$/, '').trim();
            var afterTime = label.substring(tm.index + tm[0].length);

            if (!beforeTime || beforeTime.length < 2) continue;
            // Filter out generic grid slot labels like '10:30 AM to 11:00 AM' with no title
            if (/^(?:busy|free|tentative|time slot|create event|new event)$/i.test(beforeTime)) continue;

            var startTimeStr = tm[1].trim();
            var endTimeStr = tm[2].trim();

            var baseDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());

            // Try to extract date from afterTime or beforeTime
            var fullText = label;
            var dm1 = fullText.match(/([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})/); // October 30, 2026
            var dm2 = fullText.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);  // 30 October 2026
            if (dm1) {
                var dParsed1 = Date.parse(dm1[1] + ' ' + dm1[2] + ', ' + dm1[3]);
                if (!isNaN(dParsed1)) baseDate = new Date(dParsed1);
            } else if (dm2) {
                var dParsed2 = Date.parse(dm2[2] + ' ' + dm2[1] + ', ' + dm2[3]);
                if (!isNaN(dParsed2)) baseDate = new Date(dParsed2);
            }

            var startDt = combineDateAndTime(baseDate, startTimeStr);
            var endDt = combineDateAndTime(baseDate, endTimeStr);
            if (!startDt) continue;
            if (!endDt || endDt <= startDt) {
                endDt = new Date(startDt.getTime() + 30 * 60000);
            }

            var isTeams = /Microsoft Teams Meeting|Teams Meeting/i.test(label);

            results.push({
                title: beforeTime,
                start: startDt.toISOString(),
                end: endDt.toISOString(),
                location: isTeams ? 'Microsoft Teams Meeting' : '',
                joinUrl: '',
                isAllDay: false,
                isCancelled: false
            });
        }

        if (results.length > 0) {
            sendEvents(results);
        }
    }

    function combineDateAndTime(dateObj, timeStr) {
        var m = timeStr.match(/^(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)?$/);
        if (!m) return null;
        var hh = parseInt(m[1], 10);
        var mm = parseInt(m[2], 10);
        var ampm = m[3] ? m[3].toUpperCase() : '';
        if (ampm === 'PM' && hh < 12) hh += 12;
        if (ampm === 'AM' && hh === 12) hh = 0;
        return new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), hh, mm, 0);
    }

    async function tryGraphTokenFetch() {
        try {
            for (var i = 0; i < localStorage.length; i++) {
                var k = localStorage.key(i);
                if (!k || k.toLowerCase().indexOf('accesstoken') === -1) continue;
                var raw = localStorage.getItem(k);
                if (!raw || raw.indexOf('secret') === -1) continue;
                var obj = JSON.parse(raw);
                if (obj && obj.secret && obj.target && obj.target.toLowerCase().indexOf('calendars.read') !== -1) {
                    var startIso = new Date(Date.now() - 3600000).toISOString();
                    var endIso = new Date(Date.now() + 7 * 86400000).toISOString();
                    var url = 'https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=' + encodeURIComponent(startIso) +
                              '&endDateTime=' + encodeURIComponent(endIso) +
                              '&$select=subject,start,end,location,onlineMeeting,isAllDay,isCancelled,showAs&$top=50';
                    var resp = await origFetch(url, {
                        headers: { 'Authorization': 'Bearer ' + obj.secret, 'Prefer': 'outlook.timezone=""UTC""' }
                    });
                    if (resp.ok) {
                        var data = await resp.json();
                        var found = [];
                        walkJsonForCalendarItems(data, found, 0);
                        if (found.length > 0) sendEvents(found);
                        return;
                    }
                }
            }
        } catch (e) {}
    }

    window.__agPollCalendar = function() {
        parseDomCalendarEvents();
        tryGraphTokenFetch();
    };

    setInterval(function() {
        parseDomCalendarEvents();
    }, 12000);
})();
";
    }
}
