using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace TeamsNextMeetingTile.Services;

public struct WinRect
{
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;

    public int Width => Right - Left;
    public int Height => Bottom - Top;

    public bool Intersects(WinRect other)
    {
        return Left < other.Right &&
               Right > other.Left &&
               Top < other.Bottom &&
               Bottom > other.Top;
    }
}

public class TeamsWindowState
{
    public bool IsTeamsRunning { get; set; }
    public bool IsTeamsMinimized { get; set; }
    public bool IsTileAreaOccluded { get; set; }
    public IntPtr TeamsHandle { get; set; }
    public WinRect TeamsBounds { get; set; }
    public double DpiScale { get; set; } = 1.0;
}

public static class TeamsWindowTracker
{
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out WinRect lpRect);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("dwmapi.dll")]
    private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out WinRect pvAttribute, int cbAttribute);

    [DllImport("dwmapi.dll")]
    private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out int pvAttribute, int cbAttribute);

    private const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;
    private const int DWMWA_CLOAKED = 14;
    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_TOOLWINDOW = 0x00000080;
    private const int WS_EX_TRANSPARENT = 0x00000020;

    private static readonly int CurrentPid = Environment.ProcessId;

    public static TeamsWindowState GetTeamsState(IntPtr ownOverlayHandle, IntPtr ownSettingsHandle, IntPtr ownSyncHandle, double logicalOffsetX, double logicalOffsetY, double logicalWidth, double logicalHeight)
    {
        var state = new TeamsWindowState();
        IntPtr bestHandle = IntPtr.Zero;
        WinRect bestRect = default;
        int bestScore = -1;
        bool foundMinimizedTeams = false;

        EnumWindows((hWnd, _) =>
        {
            if (hWnd == ownOverlayHandle || hWnd == ownSettingsHandle || hWnd == ownSyncHandle)
                return true;

            if (!IsWindowVisible(hWnd))
                return true;

            if (IsWindowCloaked(hWnd))
                return true;

            GetWindowThreadProcessId(hWnd, out uint pid);
            if (pid == CurrentPid || pid == 0)
                return true;

            string title = GetWindowTitle(hWnd);
            string className = GetWindowClassName(hWnd);

            bool isTeamsCandidate = false;
            if (title.Contains("Microsoft Teams", StringComparison.OrdinalIgnoreCase))
            {
                isTeamsCandidate = true;
            }
            else if (className.Contains("TeamsWebView", StringComparison.OrdinalIgnoreCase) ||
                     className.Contains("WinUIDesktopWin32WindowClass", StringComparison.OrdinalIgnoreCase))
            {
                string procName = SafeGetProcessName((int)pid);
                if (procName.Equals("ms-teams", StringComparison.OrdinalIgnoreCase) ||
                    procName.Equals("Teams", StringComparison.OrdinalIgnoreCase))
                {
                    isTeamsCandidate = true;
                }
            }

            if (!isTeamsCandidate)
                return true;

            string pName = SafeGetProcessName((int)pid);
            bool isNativeTeamsProcess = pName.Equals("ms-teams", StringComparison.OrdinalIgnoreCase) ||
                                        pName.Equals("Teams", StringComparison.OrdinalIgnoreCase);
            bool isEdgePwa = pName.Equals("msedge", StringComparison.OrdinalIgnoreCase) &&
                             title.EndsWith("Microsoft Teams", StringComparison.OrdinalIgnoreCase);

            if (!isNativeTeamsProcess && !isEdgePwa)
                return true;

            if (IsIconic(hWnd))
            {
                foundMinimizedTeams = true;
                return true;
            }

            WinRect rect = GetAccurateWindowRect(hWnd);
            if (rect.Width < 380 || rect.Height < 280)
                return true;

            int score = (rect.Width * rect.Height) / 1000;
            if (isNativeTeamsProcess) score += 5000;
            if (title.StartsWith("Chat |", StringComparison.OrdinalIgnoreCase) ||
                title.StartsWith("Activity |", StringComparison.OrdinalIgnoreCase) ||
                title.StartsWith("Calendar |", StringComparison.OrdinalIgnoreCase) ||
                title.StartsWith("Teams |", StringComparison.OrdinalIgnoreCase))
            {
                score += 10000;
            }

            if (score > bestScore)
            {
                bestScore = score;
                bestHandle = hWnd;
                bestRect = rect;
            }

            return true;
        }, IntPtr.Zero);

        if (bestHandle == IntPtr.Zero)
        {
            state.IsTeamsRunning = foundMinimizedTeams;
            state.IsTeamsMinimized = foundMinimizedTeams;
            return state;
        }

        state.IsTeamsRunning = true;
        state.IsTeamsMinimized = false;
        state.TeamsHandle = bestHandle;
        state.TeamsBounds = bestRect;

        uint dpi = GetDpiForWindow(bestHandle);
        double scale = dpi > 0 ? dpi / 96.0 : 1.0;
        state.DpiScale = scale;

        var targetTileRect = new WinRect
        {
            Left = bestRect.Left + (int)(logicalOffsetX * scale),
            Top = bestRect.Top + (int)(logicalOffsetY * scale),
            Right = bestRect.Left + (int)((logicalOffsetX + logicalWidth) * scale),
            Bottom = bestRect.Top + (int)((logicalOffsetY + logicalHeight) * scale)
        };

        state.IsTileAreaOccluded = CheckIfRectOccludedBeforeWindow(
            bestHandle,
            targetTileRect,
            ownOverlayHandle,
            ownSyncHandle);

        return state;
    }

    private static bool CheckIfRectOccludedBeforeWindow(
        IntPtr targetWindow,
        WinRect tileRect,
        IntPtr ownOverlayHandle,
        IntPtr ownSyncHandle)
    {
        bool occluded = false;

        EnumWindows((hWnd, _) =>
        {
            if (hWnd == targetWindow)
            {
                // Reached the Teams window in Z-order without hitting an occluding window above it
                return false;
            }

            if (hWnd == ownOverlayHandle || (ownSyncHandle != IntPtr.Zero && hWnd == ownSyncHandle))
                return true;

            if (!IsWindowVisible(hWnd) || IsIconic(hWnd) || IsWindowCloaked(hWnd))
                return true;

            int exStyle = GetWindowLong(hWnd, GWL_EXSTYLE);
            if ((exStyle & WS_EX_TRANSPARENT) != 0)
                return true;

            string className = GetWindowClassName(hWnd);
            if (className is "Progman" or "WorkerW" or "Shell_TrayWnd" or "Shell_SecondaryTrayWnd"
                or "ime" or "MSCTFIME UI" or "DummyDWMListenerWindow" or "EdgeUiInputTopWndClass"
                or "Windows.UI.Core.CoreWindow" or "XamlExplorerHostIslandWindow" or "TopLevelWindowForOverflowXamlIsland")
                return true;

            WinRect r = GetAccurateWindowRect(hWnd);
            if (r.Width <= 20 || r.Height <= 20)
                return true;

            if (r.Intersects(tileRect))
            {
                occluded = true;
                return false;
            }

            return true;
        }, IntPtr.Zero);

        return occluded;
    }

    private static WinRect GetAccurateWindowRect(IntPtr hWnd)
    {
        if (DwmGetWindowAttribute(hWnd, DWMWA_EXTENDED_FRAME_BOUNDS, out WinRect dwmRect, Marshal.SizeOf<WinRect>()) == 0)
        {
            if (dwmRect.Width > 0 && dwmRect.Height > 0)
                return dwmRect;
        }
        GetWindowRect(hWnd, out WinRect rect);
        return rect;
    }

    private static bool IsWindowCloaked(IntPtr hWnd)
    {
        if (DwmGetWindowAttribute(hWnd, DWMWA_CLOAKED, out int cloaked, sizeof(int)) == 0)
        {
            return cloaked != 0;
        }
        return false;
    }

    private static string GetWindowTitle(IntPtr hWnd)
    {
        var sb = new StringBuilder(256);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private static string GetWindowClassName(IntPtr hWnd)
    {
        var sb = new StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private static string SafeGetProcessName(int pid)
    {
        try
        {
            using var proc = Process.GetProcessById(pid);
            return proc.ProcessName ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }
}
