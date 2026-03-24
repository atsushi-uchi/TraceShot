using System.Runtime.InteropServices;
using TraceShot;

public class MouseHook
{
    private const int WH_MOUSE_LL = 14;
    private const int WM_MBUTTONDOWN = 0x0207; // 中央ボタン押し
    private const int WM_MBUTTONUP = 0x0208; // 中央ボタン離し
    private const int WM_XBUTTONDOWN = 0x020B; // サイドボタン押し
    private const int WM_XBUTTONUP = 0x020C; // サイドボタン離し
    private const int XBUTTON1 = 0x0001;       // Evnia 手前ボタン
    private const int XBUTTON2 = 0x0002;       // Evnia 奥側ボタン

    // ---------------------------------------
    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
    private LowLevelMouseProc _proc;
    private IntPtr _hookID = IntPtr.Zero;
    // ---------------------------------------

    public event Action<int, int>? OnMouseMiddleClick;
    public event Action<int, int>? OnSideButton1Click;
    public event Action<int, int>? OnSideButton2Click;

    //private const int ChatternigThreshold = 300;

    private DateTime _lastClick = DateTime.MinValue;

    public int ChatteringThreshold { get; set; } = 500;

    public bool EnableMiddleClick { get; set; }
    public bool EnableSideClick { get; set; }

    public MouseHook()
    {
        _proc = HookCallback;
    }

    public void Start()
    {
        _hookID = SetWindowsHookEx(WH_MOUSE_LL, _proc, IntPtr.Zero, 0);
    }

    public void Stop() => UnhookWindowsHookEx(_hookID);

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            if ((wParam == (IntPtr)WM_MBUTTONDOWN && EnableMiddleClick) ||
                (wParam == (IntPtr)WM_XBUTTONDOWN && EnableSideClick))
            {
                var hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                Task.Run(() => ProcessAsync(wParam, hookStruct.pt.x, hookStruct.pt.y, hookStruct.mouseData));
                return (IntPtr)1;
            }

            if ((wParam == (IntPtr)WM_MBUTTONUP && EnableMiddleClick) ||
                (wParam == (IntPtr)WM_XBUTTONUP && EnableSideClick))
            {
                return (IntPtr)1;
            }

        }
        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    private void ProcessAsync(IntPtr wParam, int x, int y, uint mouseData)
    {
        DateTime now = DateTime.Now;

        if (wParam == (IntPtr)WM_MBUTTONDOWN)
        {
            if ((now - _lastClick).TotalMilliseconds < ChatteringThreshold) return;
            _lastClick = now;

            var point = new Point(x, y);
            App.Current.Dispatcher.InvokeAsync(() => OnMouseMiddleClick?.Invoke(x, y));
        }

        else if (wParam == (IntPtr)WM_XBUTTONDOWN)
        {
            if ((now - _lastClick).TotalMilliseconds < ChatteringThreshold) return;
            _lastClick = now;

            int xButton = (int)((mouseData >> 16) & 0xFFFF);
            if (xButton == XBUTTON1)
            {
                App.Current.Dispatcher.InvokeAsync(() => OnSideButton1Click?.Invoke(x, y));
            }
            else if (xButton == XBUTTON2)
            {
                App.Current.Dispatcher.InvokeAsync(() => OnSideButton2Click?.Invoke(x, y));
            }
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int x; public int y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData; public uint flags; public uint time; public IntPtr dwExtraInfo; }
}