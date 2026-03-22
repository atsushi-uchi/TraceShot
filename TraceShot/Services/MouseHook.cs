using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

public class MouseHook
{
    private const int WH_MOUSE_LL = 14;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_LBUTTONUP = 0x0202;


    // --- 追加: チャタリング防止用フィールド ---
    private DateTime _lastClickTime = DateTime.MinValue;
    private TimeSpan _coolDown = TimeSpan.FromMilliseconds(500); // デフォルト500ms
    // ---------------------------------------

    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(POINT pt);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
    private LowLevelMouseProc? _proc;
    private IntPtr _hookID = IntPtr.Zero;

    public event Action<System.Drawing.Point>? OnLeftClick;

    // クールタイムを外部から調整したい場合
    public void SetCoolDown(int milliseconds) => _coolDown = TimeSpan.FromMilliseconds(milliseconds);

    public void Start()
    {
        _proc = HookCallback;
        _hookID = SetWindowsHookEx(WH_MOUSE_LL, _proc, IntPtr.Zero, 0);
    }

    public void Stop() => UnhookWindowsHookEx(_hookID);

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONUP)
        {
            // --- チャタリング防止ロジック ---
            DateTime now = DateTime.Now;
            if (now - _lastClickTime >= _coolDown)
            {
                MSLLHOOKSTRUCT hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

                if (IsOwnWindow(hookStruct.pt))
                {
                    Debug.WriteLine("MouseHook: Own window click ignored.");
                    return CallNextHookEx(_hookID, nCode, wParam, lParam);
                }

                _lastClickTime = now; // 実行時間を更新
                OnLeftClick?.Invoke(new System.Drawing.Point(hookStruct.pt.x, hookStruct.pt.y));
            }
            else
            {
                // クールタイム中のためスキップ
                Debug.WriteLine("MouseHook: Chattering blocked.");
            }
            // --------------------------------
        }
        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    private bool IsOwnWindow(POINT pt)
    {
        // クリックされた位置にあるウィンドウハンドルを取得
        IntPtr hWnd = WindowFromPoint(pt);
        if (hWnd == IntPtr.Zero) return false;

        // そのウィンドウのプロセスIDを取得
        GetWindowThreadProcessId(hWnd, out uint processId);

        // 現在実行中のプロセスIDと比較
        return processId == (uint)Environment.ProcessId;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int x; public int y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData; public uint flags; public uint time; public IntPtr dwExtraInfo; }
}