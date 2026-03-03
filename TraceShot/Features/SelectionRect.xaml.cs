using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using WpfPoint = System.Windows.Point;
using System.Windows.Forms;

namespace TraceShot
{
    public partial class SelectionRect : Window
    {
        #region Win32 API Definitions

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromPoint(POINT pt, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

        private const uint MONITOR_DEFAULTTONEAREST = 2;

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT { public int X; public int Y; }

        [StructLayout(LayoutKind.Sequential)]
        public struct MONITORINFO
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        #endregion

        public Rect SelectedRegion { get; private set; }
        public string TargetDeviceName { get; private set; } // 選択されたモニターの情報

        public SelectionRect()
        {
            InitializeComponent();

            // 全モニターをカバー
            this.Left = SystemParameters.VirtualScreenLeft;
            this.Top = SystemParameters.VirtualScreenTop;
            this.Width = SystemParameters.VirtualScreenWidth;
            this.Height = SystemParameters.VirtualScreenHeight;
        }

        private WpfPoint _startLogicalPos;
        private POINT _startPhysicalPos;

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.CaptureMouse();
            GetCursorPos(out _startPhysicalPos);
            _startLogicalPos = e.GetPosition(this);

            SelectionRectangle.Visibility = Visibility.Visible;
            Canvas.SetLeft(SelectionRectangle, _startLogicalPos.X);
            Canvas.SetTop(SelectionRectangle, _startLogicalPos.Y);
            SelectionRectangle.Width = 0;
            SelectionRectangle.Height = 0;
        }

        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            if (!this.IsMouseCaptured) return;

            WpfPoint currentLogicalPos = e.GetPosition(this);
            double x = Math.Min(_startLogicalPos.X, currentLogicalPos.X);
            double y = Math.Min(_startLogicalPos.Y, currentLogicalPos.Y);
            double w = Math.Abs(_startLogicalPos.X - currentLogicalPos.X);
            double h = Math.Abs(_startLogicalPos.Y - currentLogicalPos.Y);

            Canvas.SetLeft(SelectionRectangle, x);
            Canvas.SetTop(SelectionRectangle, y);
            SelectionRectangle.Width = w;
            SelectionRectangle.Height = h;
        }

        private void Window_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (this.IsMouseCaptured) this.ReleaseMouseCapture();

            POINT endPhysicalPos;
            GetCursorPos(out endPhysicalPos);

            // 1. 仮想画面全体での絶対座標(物理ピクセル)を計算
            int absX = Math.Min(_startPhysicalPos.X, endPhysicalPos.X);
            int absY = Math.Min(_startPhysicalPos.Y, endPhysicalPos.Y);
            int w = Math.Abs(_startPhysicalPos.X - endPhysicalPos.X);
            int h = Math.Abs(_startPhysicalPos.Y - endPhysicalPos.Y);

            // 2. マウス開始地点の物理座標から Screen オブジェクトを取得
            // System.Drawing.Point に変換して Screen.FromPoint を呼ぶのが一番簡単です
            var drawingPoint = new System.Drawing.Point(_startPhysicalPos.X, _startPhysicalPos.Y);
            Screen targetScreen = Screen.FromPoint(drawingPoint);

            if (targetScreen != null)
            {
                // 3. そのモニターの左上を原点(0,0)とした相対座標に変換
                // targetScreen.Bounds は仮想画面上でのそのモニターの範囲(物理ピクセル)
                int relativeX = absX - targetScreen.Bounds.X;
                int relativeY = absY - targetScreen.Bounds.Y;

                // 4. 偶数補正
                w = (w / 2) * 2;
                h = (h / 2) * 2;

                // 結果を格納
                this.SelectedRegion = new Rect(relativeX, relativeY, w, h);

                // デバイス名 (例: "\\.\DISPLAY1") を取得
                TargetDeviceName = targetScreen.DeviceName;

                Console.WriteLine($"Screen: {TargetDeviceName}");
                Console.WriteLine($"Relative Pos: X:{relativeX}, Y:{relativeY}, W:{w}, H:{h}");
            }

            FinishSelection();
        }

        private void FinishSelection()
        {
            this.DialogResult = true;
            this.Close();
        }
    }
}