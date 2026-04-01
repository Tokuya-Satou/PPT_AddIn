using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PPTDragDropAddIn
{
    public partial class OverlayWindow : Window
    {
        private double _dpiScale = 1.0;
        private bool _isInDrawMode = false;

        // 描画時にパレット上のマウスイベントを除外するために使用
        public PenPaletteWindow AssociatedPalette { get; set; }

        private const int WM_NCHITTEST = 0x0084;
        private const int HTTRANSPARENT = -1;

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        public OverlayWindow()
        {
            InitializeComponent();
            this.Loaded += (s, e) => {
                _dpiScale = VisualTreeHelper.GetDpi(this).DpiScaleX;
            };
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            // WM_NCHITTEST をフックしてパレット領域を click-through にする
            var source = HwndSource.FromHwnd(new WindowInteropHelper(this).Handle);
            source.AddHook(WndProc);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_NCHITTEST && _isInDrawMode)
            {
                // lParam の下位16bit=X, 上位16bit=Y (スクリーン座標)
                int lp = lParam.ToInt32();
                int x = (short)(lp & 0xFFFF);
                int y = (short)((lp >> 16) & 0xFFFF);

                if (IsScreenPointOverPalette(x, y))
                {
                    // HTTRANSPARENT = OS がこのウィンドウを無視して下のウィンドウにイベントを渡す
                    handled = true;
                    return new IntPtr(HTTRANSPARENT);
                }
            }
            return IntPtr.Zero;
        }

        // スクリーン座標でパレット領域か判定（WndProc 用）
        private bool IsScreenPointOverPalette(int screenX, int screenY)
        {
            if (AssociatedPalette == null || !AssociatedPalette.IsVisible) return false;
            var helper = new WindowInteropHelper(AssociatedPalette);
            RECT r;
            if (!GetWindowRect(helper.Handle, out r)) return false;
            return screenX >= r.Left && screenX <= r.Right &&
                   screenY >= r.Top  && screenY <= r.Bottom;
        }

        protected override void OnDpiChanged(DpiScale oldDpi, DpiScale newDpi)
        {
            base.OnDpiChanged(oldDpi, newDpi);
            _dpiScale = newDpi.DpiScaleX;
        }

        public void ShowSnapshot(string imagePath, double pixelX, double pixelY, double pixelW, double pixelH)
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.UriSource = new Uri(imagePath, UriKind.Absolute);
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.EndInit();

            DragImage.Source = bitmap;
            DragImage.Width = pixelW / _dpiScale;
            DragImage.Height = pixelH / _dpiScale;
            Canvas.SetLeft(DragImage, pixelX / _dpiScale);
            Canvas.SetTop(DragImage, pixelY / _dpiScale);
            DragImage.Visibility = Visibility.Visible;
        }

        public void HideSnapshot()
        {
            DragImage.Visibility = Visibility.Collapsed;
            DragImage.Source = null;
        }

        public void UpdatePosition(double pixelX, double pixelY)
        {
            Canvas.SetLeft(DragImage, pixelX / _dpiScale);
            Canvas.SetTop(DragImage, pixelY / _dpiScale);
        }

        // AllowsTransparency=True のレイヤードウィンドウでは alpha=0 ピクセルが OS レベルで
        // click-through になるため、描画モード時は alpha=1 の背景を設定してイベントを受け取る
        private static readonly SolidColorBrush NearlyTransparentBrush =
            new SolidColorBrush(Color.FromArgb(1, 0, 0, 0));

        public void SetPenMode(Color color, double thickness = 3.0)
        {
            _isInDrawMode = true;
            DrawingCanvas.DefaultDrawingAttributes = new DrawingAttributes
            {
                Color = color,
                Width = thickness,
                Height = thickness,
                FitToCurve = true,
                IsHighlighter = false,
            };
            DrawingCanvas.EditingMode = InkCanvasEditingMode.Ink;
            DrawingCanvas.Background = NearlyTransparentBrush;
            DrawingCanvas.IsHitTestVisible = true;
        }

        public void SetEraserMode()
        {
            _isInDrawMode = true;
            DrawingCanvas.EditingMode = InkCanvasEditingMode.EraseByStroke;
            DrawingCanvas.Background = NearlyTransparentBrush;
            DrawingCanvas.IsHitTestVisible = true;
        }

        public void SetArrowMode()
        {
            _isInDrawMode = false;
            DrawingCanvas.EditingMode = InkCanvasEditingMode.None;
            DrawingCanvas.Background = Brushes.Transparent;
            DrawingCanvas.IsHitTestVisible = false;
        }

        public void ClearDrawing()
        {
            DrawingCanvas.Strokes.Clear();
        }

        public StrokeCollection GetStrokes()
        {
            return DrawingCanvas.Strokes.Clone();
        }

        public void SetStrokes(StrokeCollection strokes)
        {
            DrawingCanvas.Strokes.Clear();
            if (strokes != null)
                foreach (var stroke in strokes)
                    DrawingCanvas.Strokes.Add(stroke);
        }
    }
}
