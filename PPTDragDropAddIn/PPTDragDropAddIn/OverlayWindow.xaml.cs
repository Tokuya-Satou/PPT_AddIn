using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace PPTDragDropAddIn
{
    public partial class OverlayWindow : Window
    {
        private double _dpiScale = 1.0;
        private bool _isInDrawMode = false;

        // 描画時にパレット上のマウスイベントを除外するために使用
        public PenPaletteWindow AssociatedPalette { get; set; }

        /// <summary>
        /// TouchGuard の矩形がタッチダウンされたときに発火するイベント。
        /// 引数はスクリーン座標 (X, Y)。
        /// </summary>
        public event EventHandler<Point> TouchGuardTouched;
        /// <summary>タッチドラッグ中の移動（スクリーン座標）</summary>
        public event EventHandler<Point> TouchDragged;
        /// <summary>タッチドラッグ終了（スクリーン座標）</summary>
        public event EventHandler<Point> TouchDragEnded;

        public Action ImmediateBlockAction { get; set; }

        // タッチドラッグ中かどうか（TouchMove/TouchUp の処理対象を絞る）
        internal bool IsDraggingViaTouch = false;

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

            // ドラッグ中はウィンドウ全体を alpha=1 にしてタッチ/ジェスチャーを吸収し
            // PowerPoint のスワイプによるスライド遷移を防ぐ
            RootGrid.Background = NearlyTransparentBrush;

            // TouchGuardCanvas は HitTest を無効化するだけにとどめる。
            // Clear() は呼ばない。タッチイベントを処理中の矩形要素を visual tree から削除すると
            // WPF のタッチルーティングが壊れてフリーズするため。
            TouchGuardCanvas.IsHitTestVisible = false;
        }

        public void HideSnapshot()
        {
            DragImage.Visibility = Visibility.Collapsed;
            DragImage.Source = null;

            // ドラッグ終了 → タッチを通過させる（click-through に戻す）
            RootGrid.Background = Brushes.Transparent;
            TouchGuardCanvas.IsHitTestVisible = true;
        }

        public void UpdatePosition(double pixelX, double pixelY)
        {
            Canvas.SetLeft(DragImage, pixelX / _dpiScale);
            Canvas.SetTop(DragImage, pixelY / _dpiScale);
        }

        /// <summary>
        /// Drag_ 図形のウィンドウ内ピクセル座標矩形リストを受け取り、
        /// TouchGuardCanvas 上にタッチ吸収用の矩形を配置する。
        /// これにより、タッチダウンの最初のイベントがオーバーレイで捕捉され、
        /// PowerPoint の Direct Manipulation（スワイプでスライド遷移）が発動しない。
        /// </summary>
        public void UpdateTouchGuardRects(List<Rect> rects)
        {
            TouchGuardCanvas.Children.Clear();
            if (_isInDrawMode) return; // 描画モード中はガード不要

            foreach (var r in rects)
            {
                var rect = new Rectangle
                {
                    Width = r.Width / _dpiScale,
                    Height = r.Height / _dpiScale,
                    // alpha=1: 人間にはほぼ見えないが、OS のヒットテストで不透明と判定される
                    Fill = NearlyTransparentBrush,
                    IsHitTestVisible = true,
                };
                Canvas.SetLeft(rect, r.X / _dpiScale);
                Canvas.SetTop(rect, r.Y / _dpiScale);

                // タッチ・スタイラスイベントを捕捉
                rect.TouchDown += TouchGuardRect_TouchDown;

                TouchGuardCanvas.Children.Add(rect);
            }
        }

        /// <summary>
        /// TouchGuard 矩形を全て削除する。
        /// </summary>
        public void ClearTouchGuardRects()
        {
            TouchGuardCanvas.Children.Clear();
        }

        private void TouchGuardRect_TouchDown(object sender, TouchEventArgs e)
        {
            // ① ジェスチャーブロッカーを即座に有効化
            ImmediateBlockAction?.Invoke();

            // ② shape.Export（重いCOM処理）より先にオーバーレイを不透明化する。
            RootGrid.Background = NearlyTransparentBrush;
            TouchGuardCanvas.IsHitTestVisible = false;

            // ③ キャプチャを解放する。
            //    RootGrid.CaptureTouch() は TouchDown 処理中の StylusInput スレッドへ
            //    「キャプチャ先変更」の通知を送るため、UI スレッドと互いに待ち合い
            //    デッドロックする。キャプチャしなくても、背景が NearlyTransparentBrush に
            //    なった後は OS のヒットテストで WM_POINTER が RootGrid へルーティングされる。
            e.TouchDevice.Capture(null);

            IsDraggingViaTouch = true;

            var pos = e.GetTouchPoint(null).Position;
            var screenPos = PointToScreen(pos);
            TouchGuardTouched?.Invoke(this, screenPos);
        }

        private void RootGrid_TouchMove(object sender, TouchEventArgs e)
        {
            if (!IsDraggingViaTouch) return;
            e.Handled = true;
            var pos = e.GetTouchPoint(null).Position;
            TouchDragged?.Invoke(this, PointToScreen(pos));
        }

        private void RootGrid_TouchUp(object sender, TouchEventArgs e)
        {
            if (!IsDraggingViaTouch) return;
            e.Handled = true;
            IsDraggingViaTouch = false;
            RootGrid.ReleaseTouchCapture(e.TouchDevice);
            var pos = e.GetTouchPoint(null).Position;
            TouchDragEnded?.Invoke(this, PointToScreen(pos));
        }

        // AllowsTransparency=True のレイヤードウィンドウでは alpha=0 ピクセルが OS レベルで
        // click-through になるため、描画モード時は alpha=1 の背景を設定してイベントを受け取る
        private static readonly SolidColorBrush NearlyTransparentBrush =
            new SolidColorBrush(Color.FromArgb(1, 0, 0, 0));

        public void SetPenMode(Color color, double thickness = 3.0, bool isHighlighter = false)
        {
            _isInDrawMode = true;
            DrawingCanvas.DefaultDrawingAttributes = new DrawingAttributes
            {
                Color = color,
                Width = thickness,
                Height = thickness,
                FitToCurve = true,
                IsHighlighter = isHighlighter,
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
            RootGrid.Background = Brushes.Transparent; // ドラッグ/描画モードの残留をリセット
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
