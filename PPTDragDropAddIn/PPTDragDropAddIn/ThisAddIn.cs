using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Media;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PPTDragDropAddIn
{
    public partial class ThisAddIn
    {
        private MouseHook _mouseHook;
        private PowerPoint.Shape _activeShape;
        private bool _isDragging = false;
        private float _offsetX;
        private float _offsetY;

        private OverlayWindow _overlayWindow;
        private PenPaletteWindow _penPaletteWindow;
        private bool _isDrawModeActive = false;

        // スライドごとの描画データを保持
        private Dictionary<int, System.Windows.Ink.StrokeCollection> _slideStrokes
            = new Dictionary<int, System.Windows.Ink.StrokeCollection>();
        private int _currentSlideIndex = -1;

        // パフォーマンス・精度改善用の変数
        private DateTime _lastMoveTime = DateTime.MinValue;
        private RECT _cachedWindowRect;
        private float _slidePixelWidth;
        private float _slidePixelHeight;
        private float _offsetXInWindow; // 黒帯のオフセット(X)
        private float _offsetYInWindow; // 黒帯のオフセット(Y)
        private const int MoveIntervalMs = 20; // 約50fps

        // 初期位置を保持するためのディクショナリ
        private Dictionary<string, (float Left, float Top)> _initialPositions = new Dictionary<string, (float Left, float Top)>();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_SHOWWINDOW = 0x0040;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private PowerPoint.SlideShowWindow _activeShowWindow;

        private System.Diagnostics.Stopwatch _moveStopwatch = new System.Diagnostics.Stopwatch();
        private const int MinMoveIntervalMs = 10; // 100fps 制御用

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SlideShowBegin += Application_SlideShowBegin;
            this.Application.SlideShowEnd += Application_SlideShowEnd;
            this.Application.SlideShowNextSlide += Application_SlideShowNextSlide;
            _mouseHook = new MouseHook();
            _mouseHook.MouseDown += MouseHook_MouseDown;
            _mouseHook.MouseMove += MouseHook_MouseMove;
            _mouseHook.MouseUp += MouseHook_MouseUp;
            _moveStopwatch.Start();
        }

        private void Application_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            try
            {
                // 初期位置を保存
                SaveInitialPositions(Wn.Presentation);
                _mouseHook.Install();

                // スライド描画データをリセット
                _slideStrokes.Clear();
                _currentSlideIndex = 1;

                // 投影用ウィンドウの特定
                _activeShowWindow = Wn;

                // オーバーレイウィンドウの作成と表示
                if (_overlayWindow == null)
                {
                    _overlayWindow = new OverlayWindow();
                }

                RECT rect;
                GetWindowRect((IntPtr)_activeShowWindow.HWND, out rect);
                _cachedWindowRect = rect;
                UpdateCoordinateContext(_activeShowWindow, rect);

                // Win32 API を使用して物理ピクセル単位で位置合わせ（DPIズレ防止）
                var helper = new System.Windows.Interop.WindowInteropHelper(_overlayWindow);
                IntPtr hwnd = helper.EnsureHandle();
                
                SetWindowPos(hwnd, HWND_TOPMOST, 
                    rect.Left, rect.Top, 
                    rect.Right - rect.Left, rect.Bottom - rect.Top, 
                    SWP_SHOWWINDOW);

                _overlayWindow.Show();
                _isDrawModeActive = false;
                _overlayWindow.Dispatcher.Invoke(() => _overlayWindow.SetArrowMode());

                // ペンパレットの作成と表示（毎回再作成して状態をリセット）
                if (_penPaletteWindow != null)
                {
                    _penPaletteWindow.Close();
                    _penPaletteWindow = null;
                }
                _penPaletteWindow = new PenPaletteWindow();
                _penPaletteWindow.Show();
                _penPaletteWindow.UpdateLayout();

                // OverlayWindow にパレット参照を渡して描画除外を有効化
                _overlayWindow.AssociatedPalette = _penPaletteWindow;

                double dpiScale = 1.0;
                try { dpiScale = System.Windows.Media.VisualTreeHelper.GetDpi(_penPaletteWindow).DpiScaleX; } catch { }
                _penPaletteWindow.Left = (rect.Left + 10) / dpiScale;
                _penPaletteWindow.Top = (rect.Bottom - 10) / dpiScale - _penPaletteWindow.ActualHeight;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SlideShowBegin Error: " + ex.Message);
            }
        }

        private void Application_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            _mouseHook.Uninstall();
            _isDrawModeActive = false;

            if (_overlayWindow != null)
            {
                _overlayWindow.Hide();
                _overlayWindow.Dispatcher.Invoke(() => _overlayWindow.ClearDrawing());
            }

            if (_penPaletteWindow != null)
                _penPaletteWindow.Hide();

            _slideStrokes.Clear();
            _currentSlideIndex = -1;

            ResetDragState();
            _activeShowWindow = null;

            // スライド終了時に元の位置に戻す
            RestoreInitialPositions(Pres);
        }

        private void SaveInitialPositions(PowerPoint.Presentation pres)
        {
            _initialPositions.Clear();
            foreach (PowerPoint.Slide slide in pres.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Name.StartsWith("Drag_"))
                    {
                        string key = $"{slide.SlideIndex}_{shape.Name}";
                        _initialPositions[key] = (shape.Left, shape.Top);
                    }
                }
            }
        }

        private void RestoreInitialPositions(PowerPoint.Presentation pres)
        {
            foreach (PowerPoint.Slide slide in pres.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    string key = $"{slide.SlideIndex}_{shape.Name}";
                    if (_initialPositions.ContainsKey(key))
                    {
                        var pos = _initialPositions[key];
                        shape.Left = pos.Left;
                        shape.Top = pos.Top;
                    }
                }
            }
        }

        private string _tempImagePath;

        private void MouseHook_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            try
            {
                if (_activeShowWindow == null) return;

                RECT rect = _cachedWindowRect;
                // 最新のウィンドウ位置をチェック（移動している可能性があるため）
                GetWindowRect((IntPtr)_activeShowWindow.HWND, out rect);
                _cachedWindowRect = rect;

                // パレット上のクリックはドラッグ検知をスキップ
                if (_penPaletteWindow != null && _penPaletteWindow.IsVisible)
                {
                    RECT paletteRect;
                    var paletteHelper = new System.Windows.Interop.WindowInteropHelper(_penPaletteWindow);
                    if (GetWindowRect(paletteHelper.Handle, out paletteRect))
                    {
                        if (e.X >= paletteRect.Left && e.X <= paletteRect.Right &&
                            e.Y >= paletteRect.Top && e.Y <= paletteRect.Bottom)
                            return;
                    }
                }

                if (!_isDrawModeActive && e.X >= rect.Left && e.X <= rect.Right && e.Y >= rect.Top && e.Y <= rect.Bottom)
                {
                    UpdateCoordinateContext(_activeShowWindow, rect);

                    float slideX = GetSlideX(e.X);
                    float slideY = GetSlideY(e.Y);

                    int shapeCount = _activeShowWindow.View.Slide.Shapes.Count;
                    for (int i = shapeCount; i >= 1; i--)
                    {
                        PowerPoint.Shape shape = _activeShowWindow.View.Slide.Shapes[i];
                        if (slideX >= shape.Left && slideX <= (shape.Left + shape.Width) &&
                            slideY >= shape.Top && slideY <= (shape.Top + shape.Height))
                        {
                            // 最前面の図形が「Drag_」で始まる場合のみドラッグを開始する
                            if (shape.Name.StartsWith("Drag_"))
                            {
                                _activeShape = shape;
                                _isDragging = true;
                                _mouseHook.IsDragging = true;
                                _offsetX = slideX - shape.Left;
                                _offsetY = slideY - shape.Top;

                                // 画像キャプチャと WPF 表示
                                _tempImagePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "drag_temp.png");
                                shape.Export(_tempImagePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);

                                // スライド上の座標からウィンドウ内のピクセル座標を計算
                                float pixelX = (shape.Left / _slideWidth * _slidePixelWidth) + _offsetXInWindow;
                                float pixelY = (shape.Top / _slideHeight * _slidePixelHeight) + _offsetYInWindow;
                                float pixelW = (shape.Width / _slideWidth * _slidePixelWidth);
                                float pixelH = (shape.Height / _slideHeight * _slidePixelHeight);

                                // 元の図形を隠す（COM操作はここで）
                                shape.Visible = Office.MsoTriState.msoFalse;

                                if (_overlayWindow != null)
                                {
                                    float px = pixelX, py = pixelY, pw = pixelW, ph = pixelH;
                                    _overlayWindow.Dispatcher.Invoke(() =>
                                        _overlayWindow.ShowSnapshot(_tempImagePath, px, py, pw, ph));
                                }

                                _moveStopwatch.Restart();
                                return;
                            }
                            else
                            {
                                // 最前面の図形がドラッグ対象でない場合は、下の図形を探さずに終了（通常のクリック操作を優先）
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MouseDown Error: " + ex.Message);
                ResetDragState();
            }
        }

        private float _slideWidth;
        private float _slideHeight;

        private void UpdateCoordinateContext(PowerPoint.SlideShowWindow sw, RECT rect)
        {
            try
            {
                float winW = rect.Right - rect.Left;
                float winH = rect.Bottom - rect.Top;
                if (winW <= 0 || winH <= 0) return;

                _slideWidth = (float)this.Application.ActivePresentation.PageSetup.SlideWidth;
                _slideHeight = (float)this.Application.ActivePresentation.PageSetup.SlideHeight;

                float ratioW = winW / _slideWidth;
                float ratioH = winH / _slideHeight;

                float scale = Math.Min(ratioW, ratioH);
                _slidePixelWidth = _slideWidth * scale;
                _slidePixelHeight = _slideHeight * scale;
                _offsetXInWindow = (winW - _slidePixelWidth) / 2;
                _offsetYInWindow = (winH - _slidePixelHeight) / 2;
            }
            catch { }
        }

        private void MouseHook_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (!_isDragging || _activeShape == null || _overlayWindow == null) return;

            try
            {
                // ここでは COM 通信をせず、WPF の更新のみを行う
                // マウス位置から新位置（ピクセル）を計算
                float pixelX = e.X - _cachedWindowRect.Left - (_offsetX / _slideWidth * _slidePixelWidth);
                float pixelY = e.Y - _cachedWindowRect.Top - (_offsetY / _slideHeight * _slidePixelHeight);

                _overlayWindow.Dispatcher.BeginInvoke(new Action(() =>
                    _overlayWindow.UpdatePosition(pixelX, pixelY)));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MouseMove Error: " + ex.Message);
            }
        }

        private void MouseHook_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (!_isDragging || _activeShape == null) return;

            try
            {
                // 最終位置の確定（ここで初めて PowerPoint を更新）
                float finalSlideX = GetSlideX(e.X) - _offsetX;
                float finalSlideY = GetSlideY(e.Y) - _offsetY;

                _activeShape.Left = finalSlideX;
                _activeShape.Top = finalSlideY;
                _activeShape.Visible = Office.MsoTriState.msoTrue;

                if (_overlayWindow != null)
                    _overlayWindow.Dispatcher.Invoke(() => _overlayWindow.HideSnapshot());
            }
            catch { }

            Task.Delay(50).ContinueWith(_ => {
                ResetDragState();
            });
        }

        private void ResetDragState()
        {
            _isDragging = false;
            _activeShape = null;
            if (_mouseHook != null) _mouseHook.IsDragging = false;
        }

        private float GetSlideX(int screenX)
        {
            if (_slidePixelWidth <= 0) return 0;
            // キャッシュした _slideWidth を使用（COM通信なし）
            return (screenX - _cachedWindowRect.Left - _offsetXInWindow) / _slidePixelWidth * _slideWidth;
        }

        private float GetSlideY(int screenY)
        {
            if (_slidePixelHeight <= 0) return 0;
            // キャッシュした _slideHeight を使用（COM通信なし）
            return (screenY - _cachedWindowRect.Top - _offsetYInWindow) / _slidePixelHeight * _slideHeight;
        }

        private void Application_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            if (_overlayWindow == null) return;

            int newIndex;
            try { newIndex = Wn.View.CurrentShowPosition; }
            catch { return; }

            _overlayWindow.Dispatcher.Invoke(() =>
            {
                // 現在のスライドの描画を保存
                if (_currentSlideIndex > 0)
                    _slideStrokes[_currentSlideIndex] = _overlayWindow.GetStrokes();

                // 新しいスライドの描画を復元（なければ空）
                if (_slideStrokes.TryGetValue(newIndex, out var saved))
                    _overlayWindow.SetStrokes(saved);
                else
                    _overlayWindow.ClearDrawing();
            });

            _currentSlideIndex = newIndex;
        }

        public void SetPenMode(Color color, double thickness = 3.0)
        {
            _isDrawModeActive = true;
            if (_overlayWindow != null)
                _overlayWindow.Dispatcher.Invoke(() => _overlayWindow.SetPenMode(color, thickness));
        }

        public void SetEraserMode()
        {
            _isDrawModeActive = true;
            if (_overlayWindow != null)
                _overlayWindow.Dispatcher.Invoke(() => _overlayWindow.SetEraserMode());
        }

        public void SetArrowMode()
        {
            _isDrawModeActive = false;
            if (_overlayWindow != null)
                _overlayWindow.Dispatcher.Invoke(() => _overlayWindow.SetArrowMode());
        }

        public void GoNextSlide()
        {
            try
            {
                var sw = _activeShowWindow;
                if (sw == null) return;
                // COM呼び出しをVSTOのSTA(UIスレッド)で実行
                System.Windows.Forms.Application.DoEvents();
                sw.View.Next();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GoNextSlide Error: " + ex.Message);
            }
        }

        public void GoPrevSlide()
        {
            try
            {
                var sw = _activeShowWindow;
                if (sw == null) return;
                System.Windows.Forms.Application.DoEvents();
                sw.View.Previous();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GoPrevSlide Error: " + ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _mouseHook?.Uninstall();
        }

        #region VSTO で生成されたコード
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new DragDropRibbon();
        }
    }
}
