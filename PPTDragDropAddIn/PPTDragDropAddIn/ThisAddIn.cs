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
        private GestureBlocker _gestureBlocker;

        // スライドごとの描画データを保持
        private Dictionary<int, System.Windows.Ink.StrokeCollection> _slideStrokes
            = new Dictionary<int, System.Windows.Ink.StrokeCollection>();
        private int _currentSlideIndex = -1;

        // Drag_ 図形のスクリーン座標矩形（WH_GETMESSAGE スレッドから COM なしで参照するために事前計算）
        private readonly object _dragRectsLock = new object();
        private List<System.Drawing.Rectangle> _dragShapeScreenRects = new List<System.Drawing.Rectangle>();

        // タッチドラッグ用に事前エクスポートした図形情報
        // タッチハンドラー内での COM 呼び出し（shape.Export 等）を排除するために使用
        private List<DragShapeInfo> _dragShapeInfos = new List<DragShapeInfo>();

        private class DragShapeInfo
        {
            public PowerPoint.Shape Shape;
            public string ExportedImagePath;
            public float PixelX, PixelY, PixelW, PixelH;
            public float SlideLeft, SlideTop;
            public System.Drawing.Rectangle ScreenRect;
        }

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

        // SetGestureConfig: スライドショーウィンドウのパン/スワイプジェスチャーを無効化
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetGestureConfig(IntPtr hwnd, uint dwReserved, uint cIDs,
            [In] GESTURECONFIG[] pGestureConfig, uint cbSize);

        [StructLayout(LayoutKind.Sequential)]
        private struct GESTURECONFIG
        {
            public uint dwID;
            public uint dwWant;
            public uint dwBlock;
        }

        private const uint GC_ALLGESTURES = 0x00000001;
        private const uint GID_PAN = 4;
        private const uint GC_PAN = 0x00000001;
        private const uint GC_PAN_WITH_SINGLE_FINGER_VERTICALLY = 0x00000002;
        private const uint GC_PAN_WITH_SINGLE_FINGER_HORIZONTALLY = 0x00000004;
        private const uint GC_PAN_WITH_GUTTER = 0x00000008;
        private const uint GC_PAN_WITH_INERTIA = 0x00000010;

        /// <summary>
        /// スライドショーウィンドウのパン（スワイプ）ジェスチャーを無効化する。
        /// これにより、タッチでのスライド遷移を OS レベルでブロックする。
        /// </summary>
        private void DisablePanGesture(IntPtr hwnd)
        {
            try
            {
                var configs = new GESTURECONFIG[]
                {
                    new GESTURECONFIG
                    {
                        dwID = GID_PAN,
                        dwWant = 0, // パンジェスチャーを一切受け付けない
                        dwBlock = GC_PAN_WITH_SINGLE_FINGER_VERTICALLY
                                | GC_PAN_WITH_SINGLE_FINGER_HORIZONTALLY
                                | GC_PAN_WITH_GUTTER
                                | GC_PAN_WITH_INERTIA
                    }
                };
                uint cbSize = (uint)Marshal.SizeOf(typeof(GESTURECONFIG));
                SetGestureConfig(hwnd, 0, 1, configs, cbSize);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("DisablePanGesture Error: " + ex.Message);
            }
        }

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
            bool dragEnabled = Properties.Settings.Default.DragDropEnabled;
            bool penEnabled  = Properties.Settings.Default.PenPaletteEnabled;

            if (!dragEnabled && !penEnabled) return;

            try
            {
                _slideStrokes.Clear();
                _currentSlideIndex = 1;
                _activeShowWindow = Wn;

                RECT rect;
                GetWindowRect((IntPtr)_activeShowWindow.HWND, out rect);
                _cachedWindowRect = rect;
                UpdateCoordinateContext(_activeShowWindow, rect);

                // ドラッグ機能が有効な場合のみ関連処理を起動
                if (dragEnabled)
                {
                    SaveInitialPositions(Wn.Presentation);
                    _mouseHook.Install();
                    DisablePanGesture((IntPtr)Wn.HWND);

                    _gestureBlocker?.Uninstall();
                    _gestureBlocker = new GestureBlocker();
                    _gestureBlocker.ShouldBlock = IsShapeAtScreenPoint;
                    _gestureBlocker.Install((IntPtr)Wn.HWND);

                    UpdateDragShapeInfos();

                    // タッチガードイベントの登録（※タッチでのドラッグにバグがあるため一時無効化）
                    // _overlayWindow.TouchGuardTouched -= OverlayWindow_TouchGuardTouched;
                    // _overlayWindow.TouchGuardTouched += OverlayWindow_TouchGuardTouched;
                    // _overlayWindow.TouchDragged -= OverlayWindow_TouchDragged;
                    // _overlayWindow.TouchDragged += OverlayWindow_TouchDragged;
                    // _overlayWindow.TouchDragEnded -= OverlayWindow_TouchDragEnded;
                    // _overlayWindow.TouchDragEnded += OverlayWindow_TouchDragEnded;
                    // _overlayWindow.ImmediateBlockAction = () => {
                    //    if (_gestureBlocker != null) _gestureBlocker.IsBlocking = true;
                    // };
                }

                // オーバーレイウィンドウはドラッグ・ペン描画どちらにも必要
                if (_overlayWindow == null)
                    _overlayWindow = new OverlayWindow();

                var helper = new System.Windows.Interop.WindowInteropHelper(_overlayWindow);
                IntPtr hwnd = helper.EnsureHandle();
                SetWindowPos(hwnd, HWND_TOPMOST,
                    rect.Left, rect.Top,
                    rect.Right - rect.Left, rect.Bottom - rect.Top,
                    SWP_SHOWWINDOW);

                _overlayWindow.Show();
                _isDrawModeActive = false;

                // 保存済みのペン色を DrawingAttributes に事前設定してからアロー表示
                _overlayWindow.Dispatcher.Invoke(() =>
                {
                    _overlayWindow.SetPenMode(GetSavedPenColor());
                    _overlayWindow.SetArrowMode();
                });

                // ペンパレットが有効な場合のみ表示
                if (penEnabled)
                {
                    if (_penPaletteWindow != null)
                    {
                        _penPaletteWindow.Close();
                        _penPaletteWindow = null;
                    }
                    _penPaletteWindow = new PenPaletteWindow();
                    _penPaletteWindow.Show();
                    _penPaletteWindow.UpdateLayout();
                    _overlayWindow.AssociatedPalette = _penPaletteWindow;

                    double dpiScale = 1.0;
                    try { dpiScale = System.Windows.Media.VisualTreeHelper.GetDpi(_penPaletteWindow).DpiScaleX; } catch { }
                    _penPaletteWindow.Left = (rect.Left + 10) / dpiScale;
                    _penPaletteWindow.Top = (rect.Bottom - 10) / dpiScale - _penPaletteWindow.ActualHeight;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SlideShowBegin Error: " + ex.Message);
            }
        }

        internal System.Windows.Media.Color GetSavedPenColor(int index = 1)
        {
            string colorName = index == 3 ? Properties.Settings.Default.Pen3Color :
                               (index == 2 ? Properties.Settings.Default.Pen2Color : Properties.Settings.Default.Pen1Color);
            switch (colorName)
            {
                case "Red":    return System.Windows.Media.Color.FromRgb(0xFF, 0x33, 0x33);
                case "Blue":   return System.Windows.Media.Color.FromRgb(0x33, 0x55, 0xFF);
                case "Green":  return System.Windows.Media.Color.FromRgb(0x33, 0xAA, 0x33);
                case "White":  return System.Windows.Media.Color.FromRgb(0xFF, 0xFF, 0xFF);
                case "Orange": return System.Windows.Media.Color.FromRgb(0xFF, 0x88, 0x00);
                case "Purple": return System.Windows.Media.Color.FromRgb(0x99, 0x33, 0xCC);
                default:       return System.Windows.Media.Color.FromRgb(0x11, 0x11, 0x11); // Black
            }
        }

        internal string GetSavedPenName(int index = 1)
        {
            string colorName = index == 3 ? Properties.Settings.Default.Pen3Color :
                               (index == 2 ? Properties.Settings.Default.Pen2Color : Properties.Settings.Default.Pen1Color);
            switch (colorName)
            {
                case "Red":    return "赤";
                case "Blue":   return "青";
                case "Green":  return "緑";
                case "White":  return "白";
                case "Orange": return "オレンジ";
                case "Purple": return "紫";
                default:       return "黒";
            }
        }

        internal System.Windows.Media.Color GetSavedMarkerColor()
        {
            switch (Properties.Settings.Default.Marker1Color)
            {
                case "Cyan":       return System.Windows.Media.Color.FromRgb(0x00, 0xCC, 0xFF);
                case "Orange":     return System.Windows.Media.Color.FromRgb(0xFF, 0x99, 0x00);
                case "LightGreen": return System.Windows.Media.Color.FromRgb(0x99, 0xFF, 0x33);
                case "Pink":       return System.Windows.Media.Color.FromRgb(0xFF, 0x66, 0xCC);
                default:           return System.Windows.Media.Color.FromRgb(0xFF, 0xFF, 0x00); // Yellow
            }
        }

        internal string GetSavedMarkerName()
        {
            switch (Properties.Settings.Default.Marker1Color)
            {
                case "Cyan":       return "水色";
                case "Orange":     return "橙";
                case "LightGreen": return "黄緑";
                case "Pink":       return "ピンク";
                default:           return "黄"; // Yellow
            }
        }

        private void Application_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            _mouseHook.Uninstall();
            _gestureBlocker?.Uninstall();
            _gestureBlocker = null;
            _isDrawModeActive = false;

            if (_overlayWindow != null)
            {
                _overlayWindow.Hide();
                _overlayWindow.Dispatcher.BeginInvoke(new Action(() =>
                {
                    _overlayWindow.HideSnapshot();
                    _overlayWindow.ClearDrawing();
                    _overlayWindow.ClearTouchGuardRects();
                }));
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
                                // 毎回ユニークなファイル名を使うことで BitmapImage の URI キャッシュ問題を回避
                                _tempImagePath = System.IO.Path.Combine(
                                    System.IO.Path.GetTempPath(),
                                    $"drag_temp_{System.Guid.NewGuid():N}.png");
                                shape.Export(_tempImagePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);

                                // スライド上の座標からウィンドウ内のピクセル座標を計算
                                float pixelX = (shape.Left / _slideWidth * _slidePixelWidth) + _offsetXInWindow;
                                float pixelY = (shape.Top / _slideHeight * _slidePixelHeight) + _offsetYInWindow;
                                float pixelW = (shape.Width / _slideWidth * _slidePixelWidth);
                                float pixelH = (shape.Height / _slideHeight * _slidePixelHeight);

                                // 元の図形を隠す（COM操作はここで）
                                shape.Visible = Office.MsoTriState.msoFalse;

                                // タッチジェスチャーによるスライド遷移をブロック
                                if (_gestureBlocker != null)
                                    _gestureBlocker.IsBlocking = true;

                                if (_overlayWindow != null)
                                {
                                    float px = pixelX, py = pixelY, pw = pixelW, ph = pixelH;
                                    string capturedPath = _tempImagePath;
                                    // WH_MOUSE_LL コールバックは別スレッドなので BeginInvoke（非同期）。
                                    // Invoke（同期）を使うとデッドロックする。
                                    _overlayWindow.Dispatcher.BeginInvoke(new Action(() =>
                                        _overlayWindow.ShowSnapshot(capturedPath, px, py, pw, ph)));
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

                // 使用済みテンポラリファイルを削除
                var pathToDelete = _tempImagePath;
                if (pathToDelete != null)
                    Task.Run(() => { try { System.IO.File.Delete(pathToDelete); } catch { } });
            }
            catch { }
            finally
            {
                // BeginInvoke（非同期）で HideSnapshot を呼ぶ。
                // WH_MOUSE_LL コールバック内で Invoke（同期）を使うとデッドロックする。
                if (_overlayWindow != null)
                    _overlayWindow.Dispatcher.BeginInvoke(new Action(() => _overlayWindow.HideSnapshot()));
            }

            // UpdateDragShapeInfos は必ず PPT STA スレッド（Dispatcher）で実行する。
            // Task.ContinueWith はスレッドプール（MTA）で動くため、そこで Shape COM 参照を
            // 取得すると MTA アパートメント経由のプロキシになり、後の STA 使用時にフリーズする。
            Task.Delay(100).ContinueWith(_ =>
                _overlayWindow?.Dispatcher.BeginInvoke(new Action(() =>
                {
                    try { UpdateDragShapeInfos(); } catch { }
                    ResetDragState();
                })));
        }

        private void ResetDragState()
        {
            _isDragging = false;
            _activeShape = null;
            if (_mouseHook != null) _mouseHook.IsDragging = false;
            if (_gestureBlocker != null) _gestureBlocker.IsBlocking = false;
            if (_overlayWindow != null) _overlayWindow.IsDraggingViaTouch = false;
        }

        /// <summary>
        /// 現在スライドの Drag_ 図形を事前エクスポートし、スクリーン座標・画像パスを保持する。
        /// タッチハンドラー内での COM 呼び出しを排除するために、操作前に必ず呼ぶこと。
        /// 必ず PPT STA スレッド（Dispatcher）上で呼ぶこと。
        /// スレッドプール（MTA）から呼ぶと Shape 参照が MTA アパートメント経由になり、
        /// 後で STA スレッドから使うときに COM マーシャリングが発生してフリーズする。
        /// </summary>
        private void UpdateDragShapeInfos()
        {
            // 旧 infos を退避（旧ファイル削除は新ファイル確定後に行う）
            List<DragShapeInfo> oldInfos;
            lock (_dragRectsLock) { oldInfos = _dragShapeInfos; }

            var infos = new List<DragShapeInfo>();
            try
            {
                if (_activeShowWindow != null && _slidePixelWidth > 0 && _slideHeight > 0)
                {
                    var winRect = _cachedWindowRect;
                    var shapes = _activeShowWindow.View.Slide.Shapes;
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        var shape = shapes[i];
                        if (!shape.Name.StartsWith("Drag_")) continue;

                        // 図形画像を事前エクスポート（タッチ時に行うと DM と競合してフリーズする）
                        string exportPath = System.IO.Path.Combine(
                            System.IO.Path.GetTempPath(),
                            $"drag_pre_{System.Guid.NewGuid():N}.png");
                        try { shape.Export(exportPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG); }
                        catch { continue; }

                        float pixelX = shape.Left / _slideWidth * _slidePixelWidth + _offsetXInWindow;
                        float pixelY = shape.Top  / _slideHeight * _slidePixelHeight + _offsetYInWindow;
                        float pixelW = shape.Width  / _slideWidth  * _slidePixelWidth;
                        float pixelH = shape.Height / _slideHeight * _slidePixelHeight;

                        int l = (int)pixelX + winRect.Left;
                        int t = (int)pixelY + winRect.Top;
                        int r = (int)(pixelX + pixelW) + winRect.Left;
                        int b = (int)(pixelY + pixelH) + winRect.Top;

                        infos.Add(new DragShapeInfo
                        {
                            Shape = shape,
                            ExportedImagePath = exportPath,
                            PixelX = pixelX, PixelY = pixelY, PixelW = pixelW, PixelH = pixelH,
                            SlideLeft = shape.Left, SlideTop = shape.Top,
                            ScreenRect = new System.Drawing.Rectangle(l, t, r - l, b - t)
                        });
                    }
                }
            }
            catch { }

            // 新ファイルが確定してから _dragShapeInfos を更新し、その後に旧ファイルを削除する。
            // 先に削除すると ShowSnapshot がファイルを読み込む前に消えて WIC エラーになる。
            var screenRects = infos.Select(info => info.ScreenRect).ToList();
            lock (_dragRectsLock)
            {
                _dragShapeInfos = infos;
                _dragShapeScreenRects = screenRects;
            }
            foreach (var old in oldInfos)
            {
                var p = old.ExportedImagePath;
                Task.Run(() => { try { System.IO.File.Delete(p); } catch { } });
            }

            // TouchGuard オーバーレイを更新（※タッチでのドラッグにバグがあるため一時無効化）
            /*
            if (_overlayWindow == null) return;
            var guardRects = new List<System.Windows.Rect>();
            foreach (var info in infos)
            {
                double x = info.ScreenRect.X - _cachedWindowRect.Left;
                double y = info.ScreenRect.Y - _cachedWindowRect.Top;
                guardRects.Add(new System.Windows.Rect(x, y, info.ScreenRect.Width, info.ScreenRect.Height));
            }
            _overlayWindow.UpdateTouchGuardRects(guardRects);
            */
        }

        /// <summary>
        /// GestureBlocker.ShouldBlock として使用。COM なし・スレッドセーフ。
        /// </summary>
        private bool IsShapeAtScreenPoint(int screenX, int screenY)
        {
            if (_isDrawModeActive) return false;
            lock (_dragRectsLock)
            {
                foreach (var r in _dragShapeScreenRects)
                    if (r.Contains(screenX, screenY)) return true;
            }
            return false;
        }

        /// <summary>
        /// TouchGuard 矩形がタッチされたときに呼ばれる。
        /// 事前エクスポート済みの _dragShapeInfos を使い、COM 呼び出しを一切行わない。
        /// （タッチハンドラー内で shape.Export 等の重い COM を呼ぶと、PowerPoint の
        ///   Direct Manipulation がスライド遷移を処理するタイミングと競合してフリーズする）
        /// </summary>
        private void OverlayWindow_TouchGuardTouched(object sender, System.Windows.Point screenPoint)
        {
            bool dragStarted = false;
            try
            {
                if (_activeShowWindow == null) return;
                if (_isDrawModeActive) return;

                int sx = (int)screenPoint.X;
                int sy = (int)screenPoint.Y;

                // 事前計算済みスクリーン矩形でヒットテスト（COM 呼び出しなし）
                DragShapeInfo found = null;
                lock (_dragRectsLock)
                {
                    // 高インデックス（最前面）優先で検索
                    for (int i = _dragShapeInfos.Count - 1; i >= 0; i--)
                    {
                        if (_dragShapeInfos[i].ScreenRect.Contains(sx, sy))
                        {
                            found = _dragShapeInfos[i];
                            break;
                        }
                    }
                }
                if (found == null) return;

                _activeShape = found.Shape;
                _isDragging = true;
                _mouseHook.IsDragging = true;
                _tempImagePath = null; // タッチドラッグでは事前エクスポート済みパスを直接使用

                float slideX = GetSlideX(sx);
                float slideY = GetSlideY(sy);
                _offsetX = slideX - found.SlideLeft;
                _offsetY = slideY - found.SlideTop;

                // 図形を非表示にする。重い処理は避けるため BeginInvoke で後回しにする。
                // オーバーレイが NearlyTransparentBrush で覆っているため、わずかな遅延は視覚的に問題ない。
                var shapeToHide = found.Shape;
                _overlayWindow.Dispatcher.BeginInvoke(
                    new Action(() => { try { shapeToHide.Visible = Office.MsoTriState.msoFalse; } catch { } }),
                    System.Windows.Threading.DispatcherPriority.Background);

                if (_gestureBlocker != null) _gestureBlocker.IsBlocking = true;

                if (_overlayWindow != null)
                {
                    float px = found.PixelX, py = found.PixelY, pw = found.PixelW, ph = found.PixelH;
                    string capturedPath = found.ExportedImagePath;
                    if (_overlayWindow.Dispatcher.CheckAccess())
                        _overlayWindow.ShowSnapshot(capturedPath, px, py, pw, ph);
                    else
                        _overlayWindow.Dispatcher.BeginInvoke(new Action(() =>
                            _overlayWindow.ShowSnapshot(capturedPath, px, py, pw, ph)));
                }

                dragStarted = true;
                _moveStopwatch.Restart();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("TouchGuardTouched Error: " + ex.Message);
                ResetDragState();
            }
            finally
            {
                if (!dragStarted && _overlayWindow != null)
                {
                    if (_overlayWindow.Dispatcher.CheckAccess())
                        _overlayWindow.HideSnapshot();
                    else
                        _overlayWindow.Dispatcher.BeginInvoke(new Action(() => _overlayWindow.HideSnapshot()));
                    if (_gestureBlocker != null)
                        _gestureBlocker.IsBlocking = false;
                }
            }
        }

        private void OverlayWindow_TouchDragged(object sender, System.Windows.Point screenPos)
        {
            if (!_isDragging || _overlayWindow == null) return;
            // TouchMove は WPF スレッド上で発火するため、UpdatePosition も直接呼べる
            float pixelX = (float)(screenPos.X - _cachedWindowRect.Left)
                           - (_offsetX / _slideWidth * _slidePixelWidth);
            float pixelY = (float)(screenPos.Y - _cachedWindowRect.Top)
                           - (_offsetY / _slideHeight * _slidePixelHeight);
            _overlayWindow.UpdatePosition(pixelX, pixelY);
        }

        private void OverlayWindow_TouchDragEnded(object sender, System.Windows.Point screenPos)
        {
            if (!_isDragging || _activeShape == null) return;
            try
            {
                float finalSlideX = GetSlideX((int)screenPos.X) - _offsetX;
                float finalSlideY = GetSlideY((int)screenPos.Y) - _offsetY;
                _activeShape.Left = finalSlideX;
                _activeShape.Top = finalSlideY;
                _activeShape.Visible = Office.MsoTriState.msoTrue;
                // 事前エクスポート済みファイルはここでは削除しない。
                // UpdateDragShapeInfos が次回呼ばれたとき（100ms 後）に旧ファイルを削除する。
            }
            catch { }
            finally
            {
                _overlayWindow?.HideSnapshot();
            }
            Task.Delay(100).ContinueWith(_ =>
                _overlayWindow?.Dispatcher.BeginInvoke(new Action(() =>
                {
                    try { UpdateDragShapeInfos(); } catch { }
                    ResetDragState();
                })));
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

            // インデックスをローカル変数にキャプチャ。
            // BeginInvoke（非同期）にするため、クロージャが _currentSlideIndex フィールドを
            // 参照すると先に更新されてしまう。ローカル変数で正しい値を保持する。
            int savedIndex = _currentSlideIndex;
            // Invoke（同期）は OverlayWindow_TouchGuardTouched 内の COM 呼び出しと
            // 同じスレッドでデッドロックを引き起こす可能性があるため BeginInvoke を使う。
            _overlayWindow.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (savedIndex > 0)
                    _slideStrokes[savedIndex] = _overlayWindow.GetStrokes();
                if (_slideStrokes.TryGetValue(newIndex, out var saved))
                    _overlayWindow.SetStrokes(saved);
                else
                    _overlayWindow.ClearDrawing();
            }));

            _currentSlideIndex = newIndex;
            UpdateDragShapeInfos(); // 新スライドの図形を事前エクスポート・TouchGuard矩形も更新

            try
            {
                if (_activeShowWindow != null)
                    DisablePanGesture((IntPtr)_activeShowWindow.HWND);
            }
            catch { }
        }

        public void SetPenMode(Color color, double thickness = 3.0, bool isHighlighter = false)
        {
            _isDrawModeActive = true;
            if (_overlayWindow != null)
                _overlayWindow.Dispatcher.Invoke(() => _overlayWindow.SetPenMode(color, thickness, isHighlighter));
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
