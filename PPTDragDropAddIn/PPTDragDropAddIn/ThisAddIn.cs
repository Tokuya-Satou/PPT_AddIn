using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SlideShowBegin += Application_SlideShowBegin;
            this.Application.SlideShowEnd += Application_SlideShowEnd;
            _mouseHook = new MouseHook();
            _mouseHook.MouseDown += MouseHook_MouseDown;
            _mouseHook.MouseMove += MouseHook_MouseMove;
            _mouseHook.MouseUp += MouseHook_MouseUp;
        }

        private void Application_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            // 初期位置を保存
            SaveInitialPositions(Wn.Presentation);
            _mouseHook.Install();
        }

        private void Application_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            _mouseHook.Uninstall();
            _isDragging = false;
            _activeShape = null;
            if (_mouseHook != null) _mouseHook.IsDragging = false;

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

        private void MouseHook_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (this.Application.SlideShowWindows.Count == 0) return;

            var sw = this.Application.SlideShowWindows[1];
            var view = sw.View;
            
            // ウィンドウ座標を取得
            GetWindowRect((IntPtr)sw.HWND, out _cachedWindowRect);
            float winW = _cachedWindowRect.Right - _cachedWindowRect.Left;
            float winH = _cachedWindowRect.Bottom - _cachedWindowRect.Top;
            if (winW <= 0 || winH <= 0) return;

            // スライドとウィンドウのアスペクト比を比較して「実際の表示領域」を計算
            float slideW = (float)this.Application.ActivePresentation.PageSetup.SlideWidth;
            float slideH = (float)this.Application.ActivePresentation.PageSetup.SlideHeight;
            float ratioW = winW / slideW;
            float ratioH = winH / slideH;

            // 小さい方の比率に合わせてスライドが表示される（黒帯ができる）
            float scale = Math.Min(ratioW, ratioH);
            _slidePixelWidth = slideW * scale;
            _slidePixelHeight = slideH * scale;
            _offsetXInWindow = (winW - _slidePixelWidth) / 2;
            _offsetYInWindow = (winH - _slidePixelHeight) / 2;

            float slideX = GetSlideX(e.X);
            float slideY = GetSlideY(e.Y);

            foreach (PowerPoint.Shape shape in view.Slide.Shapes)
            {
                if (shape.Name.StartsWith("Drag_") &&
                    slideX >= shape.Left && slideX <= (shape.Left + shape.Width) &&
                    slideY >= shape.Top && slideY <= (shape.Top + shape.Height))
                {
                    _activeShape = shape;
                    _isDragging = true;
                    _mouseHook.IsDragging = true;
                    _offsetX = slideX - shape.Left;
                    _offsetY = slideY - shape.Top;
                    _lastMoveTime = DateTime.Now;
                    break;
                }
            }
        }

        private void MouseHook_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (_isDragging && _activeShape != null)
            {
                if ((DateTime.Now - _lastMoveTime).TotalMilliseconds < MoveIntervalMs)
                    return;

                _activeShape.Left = GetSlideX(e.X) - _offsetX;
                _activeShape.Top = GetSlideY(e.Y) - _offsetY;
                
                _lastMoveTime = DateTime.Now;
            }
        }

        private void MouseHook_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            Task.Delay(50).ContinueWith(_ => {
                _isDragging = false;
                _activeShape = null;
                if (_mouseHook != null) _mouseHook.IsDragging = false;
            });
        }

        private float GetSlideX(int screenX)
        {
            if (_slidePixelWidth <= 0) return 0;
            float slideW = (float)this.Application.ActivePresentation.PageSetup.SlideWidth;
            // (マウスクリックピクセル - ウィンドウ左端 - 黒帯幅) / 表示幅 * スライドポイント幅
            return (screenX - _cachedWindowRect.Left - _offsetXInWindow) / _slidePixelWidth * slideW;
        }

        private float GetSlideY(int screenY)
        {
            if (_slidePixelHeight <= 0) return 0;
            float slideH = (float)this.Application.ActivePresentation.PageSetup.SlideHeight;
            // (マウスクリックピクセル - ウィンドウ上端 - 黒帯高) / 表示高 * スライドポイント高
            return (screenY - _cachedWindowRect.Top - _offsetYInWindow) / _slidePixelHeight * slideH;
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
