using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTDragDropAddIn
{
    [ComVisible(true)]
    public class DragDropRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        // ペン色の定義
        private static readonly Dictionary<string, Color> PenColors = new Dictionary<string, Color>
        {
            { "Black",  Color.FromArgb(0x11, 0x11, 0x11) },
            { "Red",    Color.FromArgb(0xFF, 0x33, 0x33) },
            { "Blue",   Color.FromArgb(0x33, 0x55, 0xFF) },
            { "Green",  Color.FromArgb(0x33, 0xAA, 0x33) },
            { "White",  Color.FromArgb(0xFF, 0xFF, 0xFF) },
            { "Orange", Color.FromArgb(0xFF, 0x88, 0x00) },
            { "Purple", Color.FromArgb(0x99, 0x33, 0xCC) },
        };

        // 蛍光ペン色の定義
        private static readonly Dictionary<string, Color> MarkerColors = new Dictionary<string, Color>
        {
            { "Yellow",     Color.FromArgb(0xFF, 0xFF, 0x00) },
            { "Cyan",       Color.FromArgb(0x00, 0xCC, 0xFF) },
            { "Orange",     Color.FromArgb(0xFF, 0x99, 0x00) },
            { "LightGreen", Color.FromArgb(0x99, 0xFF, 0x33) },
            { "Pink",       Color.FromArgb(0xFF, 0x66, 0xCC) },
        };

        private static readonly Dictionary<string, string> PenLabels = new Dictionary<string, string>
        {
            { "Black", "黒" }, { "Red", "赤" }, { "Blue", "青" },
            { "Green", "緑" }, { "White", "白" }, { "Orange", "オレンジ" }, { "Purple", "紫" },
        };

        private static readonly Dictionary<string, string> MarkerLabels = new Dictionary<string, string>
        {
            { "Yellow", "黄" }, { "Cyan", "水色" }, { "Orange", "橙" },
            { "LightGreen", "黄緑" }, { "Pink", "ピンク" },
        };

        public DragDropRibbon() { }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PPTDragDropAddIn.DragDropRibbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            // 起動時にすべてのコントロールを設定値で初期化
            ribbon.Invalidate();
        }

        // ─── 制御グループ ───────────────────────────────────────

        public bool GetDragDropPressed(Office.IRibbonControl control)
            => Properties.Settings.Default.DragDropEnabled;

        public void OnToggleDragDrop(Office.IRibbonControl control, bool pressed)
        {
            Properties.Settings.Default.DragDropEnabled = pressed;
            Properties.Settings.Default.Save();
        }

        public bool GetPenPalettePressed(Office.IRibbonControl control)
            => Properties.Settings.Default.PenPaletteEnabled;

        public void OnTogglePenPalette(Office.IRibbonControl control, bool pressed)
        {
            Properties.Settings.Default.PenPaletteEnabled = pressed;
            Properties.Settings.Default.Save();
        }

        // ─── ペン設定グループ ────────────────────────────────────

        public string GetPenLabel(Office.IRibbonControl control)
        {
            var key = Properties.Settings.Default.LastPenColor;
            return "ペン: " + (PenLabels.ContainsKey(key) ? PenLabels[key] : key);
        }

        public string GetMarkerLabel(Office.IRibbonControl control)
        {
            var key = Properties.Settings.Default.LastMarkerColor;
            return "蛍光: " + (MarkerLabels.ContainsKey(key) ? MarkerLabels[key] : key);
        }

        public Bitmap GetPenImage(Office.IRibbonControl control)
        {
            var key = Properties.Settings.Default.LastPenColor;
            return CreateColorSwatch(PenColors.ContainsKey(key) ? PenColors[key] : Color.Black);
        }

        public Bitmap GetMarkerImage(Office.IRibbonControl control)
        {
            var key = Properties.Settings.Default.LastMarkerColor;
            return CreateColorSwatch(MarkerColors.ContainsKey(key) ? MarkerColors[key] : Color.Yellow);
        }

        public Bitmap GetColorImage(Office.IRibbonControl control)
        {
            var tag = control.Tag;
            Color c;
            if (PenColors.TryGetValue(tag, out c) || MarkerColors.TryGetValue(tag, out c))
                return CreateColorSwatch(c);
            return CreateColorSwatch(Color.Gray);
        }

        public void OnPenMain(Office.IRibbonControl control)
        {
            // メインボタン: 現在の LastPenColor を再確認・保存（色は変えない）
            Properties.Settings.Default.Save();
        }

        public void OnMarkerMain(Office.IRibbonControl control)
        {
            Properties.Settings.Default.Save();
        }

        public void OnPenColor(Office.IRibbonControl control)
        {
            Properties.Settings.Default.LastPenColor = control.Tag;
            Properties.Settings.Default.Save();
            ribbon.InvalidateControl("btnPenMain");
        }

        public void OnMarkerColor(Office.IRibbonControl control)
        {
            Properties.Settings.Default.LastMarkerColor = control.Tag;
            Properties.Settings.Default.Save();
            ribbon.InvalidateControl("btnMarkerMain");
        }

        // ─── コンテキストメニュー（既存） ───────────────────────

        public string GetDragLabel(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                if (app.SlideShowWindows.Count > 0) return "ドラッグ設定";
                var selection = app.ActiveWindow.Selection;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 0)
                {
                    var shape = selection.ShapeRange[1];
                    return shape.Name.StartsWith("Drag_") ? "ドラッグを無効にする" : "ドラッグを有効にする";
                }
            }
            catch { }
            return "ドラッグを有効にする";
        }

        public void OnToggleDrag(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                var selection = app.ActiveWindow.Selection;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    foreach (PowerPoint.Shape shape in selection.ShapeRange)
                    {
                        if (shape.Name.StartsWith("Drag_"))
                            shape.Name = shape.Name.Substring(5);
                        else
                            shape.Name = "Drag_" + shape.Name;
                    }
                    if (ribbon != null) ribbon.Invalidate();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("エラーが発生しました: " + ex.Message);
            }
        }

        // ─── ヘルパー ────────────────────────────────────────────

        private static Bitmap CreateColorSwatch(Color color)
        {
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.FillRectangle(new SolidBrush(color), 1, 1, 14, 14);
                g.DrawRectangle(Pens.Gray, 0, 0, 15, 15);
            }
            return bmp;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            foreach (string name in resourceNames)
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (Stream stream = asm.GetManifestResourceStream(name))
                    {
                        if (stream != null)
                        {
                            using (StreamReader reader = new StreamReader(stream))
                                return reader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }

}
