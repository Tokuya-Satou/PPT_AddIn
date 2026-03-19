using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace PPTDragDropAddIn
{
    [ComVisible(true)]
    public class DragDropRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public DragDropRibbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PPTDragDropAddIn.DragDropRibbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public string GetDragLabel(Office.IRibbonControl control)
        {
            try
            {
                // PowerPointのApplicationを明示的に取得して衝突を避ける
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
                        {
                            shape.Name = shape.Name.Substring(5);
                        }
                        else
                        {
                            shape.Name = "Drag_" + shape.Name;
                        }
                    }
                    if (ribbon != null) ribbon.Invalidate();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("エラーが発生しました: " + ex.Message);
            }
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
                            {
                                return reader.ReadToEnd();
                            }
                        }
                    }
                }
            }
            return null;
        }
    }
}
