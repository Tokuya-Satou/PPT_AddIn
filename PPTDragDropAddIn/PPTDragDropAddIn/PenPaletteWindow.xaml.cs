using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace PPTDragDropAddIn
{
    public partial class PenPaletteWindow : Window
    {
        private const int WM_MOUSEACTIVATE = 0x0021;
        private const int MA_NOACTIVATE = 3;

        private bool _isCollapsed = false;

        public PenPaletteWindow()
        {
            InitializeComponent();
            this.Loaded += PenPaletteWindow_Loaded;
        }

        private void PenPaletteWindow_Loaded(object sender, RoutedEventArgs e)
        {
            UI_Pen1Ellipse.Fill = new SolidColorBrush(Globals.ThisAddIn.GetSavedPenColor(1));
            UI_Pen1Text.Text = " " + Globals.ThisAddIn.GetSavedPenName(1) + "ペン";

            UI_Pen2Ellipse.Fill = new SolidColorBrush(Globals.ThisAddIn.GetSavedPenColor(2));
            UI_Pen2Text.Text = " " + Globals.ThisAddIn.GetSavedPenName(2) + "ペン";

            UI_Pen3Ellipse.Fill = new SolidColorBrush(Globals.ThisAddIn.GetSavedPenColor(3));
            UI_Pen3Text.Text = " " + Globals.ThisAddIn.GetSavedPenName(3) + "ペン";

            UI_MarkerRect.Fill = new SolidColorBrush(Globals.ThisAddIn.GetSavedMarkerColor());
            UI_MarkerText.Text = " " + Globals.ThisAddIn.GetSavedMarkerName() + "マーカー";
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var source = HwndSource.FromHwnd(new WindowInteropHelper(this).Handle);
            source.AddHook(WndProc);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_MOUSEACTIVATE)
            {
                handled = true;
                return new IntPtr(MA_NOACTIVATE);
            }
            return IntPtr.Zero;
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                DragMove();
        }

        private void BtnToggle_Click(object sender, RoutedEventArgs e)
        {
            _isCollapsed = !_isCollapsed;
            PanelContent.Visibility = _isCollapsed ? Visibility.Collapsed : Visibility.Visible;
            BtnToggle.Content = _isCollapsed ? "＋" : "−";
        }

        private void BtnPen1_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.SetPenMode(Globals.ThisAddIn.GetSavedPenColor(1));

        private void BtnPen2_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.SetPenMode(Globals.ThisAddIn.GetSavedPenColor(2));

        private void BtnPen3_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.SetPenMode(Globals.ThisAddIn.GetSavedPenColor(3));

        private void BtnMarker_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.SetPenMode(Globals.ThisAddIn.GetSavedMarkerColor(), thickness: 12.0, isHighlighter: true);

        private void BtnEraser_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.SetEraserMode();

        private void BtnArrow_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.SetArrowMode();

        private void BtnPrev_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.GoPrevSlide();

        private void BtnNext_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.GoNextSlide();
    }
}
