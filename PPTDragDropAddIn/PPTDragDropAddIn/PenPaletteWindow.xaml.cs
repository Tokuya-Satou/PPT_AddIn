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

        private void BtnBlackPen_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.SetPenMode(Colors.Black);

        private void BtnRedPen_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.SetPenMode(Colors.Red);

        private void BtnBluePen_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.SetPenMode(Colors.Blue);

        private void BtnYellowMarker_Click(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.SetPenMode(Color.FromRgb(255, 255, 0), thickness: 12.0, isHighlighter: true);

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
