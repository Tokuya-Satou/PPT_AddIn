using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PPTDragDropAddIn
{
    public class MouseHook
    {
        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;
        private const int WM_MOUSEMOVE = 0x0200;

        public event EventHandler<MouseEventArgs> MouseDown;
        public event EventHandler<MouseEventArgs> MouseUp;
        public event EventHandler<MouseEventArgs> MouseMove;

        private LowLevelMouseProc _proc;
        private IntPtr _hookId = IntPtr.Zero;

        // ドラッグ中かどうかを保持し、ドラッグ中のクリックイベントを遮断するか判定
        public bool IsDragging { get; set; } = false;

        public MouseHook()
        {
            _proc = HookCallback;
        }

        public void Install()
        {
            _hookId = SetHook(_proc);
        }

        public void Uninstall()
        {
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }

        private IntPtr SetHook(LowLevelMouseProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                MouseEventArgs e = new MouseEventArgs(MouseButtons.Left, 0, hookStruct.pt.x, hookStruct.pt.y, 0);

                int message = (int)wParam;

                if (message == WM_LBUTTONDOWN)
                {
                    MouseDown?.Invoke(this, e);
                    // ドラッグが開始された場合、このイベントをOSに流さないようにすることも検討できますが、
                    // ここではイベントを通知した上で、IsDraggingの状態によって後のUpを制御します。
                }
                else if (message == WM_LBUTTONUP)
                {
                    MouseUp?.Invoke(this, e);
                    // ドラッグ中だった場合、このマウスアップをフックして遮断することで、
                    // PowerPoint側に「クリック」として認識されるのを防ぎ、スライド遷移を抑制します。
                    if (IsDragging)
                    {
                        return (IntPtr)1; // イベントを消費（遮断）
                    }
                }
                else if (message == WM_MOUSEMOVE)
                {
                    MouseMove?.Invoke(this, e);
                }
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
    }
}
