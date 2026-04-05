using System;
using System.Runtime.InteropServices;

namespace PPTDragDropAddIn
{
    /// <summary>
    /// ドラッグ中に PowerPoint スライドショーウィンドウへの
    /// タッチ/ジェスチャーメッセージをブロックするクラス。
    /// VSTO は PowerPoint プロセス内で動くため、WH_GETMESSAGE フックを
    /// スライドショーウィンドウのスレッドに直接インストールできる。
    /// </summary>
    internal class GestureBlocker
    {
        private const int WH_GETMESSAGE = 3;
        private const uint WM_NULL            = 0x0000;
        private const uint WM_GESTURE         = 0x0119;
        private const uint WM_GESTURENOTIFY   = 0x011A;
        private const uint WM_TOUCH           = 0x0240;
        private const uint WM_POINTERDOWN     = 0x0246;
        private const uint WM_POINTERUPDATE   = 0x0245;
        private const uint WM_POINTERUP       = 0x0247;
        private const uint WM_POINTERENTER    = 0x0249;
        private const uint WM_POINTERLEAVE    = 0x024A;

        // ドラッグ中かどうか（volatile でクロススレッド読み書きを安全に）
        public volatile bool IsBlocking = false;

        /// <summary>
        /// WM_POINTERDOWN 受信時にスクリーン座標を渡してブロックするか判定するコールバック。
        /// PowerPoint のメインスレッドから呼ばれるため COM アクセス可。
        /// </summary>
        public Func<int, int, bool> ShouldBlock { get; set; }

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
        private HookProc _hookProc; // GC に回収されないよう保持
        private IntPtr _hookId = IntPtr.Zero;

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hwnd;
            public uint   message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint   time;
            public int    ptX;
            public int    ptY;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public void Install(IntPtr slideshowHwnd)
        {
            if (_hookId != IntPtr.Zero) return;
            uint processId;
            uint threadId = GetWindowThreadProcessId(slideshowHwnd, out processId);
            if (threadId == 0) return;

            _hookProc = HookCallback;
            // インプロセスフックは hMod = IntPtr.Zero で OK
            _hookId = SetWindowsHookEx(WH_GETMESSAGE, _hookProc, IntPtr.Zero, threadId);
        }

        public void Uninstall()
        {
            IsBlocking = false;
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                // wParam == 1 (PM_REMOVE): メッセージがキューから取り出される瞬間
                if (nCode >= 0 && wParam == (IntPtr)1)
                {
                    MSG msg = Marshal.PtrToStructure<MSG>(lParam);

                    // WM_POINTERDOWN 時点でまだブロック中でなければ当たり判定を試みる。
                    // これにより「タッチ検出 → ブロック設定」の間に PowerPoint が
                    // WM_POINTERDOWN を処理してしまう競合状態を回避する。
                    uint originalMessage = msg.message;

                    if (originalMessage == WM_POINTERDOWN && !IsBlocking && ShouldBlock != null)
                    {
                        // WM_POINTER* の座標は MSG 構造体の pt フィールド（ptX/ptY）にある
                        if (ShouldBlock(msg.ptX, msg.ptY))
                            IsBlocking = true;
                    }

                    // WM_POINTERUP でブロック終了（メッセージは書き換えない）
                    if (IsBlocking && originalMessage == WM_POINTERUP)
                        IsBlocking = false;

                    // ★ WM_NULL への書き換えはしない。
                    // WM_NULL への置換は PowerPoint の Direct Manipulation の状態機械を
                    // 壊してフリーズを引き起こす。タッチ吸収は TouchGuard（オーバーレイの
                    // 不透明化 + RootGrid.CaptureTouch）で行う。
                }
            }
            catch { }

            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }
    }
}
