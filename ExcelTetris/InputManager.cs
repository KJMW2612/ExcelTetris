using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelTetris
{
    public class InputManager
    {
        private const int WH_KEYBOARD = 2;
        private const int HC_ACTION = 0;

        private LowLevelKeyboardProc _proc;
        private IntPtr _hookID = IntPtr.Zero;
        private Action<Keys> _onKeyPress;
        private Func<bool> _isBossModeActiveCheck;
        private Func<bool> _isInputActiveCheck;

        public InputManager(Action<Keys> onKeyPress, Func<bool> isBossModeActiveCheck, Func<bool> isInputActiveCheck)
        {
            _proc = HookCallback;
            _onKeyPress = onKeyPress;
            _isBossModeActiveCheck = isBossModeActiveCheck;
            _isInputActiveCheck = isInputActiveCheck;
        }

        public void Start()
        {
            if (_hookID == IntPtr.Zero)
            {
                uint threadId = GetCurrentThreadId();
                _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, threadId);

                if (_hookID == IntPtr.Zero)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    Debug.WriteLine($"로컬 키보드 훅 등록 실패. 에러 코드: {errorCode}");
                }
            }
        }

        public void Stop()
        {
            if (_hookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int vkCode = (int)wParam;
                Keys key = (Keys)vkCode;

                long lParamVal = lParam.ToInt64();
                bool isKeyDown = (lParamVal & 0x80000000) == 0;

                if (isKeyDown && IsExcelActiveWindow())
                {
                    if (IsGameKey(key))
                    {
                        ThisAddIn.Instance.SafeInvoke(() =>
                        {
                            try
                            {
                                _onKeyPress?.Invoke(key);
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"입력 콜백 실행 실패: {ex.Message}");
                            }
                        });

                        return (IntPtr)1;
                    }
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private bool IsExcelActiveWindow()
        {
            IntPtr hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return false;

            uint foregroundPid;
            GetWindowThreadProcessId(hwnd, out foregroundPid);

            return foregroundPid == (uint)Process.GetCurrentProcess().Id;
        }

        private bool IsGameKey(Keys key)
        {
            if (key == Keys.Escape) return true;

            if (_isInputActiveCheck != null && !_isInputActiveCheck())
            {
                return false;
            }

            // A, D, S, W, E, Q, Space 이외에 흑백 카무플라주 스위치용 F 키를 감지 대상에 결합합니다.
            return key == Keys.A || key == Keys.D || key == Keys.S || key == Keys.W ||
                   key == Keys.E || key == Keys.Q || key == Keys.Space || key == Keys.F;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    }
}