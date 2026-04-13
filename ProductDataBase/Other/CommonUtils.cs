using System.Data;
using System.Runtime.InteropServices;

namespace ProductDatabase.Other {

    internal static partial class NativeMethods {
        [LibraryImport("user32.dll", SetLastError = true)]
        public static partial void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
    }

    internal partial class CommonUtils {

        // STAスレッドで処理を実行し Task として返す（COM Interop・ダイアログ用）
        // RunContinuationsAsynchronously: await後の継続処理がSTAスレッドで実行されるのを防ぐ
        public static Task RunOnStaThreadAsync(Action action) {
            var tcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var thread = new Thread(() => {
                try {
                    action();
                    tcs.SetResult();
                } catch (Exception ex) {
                    tcs.SetException(ex);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return tcs.Task;
        }

        /// <summary>
        /// 月コードを取得します（10月→X, 11月→Y, 12月→Z, それ以外→MM形式）。
        /// </summary>
        public static string ToMonthCode(DateTime date) =>
            date.Month switch {
                10 => "X",
                11 => "Y",
                12 => "Z",
                var m => m.ToString("00")
            };

        // CapsLockがオンになっていたらCapsLockを解除する
        public static partial class Keyboard {
            private const byte VK_CAPITAL = 0x14; // CapsLock の仮想キーコード
            private const int KEYEVENTF_EXTENDEDKEY = 0x1;
            private const int KEYEVENTF_KEYUP = 0x2;

            // CapsLock の状態を切り替える
            public static void CapsDisable() {
                if (Control.IsKeyLocked(Keys.CapsLock)) {
                    NativeMethods.keybd_event(VK_CAPITAL, 0, KEYEVENTF_EXTENDEDKEY, 0);
                    NativeMethods.keybd_event(VK_CAPITAL, 0, KEYEVENTF_KEYUP, 0);
                }
            }
        }
    }
}
