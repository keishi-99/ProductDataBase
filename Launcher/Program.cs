using System.Diagnostics;
using System.Threading;

namespace Launcher {
    internal static class Program {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        private static Mutex? s_mutex;
        [STAThread]
        static void Main() {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.

            try {
                const string MutexName = "LauncherAppMutex"; // アプリケーション固有の名前を指定

                // ミューテックスを作成
                s_mutex = new Mutex(true, MutexName, out var isNewInstance);

                if (!isNewInstance) {
                    // すでにインスタンスが存在する場合
                    MessageBox.Show("このアプリケーションはすでに起動しています。", "重複起動の防止",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                SplashForm splash = new SplashForm();
                splash.Show();
                splash.Refresh();

                string mainAppPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ProductDatabase.exe");

                if (File.Exists(mainAppPath)) {

                    Process startApp = new Process();
                    startApp.StartInfo.FileName = mainAppPath;

                    startApp.Start();

                    // 本体のウィンドウが出るまでループで待機（応答なし防止）
                    Stopwatch sw = Stopwatch.StartNew();
                    while (!startApp.HasExited && sw.ElapsedMilliseconds < 20000) {
                        // 本体のウィンドウハンドルができたらループを抜ける
                        if (startApp.MainWindowHandle != IntPtr.Zero) break;

                        Application.DoEvents(); // スプラッシュが固まるのを防ぐ
                        Thread.Sleep(100);      // CPU負荷を抑える
                    }

                    // 4. 【重要】本体のウィンドウが表示されるまで待機
                    // これをしないと、本体が出る前にランチャーが消えて「一瞬無反応」に見えます。
                    // ただし、最大20秒などでタイムアウトさせるのが安全です。
                    try {
                        // 本体の初期化（アイドル状態）を待つ
                        startApp.WaitForInputIdle(20000);
                    } catch { /* タイムアウト等は無視 */ }
                }
                else {
                    MessageBox.Show("本体ファイル(ProductDatabase.exe)が見つかりません。");
                }

                // 5. ランチャーを終了（スプラッシュも一緒に消える）
                splash.Close();
                Application.Exit();

            } finally {
                if (s_mutex != null) {
                    s_mutex.ReleaseMutex();
                    s_mutex.Dispose();
                }
            }
        }
    }
}