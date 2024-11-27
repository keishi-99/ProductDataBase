namespace ProductDatabase {
    internal static class Program {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        ///
        private static Mutex? s_mutex;
        [STAThread]
        private static void Main() {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            const string MutexName = "MyUniqueAppMutex"; // アプリケーション固有の名前を指定

            // ミューテックスを作成
            s_mutex = new Mutex(true, MutexName, out var isNewInstance);

            if (!isNewInstance) {
                // すでにインスタンスが存在する場合
                MessageBox.Show("このアプリケーションはすでに起動しています。", "重複起動の防止",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // アプリケーションを実行
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainWindow());

            // アプリケーション終了時にミューテックスを解放
            s_mutex.ReleaseMutex();
        }
    }
}