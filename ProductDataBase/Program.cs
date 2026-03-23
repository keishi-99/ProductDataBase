namespace ProductDatabase {
    internal static class Program {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        ///
        private static Mutex? _mutex;
        [STAThread]
        private static void Main() {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            const string MutexName = "MyUniqueAppMutex"; // �A�v���P�[�V�����ŗL�̖��O���w��

            // �~���[�e�b�N�X���쐬
            _mutex = new Mutex(true, MutexName, out var isNewInstance);

            if (!isNewInstance) {
                // ���łɃC���X�^���X�����݂���ꍇ
                MessageBox.Show("���̃A�v���P�[�V�����͂��łɋN�����Ă��܂��B", "�d���N���̖h�~",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // �A�v���P�[�V���������s
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // UI�X���b�h��O
            Application.ThreadException += (s, e) => {
                MessageBox.Show(e.Exception.ToString(), "UI��O");
            };

            // ��UI�X���b�h��O
            AppDomain.CurrentDomain.UnhandledException += (s, e) => {
                if (e.ExceptionObject is Exception ex)
                    MessageBox.Show(ex.ToString(), "��UI��O");
            };

            try {
                Application.Run(new MainWindow());
            } finally {
                // アプリケーション終了時にミューテックスを確実に解放・破棄
                _mutex?.ReleaseMutex();
                _mutex?.Dispose();
            }
        }
    }
}