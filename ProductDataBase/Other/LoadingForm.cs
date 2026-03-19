namespace ProductDatabase.Other {
    public partial class LoadingForm : Form {
        public LoadingForm() {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;

            // 「処理中...」ラベルをフォーム下部に追加
            var label = new Label {
                Text = "処理中...",
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Bottom,
                Height = 28
            };
            Controls.Add(label);
        }
    }
}
