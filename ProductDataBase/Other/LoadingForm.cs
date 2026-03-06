namespace ProductDatabase.Other {
    public partial class LoadingForm : Form {
        public LoadingForm() {
            InitializeComponent();
            ControlBox = false;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            ShowIcon = false;
            ShowInTaskbar = false;
        }

        // フォーム表示後にControlBoxを無効化してユーザーが手動で閉じられないようにする
        protected override void OnShown(EventArgs e) {
            base.OnShown(e);
            // ユーザーが閉じられないようにする
            this.ControlBox = false;
        }
    }
}
