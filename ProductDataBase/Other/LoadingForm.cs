namespace ProductDatabase.Other {
    public partial class LoadingForm : Form {
        public LoadingForm() {
            InitializeComponent();
        }

        protected override void OnShown(EventArgs e) {
            base.OnShown(e);
            // ユーザーが閉じられないようにする
            this.ControlBox = false;
        }
    }
}
