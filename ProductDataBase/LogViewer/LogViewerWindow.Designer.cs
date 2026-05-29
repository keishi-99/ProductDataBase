namespace ProductDatabase.LogViewer {
    partial class LogViewerWindow {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing) {
            if (disposing) {
                components?.Dispose();
                _loadCts?.Dispose();
                _errorLoadCts?.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent() {
            this.MainTabControl = new TabControl();
            this.OperationLogTabPage = new TabPage();
            this.ErrorLogTabPage = new TabPage();

            this.TopPanel = new Panel();
            this.YearMonthLabel = new Label();
            this.YearMonthComboBox = new ComboBox();
            this.OperationTypeLabel = new Label();
            this.OperationTypeComboBox = new ComboBox();
            this.SearchLabel = new Label();
            this.SearchTextBox = new TextBox();
            this.SearchButton = new Button();
            this.LogDataGridView = new DataGridView();

            this.ErrorTopPanel = new Panel();
            this.ErrorYearMonthLabel = new Label();
            this.ErrorYearMonthComboBox = new ComboBox();
            this.ErrorSearchLabel = new Label();
            this.ErrorSearchTextBox = new TextBox();
            this.ErrorSearchButton = new Button();
            this.ErrorDataGridView = new DataGridView();

            this.BottomStatusStrip = new StatusStrip();
            this.CountLabel = new ToolStripStatusLabel();

            this.MainTabControl.SuspendLayout();
            this.OperationLogTabPage.SuspendLayout();
            this.ErrorLogTabPage.SuspendLayout();
            this.TopPanel.SuspendLayout();
            this.ErrorTopPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)this.LogDataGridView).BeginInit();
            ((System.ComponentModel.ISupportInitialize)this.ErrorDataGridView).BeginInit();
            this.BottomStatusStrip.SuspendLayout();
            this.SuspendLayout();

            // MainTabControl
            this.MainTabControl.Controls.Add(this.OperationLogTabPage);
            this.MainTabControl.Controls.Add(this.ErrorLogTabPage);
            this.MainTabControl.Dock = DockStyle.Fill;
            this.MainTabControl.Name = "MainTabControl";
            this.MainTabControl.TabIndex = 0;
            this.MainTabControl.SelectedIndexChanged += this.MainTabControl_SelectedIndexChanged;

            // OperationLogTabPage
            this.OperationLogTabPage.Controls.Add(this.LogDataGridView);
            this.OperationLogTabPage.Controls.Add(this.TopPanel);
            this.OperationLogTabPage.Name = "OperationLogTabPage";
            this.OperationLogTabPage.Padding = new Padding(3);
            this.OperationLogTabPage.TabIndex = 0;
            this.OperationLogTabPage.Text = "操作ログ";

            // ErrorLogTabPage
            this.ErrorLogTabPage.Controls.Add(this.ErrorDataGridView);
            this.ErrorLogTabPage.Controls.Add(this.ErrorTopPanel);
            this.ErrorLogTabPage.Name = "ErrorLogTabPage";
            this.ErrorLogTabPage.Padding = new Padding(3);
            this.ErrorLogTabPage.TabIndex = 1;
            this.ErrorLogTabPage.Text = "エラーログ";

            // TopPanel
            this.TopPanel.Controls.Add(this.YearMonthLabel);
            this.TopPanel.Controls.Add(this.YearMonthComboBox);
            this.TopPanel.Controls.Add(this.OperationTypeLabel);
            this.TopPanel.Controls.Add(this.OperationTypeComboBox);
            this.TopPanel.Controls.Add(this.SearchLabel);
            this.TopPanel.Controls.Add(this.SearchTextBox);
            this.TopPanel.Controls.Add(this.SearchButton);
            this.TopPanel.Dock = DockStyle.Top;
            this.TopPanel.Height = 40;
            this.TopPanel.Name = "TopPanel";
            this.TopPanel.TabIndex = 0;

            // YearMonthLabel
            this.YearMonthLabel.AutoSize = true;
            this.YearMonthLabel.Location = new Point(10, 12);
            this.YearMonthLabel.Name = "YearMonthLabel";
            this.YearMonthLabel.Text = "年月：";

            // YearMonthComboBox
            this.YearMonthComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.YearMonthComboBox.Location = new Point(52, 8);
            this.YearMonthComboBox.Name = "YearMonthComboBox";
            this.YearMonthComboBox.Size = new Size(130, 23);
            this.YearMonthComboBox.TabIndex = 1;
            this.YearMonthComboBox.SelectedIndexChanged += this.YearMonthComboBox_SelectedIndexChanged;

            // OperationTypeLabel
            this.OperationTypeLabel.AutoSize = true;
            this.OperationTypeLabel.Location = new Point(200, 12);
            this.OperationTypeLabel.Name = "OperationTypeLabel";
            this.OperationTypeLabel.Text = "操作種別：";

            // OperationTypeComboBox
            this.OperationTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.OperationTypeComboBox.Location = new Point(268, 8);
            this.OperationTypeComboBox.Name = "OperationTypeComboBox";
            this.OperationTypeComboBox.Size = new Size(150, 23);
            this.OperationTypeComboBox.TabIndex = 2;
            this.OperationTypeComboBox.SelectedIndexChanged += this.OperationTypeComboBox_SelectedIndexChanged;

            // SearchLabel
            this.SearchLabel.AutoSize = true;
            this.SearchLabel.Location = new Point(435, 12);
            this.SearchLabel.Name = "SearchLabel";
            this.SearchLabel.Text = "検索：";

            // SearchTextBox
            this.SearchTextBox.Location = new Point(472, 8);
            this.SearchTextBox.Name = "SearchTextBox";
            this.SearchTextBox.Size = new Size(200, 23);
            this.SearchTextBox.TabIndex = 3;
            this.SearchTextBox.KeyDown += this.SearchTextBox_KeyDown;

            // SearchButton
            this.SearchButton.Location = new Point(680, 7);
            this.SearchButton.Name = "SearchButton";
            this.SearchButton.Size = new Size(60, 25);
            this.SearchButton.TabIndex = 4;
            this.SearchButton.Text = "検索";
            this.SearchButton.UseVisualStyleBackColor = true;
            this.SearchButton.Click += this.SearchButton_Click;

            // LogDataGridView
            this.LogDataGridView.BackgroundColor = Color.White;
            this.LogDataGridView.BorderStyle = BorderStyle.None;
            this.LogDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.LogDataGridView.Dock = DockStyle.Fill;
            this.LogDataGridView.Name = "LogDataGridView";
            this.LogDataGridView.TabIndex = 5;

            // ErrorTopPanel
            this.ErrorTopPanel.Controls.Add(this.ErrorYearMonthLabel);
            this.ErrorTopPanel.Controls.Add(this.ErrorYearMonthComboBox);
            this.ErrorTopPanel.Controls.Add(this.ErrorSearchLabel);
            this.ErrorTopPanel.Controls.Add(this.ErrorSearchTextBox);
            this.ErrorTopPanel.Controls.Add(this.ErrorSearchButton);
            this.ErrorTopPanel.Dock = DockStyle.Top;
            this.ErrorTopPanel.Height = 40;
            this.ErrorTopPanel.Name = "ErrorTopPanel";
            this.ErrorTopPanel.TabIndex = 0;

            // ErrorYearMonthLabel
            this.ErrorYearMonthLabel.AutoSize = true;
            this.ErrorYearMonthLabel.Location = new Point(10, 12);
            this.ErrorYearMonthLabel.Name = "ErrorYearMonthLabel";
            this.ErrorYearMonthLabel.Text = "年月：";

            // ErrorYearMonthComboBox
            this.ErrorYearMonthComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.ErrorYearMonthComboBox.Location = new Point(52, 8);
            this.ErrorYearMonthComboBox.Name = "ErrorYearMonthComboBox";
            this.ErrorYearMonthComboBox.Size = new Size(130, 23);
            this.ErrorYearMonthComboBox.TabIndex = 1;
            this.ErrorYearMonthComboBox.SelectedIndexChanged += this.ErrorYearMonthComboBox_SelectedIndexChanged;

            // ErrorSearchLabel
            this.ErrorSearchLabel.AutoSize = true;
            this.ErrorSearchLabel.Location = new Point(200, 12);
            this.ErrorSearchLabel.Name = "ErrorSearchLabel";
            this.ErrorSearchLabel.Text = "検索：";

            // ErrorSearchTextBox
            this.ErrorSearchTextBox.Location = new Point(242, 8);
            this.ErrorSearchTextBox.Name = "ErrorSearchTextBox";
            this.ErrorSearchTextBox.Size = new Size(200, 23);
            this.ErrorSearchTextBox.TabIndex = 2;
            this.ErrorSearchTextBox.KeyDown += this.ErrorSearchTextBox_KeyDown;

            // ErrorSearchButton
            this.ErrorSearchButton.Location = new Point(450, 7);
            this.ErrorSearchButton.Name = "ErrorSearchButton";
            this.ErrorSearchButton.Size = new Size(60, 25);
            this.ErrorSearchButton.TabIndex = 3;
            this.ErrorSearchButton.Text = "検索";
            this.ErrorSearchButton.UseVisualStyleBackColor = true;
            this.ErrorSearchButton.Click += this.ErrorSearchButton_Click;

            // ErrorDataGridView
            this.ErrorDataGridView.BackgroundColor = Color.White;
            this.ErrorDataGridView.BorderStyle = BorderStyle.None;
            this.ErrorDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ErrorDataGridView.Dock = DockStyle.Fill;
            this.ErrorDataGridView.Name = "ErrorDataGridView";
            this.ErrorDataGridView.TabIndex = 5;

            // BottomStatusStrip
            this.BottomStatusStrip.Items.AddRange(new ToolStripItem[] { this.CountLabel });
            this.BottomStatusStrip.Location = new Point(0, 628);
            this.BottomStatusStrip.Name = "BottomStatusStrip";
            this.BottomStatusStrip.SizingGrip = false;
            this.BottomStatusStrip.Size = new Size(1150, 22);
            this.BottomStatusStrip.TabIndex = 6;

            // CountLabel
            this.CountLabel.Name = "CountLabel";
            this.CountLabel.Size = new Size(89, 17);
            this.CountLabel.Text = "0 件表示中";

            // LogViewerWindow
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1150, 650);
            this.Controls.Add(this.MainTabControl);
            this.Controls.Add(this.BottomStatusStrip);
            this.Font = new Font("Meiryo UI", 9F);
            this.Name = "LogViewerWindow";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "ログ閲覧";

            this.MainTabControl.ResumeLayout(false);
            this.OperationLogTabPage.ResumeLayout(false);
            this.ErrorLogTabPage.ResumeLayout(false);
            this.TopPanel.ResumeLayout(false);
            this.TopPanel.PerformLayout();
            this.ErrorTopPanel.ResumeLayout(false);
            this.ErrorTopPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)this.LogDataGridView).EndInit();
            ((System.ComponentModel.ISupportInitialize)this.ErrorDataGridView).EndInit();
            this.BottomStatusStrip.ResumeLayout(false);
            this.BottomStatusStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private TabControl MainTabControl;
        private TabPage OperationLogTabPage;
        private TabPage ErrorLogTabPage;
        private Panel TopPanel;
        private Label YearMonthLabel;
        private ComboBox YearMonthComboBox;
        private Label OperationTypeLabel;
        private ComboBox OperationTypeComboBox;
        private Label SearchLabel;
        private TextBox SearchTextBox;
        private Button SearchButton;
        private DataGridView LogDataGridView;
        private Panel ErrorTopPanel;
        private Label ErrorYearMonthLabel;
        private ComboBox ErrorYearMonthComboBox;
        private Label ErrorSearchLabel;
        private TextBox ErrorSearchTextBox;
        private Button ErrorSearchButton;
        private DataGridView ErrorDataGridView;
        private StatusStrip BottomStatusStrip;
        private ToolStripStatusLabel CountLabel;
    }
}
