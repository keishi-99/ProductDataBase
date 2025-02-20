namespace ProductDatabase {
    partial class HistoryWindow {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.DataBaseDataGridView = new DataGridView();
            this.CategoryComboBox = new ComboBox();
            this.FilterStringTextBox = new TextBox();
            this.CategoryRadioButton1 = new RadioButton();
            this.CategoryRadioButton2 = new RadioButton();
            this.StockCheckBox = new CheckBox();
            this.AllSubstrateStockCheckBox = new CheckBox();
            this.GenerateReportButton = new Button();
            this.GenerateListButton = new Button();
            this.CategoryRadioButton3 = new RadioButton();
            this.GenerateCheckSheetButton = new Button();
            this.AllSubstrateCheckBox = new CheckBox();
            this.menuStrip1 = new MenuStrip();
            this.ファイルToolStripMenuItem = new ToolStripMenuItem();
            this.編集モードToolStripMenuItem = new ToolStripMenuItem();
            this.編集終了ToolStripMenuItem = new ToolStripMenuItem();
            this.GroupModelCheckBox = new CheckBox();
            ((System.ComponentModel.ISupportInitialize)this.DataBaseDataGridView).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // DataBaseDataGridView
            // 
            this.DataBaseDataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.DataBaseDataGridView.EditMode = DataGridViewEditMode.EditOnF2;
            this.DataBaseDataGridView.Location = new Point(0, 27);
            this.DataBaseDataGridView.Name = "DataBaseDataGridView";
            this.DataBaseDataGridView.Size = new Size(1184, 467);
            this.DataBaseDataGridView.TabIndex = 0;
            // 
            // CategoryComboBox
            // 
            this.CategoryComboBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.CategoryComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.CategoryComboBox.FormattingEnabled = true;
            this.CategoryComboBox.Location = new Point(12, 502);
            this.CategoryComboBox.Name = "CategoryComboBox";
            this.CategoryComboBox.Size = new Size(121, 23);
            this.CategoryComboBox.TabIndex = 1;
            this.CategoryComboBox.SelectedIndexChanged += this.CategoryComboBox_SelectedIndexChanged;
            // 
            // FilterStringTextBox
            // 
            this.FilterStringTextBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.FilterStringTextBox.Location = new Point(139, 502);
            this.FilterStringTextBox.MaxLength = 50;
            this.FilterStringTextBox.Name = "FilterStringTextBox";
            this.FilterStringTextBox.Size = new Size(100, 23);
            this.FilterStringTextBox.TabIndex = 2;
            this.FilterStringTextBox.TextChanged += this.FilterStringTextBox_TextChanged;
            // 
            // CategoryRadioButton1
            // 
            this.CategoryRadioButton1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.CategoryRadioButton1.Appearance = Appearance.Button;
            this.CategoryRadioButton1.Location = new Point(261, 500);
            this.CategoryRadioButton1.Name = "CategoryRadioButton1";
            this.CategoryRadioButton1.Size = new Size(92, 25);
            this.CategoryRadioButton1.TabIndex = 3;
            this.CategoryRadioButton1.TabStop = true;
            this.CategoryRadioButton1.Tag = "1";
            this.CategoryRadioButton1.Text = "登録履歴";
            this.CategoryRadioButton1.TextAlign = ContentAlignment.MiddleCenter;
            this.CategoryRadioButton1.UseVisualStyleBackColor = true;
            this.CategoryRadioButton1.CheckedChanged += this.CategoryRadioButton_CheckedChanged;
            // 
            // CategoryRadioButton2
            // 
            this.CategoryRadioButton2.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.CategoryRadioButton2.Appearance = Appearance.Button;
            this.CategoryRadioButton2.Location = new Point(359, 500);
            this.CategoryRadioButton2.Name = "CategoryRadioButton2";
            this.CategoryRadioButton2.Size = new Size(92, 25);
            this.CategoryRadioButton2.TabIndex = 4;
            this.CategoryRadioButton2.TabStop = true;
            this.CategoryRadioButton2.Tag = "2";
            this.CategoryRadioButton2.Text = "Button2";
            this.CategoryRadioButton2.TextAlign = ContentAlignment.MiddleCenter;
            this.CategoryRadioButton2.UseVisualStyleBackColor = true;
            this.CategoryRadioButton2.CheckedChanged += this.CategoryRadioButton_CheckedChanged;
            // 
            // StockCheckBox
            // 
            this.StockCheckBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.StockCheckBox.AutoSize = true;
            this.StockCheckBox.Location = new Point(623, 506);
            this.StockCheckBox.Name = "StockCheckBox";
            this.StockCheckBox.Size = new Size(83, 19);
            this.StockCheckBox.TabIndex = 5;
            this.StockCheckBox.Text = "在庫有のみ";
            this.StockCheckBox.UseVisualStyleBackColor = true;
            this.StockCheckBox.CheckedChanged += this.StockCheckBox_CheckedChanged;
            // 
            // AllSubstrateStockCheckBox
            // 
            this.AllSubstrateStockCheckBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.AllSubstrateStockCheckBox.AutoSize = true;
            this.AllSubstrateStockCheckBox.Location = new Point(712, 506);
            this.AllSubstrateStockCheckBox.Name = "AllSubstrateStockCheckBox";
            this.AllSubstrateStockCheckBox.Size = new Size(62, 19);
            this.AllSubstrateStockCheckBox.TabIndex = 6;
            this.AllSubstrateStockCheckBox.Text = "他基板";
            this.AllSubstrateStockCheckBox.UseVisualStyleBackColor = true;
            this.AllSubstrateStockCheckBox.CheckedChanged += this.AllSubstrateStockCheckBox_CheckedChanged;
            // 
            // GenerateReportButton
            // 
            this.GenerateReportButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.GenerateReportButton.Location = new Point(1080, 500);
            this.GenerateReportButton.Name = "GenerateReportButton";
            this.GenerateReportButton.Size = new Size(92, 25);
            this.GenerateReportButton.TabIndex = 7;
            this.GenerateReportButton.Text = "成績書作成";
            this.GenerateReportButton.UseVisualStyleBackColor = true;
            this.GenerateReportButton.Click += this.GenerateReportButton_Click;
            // 
            // GenerateListButton
            // 
            this.GenerateListButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.GenerateListButton.Location = new Point(982, 500);
            this.GenerateListButton.Name = "GenerateListButton";
            this.GenerateListButton.Size = new Size(92, 25);
            this.GenerateListButton.TabIndex = 8;
            this.GenerateListButton.Text = "リスト作成";
            this.GenerateListButton.UseVisualStyleBackColor = true;
            this.GenerateListButton.Click += this.GenerateListButton_Click;
            // 
            // CategoryRadioButton3
            // 
            this.CategoryRadioButton3.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.CategoryRadioButton3.Appearance = Appearance.Button;
            this.CategoryRadioButton3.Location = new Point(457, 500);
            this.CategoryRadioButton3.Name = "CategoryRadioButton3";
            this.CategoryRadioButton3.Size = new Size(92, 25);
            this.CategoryRadioButton3.TabIndex = 9;
            this.CategoryRadioButton3.TabStop = true;
            this.CategoryRadioButton3.Tag = "3";
            this.CategoryRadioButton3.Text = "Button3";
            this.CategoryRadioButton3.TextAlign = ContentAlignment.MiddleCenter;
            this.CategoryRadioButton3.UseVisualStyleBackColor = true;
            this.CategoryRadioButton3.CheckedChanged += this.CategoryRadioButton_CheckedChanged;
            // 
            // GenerateCheckSheetButton
            // 
            this.GenerateCheckSheetButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.GenerateCheckSheetButton.Location = new Point(884, 500);
            this.GenerateCheckSheetButton.Name = "GenerateCheckSheetButton";
            this.GenerateCheckSheetButton.Size = new Size(92, 25);
            this.GenerateCheckSheetButton.TabIndex = 10;
            this.GenerateCheckSheetButton.Text = "チェックシート";
            this.GenerateCheckSheetButton.UseVisualStyleBackColor = true;
            this.GenerateCheckSheetButton.Click += this.GenerateCheckSheetButton_Click;
            // 
            // AllSubstrateCheckBox
            // 
            this.AllSubstrateCheckBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.AllSubstrateCheckBox.AutoSize = true;
            this.AllSubstrateCheckBox.Location = new Point(555, 506);
            this.AllSubstrateCheckBox.Name = "AllSubstrateCheckBox";
            this.AllSubstrateCheckBox.Size = new Size(62, 19);
            this.AllSubstrateCheckBox.TabIndex = 11;
            this.AllSubstrateCheckBox.Text = "他基板";
            this.AllSubstrateCheckBox.UseVisualStyleBackColor = true;
            this.AllSubstrateCheckBox.CheckedChanged += this.AllSubstrateCheckBox_CheckedChanged;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new ToolStripItem[] { this.ファイルToolStripMenuItem });
            this.menuStrip1.Location = new Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new Size(1184, 24);
            this.menuStrip1.TabIndex = 12;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // ファイルToolStripMenuItem
            // 
            this.ファイルToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { this.編集モードToolStripMenuItem, this.編集終了ToolStripMenuItem });
            this.ファイルToolStripMenuItem.Name = "ファイルToolStripMenuItem";
            this.ファイルToolStripMenuItem.Size = new Size(53, 20);
            this.ファイルToolStripMenuItem.Text = "ファイル";
            // 
            // 編集モードToolStripMenuItem
            // 
            this.編集モードToolStripMenuItem.Enabled = false;
            this.編集モードToolStripMenuItem.Name = "編集モードToolStripMenuItem";
            this.編集モードToolStripMenuItem.Size = new Size(125, 22);
            this.編集モードToolStripMenuItem.Text = "編集モード";
            this.編集モードToolStripMenuItem.Click += this.編集ToolStripMenuItem_Click;
            // 
            // 編集終了ToolStripMenuItem
            // 
            this.編集終了ToolStripMenuItem.Enabled = false;
            this.編集終了ToolStripMenuItem.Name = "編集終了ToolStripMenuItem";
            this.編集終了ToolStripMenuItem.Size = new Size(125, 22);
            this.編集終了ToolStripMenuItem.Text = "編集終了";
            this.編集終了ToolStripMenuItem.Click += this.編集終了ToolStripMenuItem_Click;
            // 
            // GroupModelCheckBox
            // 
            this.GroupModelCheckBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.GroupModelCheckBox.AutoSize = true;
            this.GroupModelCheckBox.Location = new Point(780, 506);
            this.GroupModelCheckBox.Name = "GroupModelCheckBox";
            this.GroupModelCheckBox.Size = new Size(62, 19);
            this.GroupModelCheckBox.TabIndex = 13;
            this.GroupModelCheckBox.Text = "型式毎";
            this.GroupModelCheckBox.UseVisualStyleBackColor = true;
            this.GroupModelCheckBox.CheckedChanged += this.GroupModelCheckBox_CheckedChanged;
            // 
            // HistoryWindow
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1184, 537);
            this.Controls.Add(this.GroupModelCheckBox);
            this.Controls.Add(this.AllSubstrateCheckBox);
            this.Controls.Add(this.GenerateCheckSheetButton);
            this.Controls.Add(this.CategoryRadioButton3);
            this.Controls.Add(this.GenerateListButton);
            this.Controls.Add(this.GenerateReportButton);
            this.Controls.Add(this.AllSubstrateStockCheckBox);
            this.Controls.Add(this.StockCheckBox);
            this.Controls.Add(this.CategoryRadioButton2);
            this.Controls.Add(this.CategoryRadioButton1);
            this.Controls.Add(this.FilterStringTextBox);
            this.Controls.Add(this.CategoryComboBox);
            this.Controls.Add(this.DataBaseDataGridView);
            this.Controls.Add(this.menuStrip1);
            this.Font = new Font("Meiryo UI", 9F);
            this.MainMenuStrip = this.menuStrip1;
            this.MinimizeBox = false;
            this.Name = "HistoryWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "データベース";
            this.Load += this.HistoryWindow_Load;
            ((System.ComponentModel.ISupportInitialize)this.DataBaseDataGridView).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private DataGridView DataBaseDataGridView;
        private ComboBox CategoryComboBox;
        private TextBox FilterStringTextBox;
        private RadioButton CategoryRadioButton1;
        private RadioButton CategoryRadioButton2;
        private CheckBox StockCheckBox;
        private CheckBox AllSubstrateStockCheckBox;
        private Button GenerateReportButton;
        private Button GenerateListButton;
        private RadioButton CategoryRadioButton3;
        private Button GenerateCheckSheetButton;
        private CheckBox AllSubstrateCheckBox;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem ファイルToolStripMenuItem;
        private ToolStripMenuItem 編集モードToolStripMenuItem;
        private ToolStripMenuItem 編集終了ToolStripMenuItem;
        private CheckBox GroupModelCheckBox;
    }
}