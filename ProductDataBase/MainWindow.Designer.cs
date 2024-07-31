namespace ProductDatabase {
    partial class MainWindow {
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
            this.MainWindowMenuStrip = new MenuStrip();
            this.ファイルToolStripMenuItem = new ToolStripMenuItem();
            this.終了ToolStripMenuItem = new ToolStripMenuItem();
            this.CategoryRadioButton1 = new RadioButton();
            this.CategoryRadioButton2 = new RadioButton();
            this.CategoryRadioButton3 = new RadioButton();
            this.CategoryRadioButton4 = new RadioButton();
            this.CategoryListBox1 = new ListBox();
            this.CategoryListBox2 = new ListBox();
            this.CategoryListBox3 = new ListBox();
            this.HistoryButton = new Button();
            this.RegisterButton = new Button();
            this.FontSizePanel = new Panel();
            this.FontSize16RadioButton = new RadioButton();
            this.FontSize12RadioButton = new RadioButton();
            this.FontSize9RadioButton = new RadioButton();
            this.FontSizeLabel = new Label();
            this.QRCodePanel = new Panel();
            this.QRCodeButton = new Button();
            this.QRCodeTextBox = new TextBox();
            this.RadioButtonBarcode = new RadioButton();
            this.RadioButtonQR = new RadioButton();
            this.MainWindowMenuStrip.SuspendLayout();
            this.FontSizePanel.SuspendLayout();
            this.QRCodePanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // MainWindowMenuStrip
            // 
            this.MainWindowMenuStrip.Items.AddRange(new ToolStripItem[] { this.ファイルToolStripMenuItem });
            this.MainWindowMenuStrip.Location = new Point(0, 0);
            this.MainWindowMenuStrip.Name = "MainWindowMenuStrip";
            this.MainWindowMenuStrip.Size = new Size(684, 24);
            this.MainWindowMenuStrip.TabIndex = 0;
            this.MainWindowMenuStrip.Text = "menuStrip1";
            // 
            // ファイルToolStripMenuItem
            // 
            this.ファイルToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { this.終了ToolStripMenuItem });
            this.ファイルToolStripMenuItem.Name = "ファイルToolStripMenuItem";
            this.ファイルToolStripMenuItem.Size = new Size(53, 20);
            this.ファイルToolStripMenuItem.Text = "ファイル";
            // 
            // 終了ToolStripMenuItem
            // 
            this.終了ToolStripMenuItem.Name = "終了ToolStripMenuItem";
            this.終了ToolStripMenuItem.Size = new Size(98, 22);
            this.終了ToolStripMenuItem.Text = "終了";
            this.終了ToolStripMenuItem.Click += this.終了ToolStripMenuItem_Click;
            // 
            // CategoryRadioButton1
            // 
            this.CategoryRadioButton1.Appearance = Appearance.Button;
            this.CategoryRadioButton1.Location = new Point(42, 47);
            this.CategoryRadioButton1.Name = "CategoryRadioButton1";
            this.CategoryRadioButton1.Size = new Size(70, 30);
            this.CategoryRadioButton1.TabIndex = 1;
            this.CategoryRadioButton1.Tag = "1";
            this.CategoryRadioButton1.Text = "基盤登録";
            this.CategoryRadioButton1.TextAlign = ContentAlignment.MiddleCenter;
            this.CategoryRadioButton1.UseVisualStyleBackColor = true;
            this.CategoryRadioButton1.CheckedChanged += this.CategoryRadioButton_CheckedChanged;
            // 
            // CategoryRadioButton2
            // 
            this.CategoryRadioButton2.Appearance = Appearance.Button;
            this.CategoryRadioButton2.Location = new Point(118, 47);
            this.CategoryRadioButton2.Name = "CategoryRadioButton2";
            this.CategoryRadioButton2.Size = new Size(70, 30);
            this.CategoryRadioButton2.TabIndex = 2;
            this.CategoryRadioButton2.Tag = "2";
            this.CategoryRadioButton2.Text = "製品登録";
            this.CategoryRadioButton2.TextAlign = ContentAlignment.MiddleCenter;
            this.CategoryRadioButton2.UseVisualStyleBackColor = true;
            this.CategoryRadioButton2.CheckedChanged += this.CategoryRadioButton_CheckedChanged;
            // 
            // CategoryRadioButton3
            // 
            this.CategoryRadioButton3.Appearance = Appearance.Button;
            this.CategoryRadioButton3.Location = new Point(194, 47);
            this.CategoryRadioButton3.Name = "CategoryRadioButton3";
            this.CategoryRadioButton3.Size = new Size(70, 30);
            this.CategoryRadioButton3.TabIndex = 3;
            this.CategoryRadioButton3.Tag = "3";
            this.CategoryRadioButton3.Text = "再印刷";
            this.CategoryRadioButton3.TextAlign = ContentAlignment.MiddleCenter;
            this.CategoryRadioButton3.UseVisualStyleBackColor = true;
            this.CategoryRadioButton3.CheckedChanged += this.CategoryRadioButton_CheckedChanged;
            // 
            // CategoryRadioButton4
            // 
            this.CategoryRadioButton4.Appearance = Appearance.Button;
            this.CategoryRadioButton4.Location = new Point(270, 47);
            this.CategoryRadioButton4.Name = "CategoryRadioButton4";
            this.CategoryRadioButton4.Size = new Size(70, 30);
            this.CategoryRadioButton4.TabIndex = 4;
            this.CategoryRadioButton4.Tag = "4";
            this.CategoryRadioButton4.Text = "基盤変更";
            this.CategoryRadioButton4.TextAlign = ContentAlignment.MiddleCenter;
            this.CategoryRadioButton4.UseVisualStyleBackColor = true;
            this.CategoryRadioButton4.CheckedChanged += this.CategoryRadioButton_CheckedChanged;
            // 
            // CategoryListBox1
            // 
            this.CategoryListBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.CategoryListBox1.FormattingEnabled = true;
            this.CategoryListBox1.HorizontalScrollbar = true;
            this.CategoryListBox1.ItemHeight = 15;
            this.CategoryListBox1.Location = new Point(42, 97);
            this.CategoryListBox1.Name = "CategoryListBox1";
            this.CategoryListBox1.Size = new Size(60, 214);
            this.CategoryListBox1.TabIndex = 5;
            this.CategoryListBox1.SelectedIndexChanged += this.CategoryListBox1_SelectedIndexChanged;
            // 
            // CategoryListBox2
            // 
            this.CategoryListBox2.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.CategoryListBox2.FormattingEnabled = true;
            this.CategoryListBox2.HorizontalScrollbar = true;
            this.CategoryListBox2.ItemHeight = 15;
            this.CategoryListBox2.Location = new Point(142, 97);
            this.CategoryListBox2.Name = "CategoryListBox2";
            this.CategoryListBox2.Size = new Size(210, 214);
            this.CategoryListBox2.TabIndex = 6;
            this.CategoryListBox2.SelectedIndexChanged += this.CategoryListBox2_SelectedIndexChanged;
            // 
            // CategoryListBox3
            // 
            this.CategoryListBox3.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.CategoryListBox3.FormattingEnabled = true;
            this.CategoryListBox3.HorizontalScrollbar = true;
            this.CategoryListBox3.ItemHeight = 15;
            this.CategoryListBox3.Location = new Point(392, 97);
            this.CategoryListBox3.Name = "CategoryListBox3";
            this.CategoryListBox3.Size = new Size(250, 214);
            this.CategoryListBox3.TabIndex = 7;
            this.CategoryListBox3.SelectedIndexChanged += this.CategoryListBox3_SelectedIndexChanged;
            // 
            // HistoryButton
            // 
            this.HistoryButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.HistoryButton.Location = new Point(392, 354);
            this.HistoryButton.Name = "HistoryButton";
            this.HistoryButton.Size = new Size(75, 25);
            this.HistoryButton.TabIndex = 8;
            this.HistoryButton.Text = "履歴";
            this.HistoryButton.UseVisualStyleBackColor = true;
            this.HistoryButton.Click += this.HistoryButton_Click;
            // 
            // RegisterButton
            // 
            this.RegisterButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.RegisterButton.Location = new Point(567, 354);
            this.RegisterButton.Name = "RegisterButton";
            this.RegisterButton.Size = new Size(75, 25);
            this.RegisterButton.TabIndex = 9;
            this.RegisterButton.Text = "登録";
            this.RegisterButton.UseVisualStyleBackColor = true;
            this.RegisterButton.Click += this.RegisterButton_Click;
            // 
            // FontSizePanel
            // 
            this.FontSizePanel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.FontSizePanel.Controls.Add(this.FontSize16RadioButton);
            this.FontSizePanel.Controls.Add(this.FontSize12RadioButton);
            this.FontSizePanel.Controls.Add(this.FontSize9RadioButton);
            this.FontSizePanel.Controls.Add(this.FontSizeLabel);
            this.FontSizePanel.Location = new Point(558, 27);
            this.FontSizePanel.Name = "FontSizePanel";
            this.FontSizePanel.Size = new Size(125, 64);
            this.FontSizePanel.TabIndex = 600;
            // 
            // FontSize16RadioButton
            // 
            this.FontSize16RadioButton.Appearance = Appearance.Button;
            this.FontSize16RadioButton.Location = new Point(82, 28);
            this.FontSize16RadioButton.Name = "FontSize16RadioButton";
            this.FontSize16RadioButton.Size = new Size(30, 22);
            this.FontSize16RadioButton.TabIndex = 604;
            this.FontSize16RadioButton.Text = "16";
            this.FontSize16RadioButton.TextAlign = ContentAlignment.MiddleCenter;
            this.FontSize16RadioButton.UseVisualStyleBackColor = true;
            this.FontSize16RadioButton.CheckedChanged += this.FontSize_CheckedChanged;
            // 
            // FontSize12RadioButton
            // 
            this.FontSize12RadioButton.Appearance = Appearance.Button;
            this.FontSize12RadioButton.Location = new Point(49, 28);
            this.FontSize12RadioButton.Name = "FontSize12RadioButton";
            this.FontSize12RadioButton.Size = new Size(30, 22);
            this.FontSize12RadioButton.TabIndex = 603;
            this.FontSize12RadioButton.Text = "12";
            this.FontSize12RadioButton.TextAlign = ContentAlignment.MiddleCenter;
            this.FontSize12RadioButton.UseVisualStyleBackColor = true;
            this.FontSize12RadioButton.CheckedChanged += this.FontSize_CheckedChanged;
            // 
            // FontSize9RadioButton
            // 
            this.FontSize9RadioButton.Appearance = Appearance.Button;
            this.FontSize9RadioButton.Checked = true;
            this.FontSize9RadioButton.Location = new Point(16, 28);
            this.FontSize9RadioButton.Name = "FontSize9RadioButton";
            this.FontSize9RadioButton.Size = new Size(30, 22);
            this.FontSize9RadioButton.TabIndex = 602;
            this.FontSize9RadioButton.TabStop = true;
            this.FontSize9RadioButton.Text = "9";
            this.FontSize9RadioButton.TextAlign = ContentAlignment.MiddleCenter;
            this.FontSize9RadioButton.UseVisualStyleBackColor = true;
            this.FontSize9RadioButton.CheckedChanged += this.FontSize_CheckedChanged;
            // 
            // FontSizeLabel
            // 
            this.FontSizeLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.FontSizeLabel.Location = new Point(29, 11);
            this.FontSizeLabel.Name = "FontSizeLabel";
            this.FontSizeLabel.Size = new Size(69, 15);
            this.FontSizeLabel.TabIndex = 601;
            this.FontSizeLabel.Text = "フォントサイズ";
            // 
            // QRCodePanel
            // 
            this.QRCodePanel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.QRCodePanel.BorderStyle = BorderStyle.FixedSingle;
            this.QRCodePanel.Controls.Add(this.RadioButtonBarcode);
            this.QRCodePanel.Controls.Add(this.RadioButtonQR);
            this.QRCodePanel.Controls.Add(this.QRCodeButton);
            this.QRCodePanel.Controls.Add(this.QRCodeTextBox);
            this.QRCodePanel.Location = new Point(42, 333);
            this.QRCodePanel.Name = "QRCodePanel";
            this.QRCodePanel.Size = new Size(218, 56);
            this.QRCodePanel.TabIndex = 500;
            // 
            // QRCodeButton
            // 
            this.QRCodeButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.QRCodeButton.Location = new Point(124, 28);
            this.QRCodeButton.Name = "QRCodeButton";
            this.QRCodeButton.Size = new Size(75, 23);
            this.QRCodeButton.TabIndex = 602;
            this.QRCodeButton.Text = "OK";
            this.QRCodeButton.UseVisualStyleBackColor = true;
            this.QRCodeButton.Click += this.QRCodeButton_Click;
            // 
            // QRCodeTextBox
            // 
            this.QRCodeTextBox.Location = new Point(18, 28);
            this.QRCodeTextBox.MaxLength = 50;
            this.QRCodeTextBox.Name = "QRCodeTextBox";
            this.QRCodeTextBox.Size = new Size(100, 23);
            this.QRCodeTextBox.TabIndex = 601;
            this.QRCodeTextBox.KeyDown += this.QRCodeTextBox_KeyDown;
            // 
            // RadioButtonBarcode
            // 
            this.RadioButtonBarcode.AutoSize = true;
            this.RadioButtonBarcode.Location = new Point(97, 3);
            this.RadioButtonBarcode.Name = "RadioButtonBarcode";
            this.RadioButtonBarcode.Size = new Size(97, 19);
            this.RadioButtonBarcode.TabIndex = 606;
            this.RadioButtonBarcode.Text = "手配管理番号";
            this.RadioButtonBarcode.UseVisualStyleBackColor = true;
            // 
            // RadioButtonQR
            // 
            this.RadioButtonQR.AutoSize = true;
            this.RadioButtonQR.Checked = true;
            this.RadioButtonQR.Location = new Point(23, 3);
            this.RadioButtonQR.Name = "RadioButtonQR";
            this.RadioButtonQR.Size = new Size(68, 19);
            this.RadioButtonQR.TabIndex = 605;
            this.RadioButtonQR.TabStop = true;
            this.RadioButtonQR.Text = "QRコード";
            this.RadioButtonQR.UseVisualStyleBackColor = true;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(684, 391);
            this.Controls.Add(this.QRCodePanel);
            this.Controls.Add(this.FontSizePanel);
            this.Controls.Add(this.RegisterButton);
            this.Controls.Add(this.HistoryButton);
            this.Controls.Add(this.CategoryListBox3);
            this.Controls.Add(this.CategoryListBox2);
            this.Controls.Add(this.CategoryListBox1);
            this.Controls.Add(this.CategoryRadioButton4);
            this.Controls.Add(this.CategoryRadioButton3);
            this.Controls.Add(this.CategoryRadioButton2);
            this.Controls.Add(this.CategoryRadioButton1);
            this.Controls.Add(this.MainWindowMenuStrip);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.Icon = (Icon)resources.GetObject("$this.Icon");
            this.MainMenuStrip = this.MainWindowMenuStrip;
            this.MaximizeBox = false;
            this.Name = "MainWindow";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "ProductDataBase";
            this.Load += this.MainWindow_Load;
            this.MainWindowMenuStrip.ResumeLayout(false);
            this.MainWindowMenuStrip.PerformLayout();
            this.FontSizePanel.ResumeLayout(false);
            this.QRCodePanel.ResumeLayout(false);
            this.QRCodePanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private MenuStrip MainWindowMenuStrip;
        private ToolStripMenuItem ファイルToolStripMenuItem;
        private ToolStripMenuItem 終了ToolStripMenuItem;
        private RadioButton CategoryRadioButton1;
        private RadioButton CategoryRadioButton2;
        private RadioButton CategoryRadioButton3;
        private RadioButton CategoryRadioButton4;
        private ListBox CategoryListBox1;
        private ListBox CategoryListBox2;
        private ListBox CategoryListBox3;
        private Button HistoryButton;
        private Button RegisterButton;
        private Panel FontSizePanel;
        private RadioButton FontSize9RadioButton;
        private Label FontSizeLabel;
        private RadioButton FontSize16RadioButton;
        private RadioButton FontSize12RadioButton;
        private Panel QRCodePanel;
        private TextBox QRCodeTextBox;
        private Button QRCodeButton;
        private RadioButton RadioButtonBarcode;
        private RadioButton RadioButtonQR;
    }
}