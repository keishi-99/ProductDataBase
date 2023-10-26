namespace ProductDataBase {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
            MainWindowMenuStrip = new MenuStrip();
            ファイルToolStripMenuItem = new ToolStripMenuItem();
            終了ToolStripMenuItem = new ToolStripMenuItem();
            CategoryRadioButton1 = new RadioButton();
            CategoryRadioButton2 = new RadioButton();
            CategoryRadioButton3 = new RadioButton();
            CategoryRadioButton4 = new RadioButton();
            CategoryListBox1 = new ListBox();
            CategoryListBox2 = new ListBox();
            CategoryListBox3 = new ListBox();
            HistoryButton = new Button();
            RegisterButton = new Button();
            FontSizePanel = new Panel();
            FontSize16RadioButton = new RadioButton();
            FontSize12RadioButton = new RadioButton();
            FontSize9RadioButton = new RadioButton();
            FontSizeLabel = new Label();
            QRCodePanel = new Panel();
            QRCodeButton = new Button();
            QRCodeTextBox = new TextBox();
            QRCodeCheckBox = new CheckBox();
            MainWindowMenuStrip.SuspendLayout();
            FontSizePanel.SuspendLayout();
            QRCodePanel.SuspendLayout();
            SuspendLayout();
            // 
            // MainWindowMenuStrip
            // 
            MainWindowMenuStrip.Items.AddRange(new ToolStripItem[] { ファイルToolStripMenuItem });
            MainWindowMenuStrip.Location = new Point(0, 0);
            MainWindowMenuStrip.Name = "MainWindowMenuStrip";
            MainWindowMenuStrip.Size = new Size(684, 24);
            MainWindowMenuStrip.TabIndex = 0;
            MainWindowMenuStrip.Text = "menuStrip1";
            // 
            // ファイルToolStripMenuItem
            // 
            ファイルToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 終了ToolStripMenuItem });
            ファイルToolStripMenuItem.Name = "ファイルToolStripMenuItem";
            ファイルToolStripMenuItem.Size = new Size(53, 20);
            ファイルToolStripMenuItem.Text = "ファイル";
            // 
            // 終了ToolStripMenuItem
            // 
            終了ToolStripMenuItem.Name = "終了ToolStripMenuItem";
            終了ToolStripMenuItem.Size = new Size(98, 22);
            終了ToolStripMenuItem.Text = "終了";
            終了ToolStripMenuItem.Click += 終了ToolStripMenuItem_Click;
            // 
            // CategoryRadioButton1
            // 
            CategoryRadioButton1.Appearance = Appearance.Button;
            CategoryRadioButton1.Location = new Point(42, 47);
            CategoryRadioButton1.Name = "CategoryRadioButton1";
            CategoryRadioButton1.Size = new Size(70, 30);
            CategoryRadioButton1.TabIndex = 1;
            CategoryRadioButton1.Text = "基盤登録";
            CategoryRadioButton1.TextAlign = ContentAlignment.MiddleCenter;
            CategoryRadioButton1.UseVisualStyleBackColor = true;
            CategoryRadioButton1.CheckedChanged += CategoryRadioButton_CheckedChanged;
            // 
            // CategoryRadioButton2
            // 
            CategoryRadioButton2.Appearance = Appearance.Button;
            CategoryRadioButton2.Location = new Point(118, 47);
            CategoryRadioButton2.Name = "CategoryRadioButton2";
            CategoryRadioButton2.Size = new Size(70, 30);
            CategoryRadioButton2.TabIndex = 2;
            CategoryRadioButton2.Text = "製品登録";
            CategoryRadioButton2.TextAlign = ContentAlignment.MiddleCenter;
            CategoryRadioButton2.UseVisualStyleBackColor = true;
            CategoryRadioButton2.CheckedChanged += CategoryRadioButton_CheckedChanged;
            // 
            // CategoryRadioButton3
            // 
            CategoryRadioButton3.Appearance = Appearance.Button;
            CategoryRadioButton3.Location = new Point(194, 47);
            CategoryRadioButton3.Name = "CategoryRadioButton3";
            CategoryRadioButton3.Size = new Size(70, 30);
            CategoryRadioButton3.TabIndex = 3;
            CategoryRadioButton3.Text = "再印刷";
            CategoryRadioButton3.TextAlign = ContentAlignment.MiddleCenter;
            CategoryRadioButton3.UseVisualStyleBackColor = true;
            CategoryRadioButton3.CheckedChanged += CategoryRadioButton_CheckedChanged;
            // 
            // CategoryRadioButton4
            // 
            CategoryRadioButton4.Appearance = Appearance.Button;
            CategoryRadioButton4.Location = new Point(270, 47);
            CategoryRadioButton4.Name = "CategoryRadioButton4";
            CategoryRadioButton4.Size = new Size(70, 30);
            CategoryRadioButton4.TabIndex = 4;
            CategoryRadioButton4.Text = "基盤変更";
            CategoryRadioButton4.TextAlign = ContentAlignment.MiddleCenter;
            CategoryRadioButton4.UseVisualStyleBackColor = true;
            CategoryRadioButton4.CheckedChanged += CategoryRadioButton_CheckedChanged;
            // 
            // CategoryListBox1
            // 
            CategoryListBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            CategoryListBox1.FormattingEnabled = true;
            CategoryListBox1.HorizontalScrollbar = true;
            CategoryListBox1.ItemHeight = 15;
            CategoryListBox1.Location = new Point(42, 97);
            CategoryListBox1.Name = "CategoryListBox1";
            CategoryListBox1.Size = new Size(60, 214);
            CategoryListBox1.TabIndex = 5;
            CategoryListBox1.SelectedIndexChanged += CategoryListBox1_SelectedIndexChanged;
            // 
            // CategoryListBox2
            // 
            CategoryListBox2.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            CategoryListBox2.FormattingEnabled = true;
            CategoryListBox2.HorizontalScrollbar = true;
            CategoryListBox2.ItemHeight = 15;
            CategoryListBox2.Location = new Point(142, 97);
            CategoryListBox2.Name = "CategoryListBox2";
            CategoryListBox2.Size = new Size(210, 214);
            CategoryListBox2.TabIndex = 6;
            CategoryListBox2.SelectedIndexChanged += CategoryListBox2_SelectedIndexChanged;
            // 
            // CategoryListBox3
            // 
            CategoryListBox3.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            CategoryListBox3.FormattingEnabled = true;
            CategoryListBox3.HorizontalScrollbar = true;
            CategoryListBox3.ItemHeight = 15;
            CategoryListBox3.Location = new Point(392, 97);
            CategoryListBox3.Name = "CategoryListBox3";
            CategoryListBox3.Size = new Size(250, 214);
            CategoryListBox3.TabIndex = 7;
            CategoryListBox3.SelectedIndexChanged += CategoryListBox3_SelectedIndexChanged;
            // 
            // HistoryButton
            // 
            HistoryButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            HistoryButton.Location = new Point(392, 344);
            HistoryButton.Name = "HistoryButton";
            HistoryButton.Size = new Size(75, 25);
            HistoryButton.TabIndex = 8;
            HistoryButton.Text = "履歴";
            HistoryButton.UseVisualStyleBackColor = true;
            HistoryButton.Click += HistoryButton_Click;
            // 
            // RegisterButton
            // 
            RegisterButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            RegisterButton.Location = new Point(567, 344);
            RegisterButton.Name = "RegisterButton";
            RegisterButton.Size = new Size(75, 25);
            RegisterButton.TabIndex = 9;
            RegisterButton.Text = "登録";
            RegisterButton.UseVisualStyleBackColor = true;
            RegisterButton.Click += RegisterButton_Click;
            // 
            // FontSizePanel
            // 
            FontSizePanel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            FontSizePanel.Controls.Add(FontSize16RadioButton);
            FontSizePanel.Controls.Add(FontSize12RadioButton);
            FontSizePanel.Controls.Add(FontSize9RadioButton);
            FontSizePanel.Controls.Add(FontSizeLabel);
            FontSizePanel.Location = new Point(558, 27);
            FontSizePanel.Name = "FontSizePanel";
            FontSizePanel.Size = new Size(125, 64);
            FontSizePanel.TabIndex = 600;
            // 
            // FontSize16RadioButton
            // 
            FontSize16RadioButton.Appearance = Appearance.Button;
            FontSize16RadioButton.Location = new Point(82, 28);
            FontSize16RadioButton.Name = "FontSize16RadioButton";
            FontSize16RadioButton.Size = new Size(30, 22);
            FontSize16RadioButton.TabIndex = 604;
            FontSize16RadioButton.Text = "16";
            FontSize16RadioButton.TextAlign = ContentAlignment.MiddleCenter;
            FontSize16RadioButton.UseVisualStyleBackColor = true;
            FontSize16RadioButton.CheckedChanged += FontSize_CheckedChanged;
            // 
            // FontSize12RadioButton
            // 
            FontSize12RadioButton.Appearance = Appearance.Button;
            FontSize12RadioButton.Location = new Point(49, 28);
            FontSize12RadioButton.Name = "FontSize12RadioButton";
            FontSize12RadioButton.Size = new Size(30, 22);
            FontSize12RadioButton.TabIndex = 603;
            FontSize12RadioButton.Text = "12";
            FontSize12RadioButton.TextAlign = ContentAlignment.MiddleCenter;
            FontSize12RadioButton.UseVisualStyleBackColor = true;
            FontSize12RadioButton.CheckedChanged += FontSize_CheckedChanged;
            // 
            // FontSize9RadioButton
            // 
            FontSize9RadioButton.Appearance = Appearance.Button;
            FontSize9RadioButton.Checked = true;
            FontSize9RadioButton.Location = new Point(16, 28);
            FontSize9RadioButton.Name = "FontSize9RadioButton";
            FontSize9RadioButton.Size = new Size(30, 22);
            FontSize9RadioButton.TabIndex = 602;
            FontSize9RadioButton.TabStop = true;
            FontSize9RadioButton.Text = "9";
            FontSize9RadioButton.TextAlign = ContentAlignment.MiddleCenter;
            FontSize9RadioButton.UseVisualStyleBackColor = true;
            FontSize9RadioButton.CheckedChanged += FontSize_CheckedChanged;
            // 
            // FontSizeLabel
            // 
            FontSizeLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            FontSizeLabel.Location = new Point(29, 11);
            FontSizeLabel.Name = "FontSizeLabel";
            FontSizeLabel.Size = new Size(69, 15);
            FontSizeLabel.TabIndex = 601;
            FontSizeLabel.Text = "フォントサイズ";
            // 
            // QRCodePanel
            // 
            QRCodePanel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            QRCodePanel.BorderStyle = BorderStyle.FixedSingle;
            QRCodePanel.Controls.Add(QRCodeButton);
            QRCodePanel.Controls.Add(QRCodeTextBox);
            QRCodePanel.Location = new Point(42, 333);
            QRCodePanel.Name = "QRCodePanel";
            QRCodePanel.Size = new Size(218, 46);
            QRCodePanel.TabIndex = 500;
            // 
            // QRCodeButton
            // 
            QRCodeButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            QRCodeButton.Location = new Point(124, 16);
            QRCodeButton.Name = "QRCodeButton";
            QRCodeButton.Size = new Size(75, 23);
            QRCodeButton.TabIndex = 602;
            QRCodeButton.Text = "OK";
            QRCodeButton.UseVisualStyleBackColor = true;
            // 
            // QRCodeTextBox
            // 
            QRCodeTextBox.Location = new Point(18, 16);
            QRCodeTextBox.MaxLength = 50;
            QRCodeTextBox.Name = "QRCodeTextBox";
            QRCodeTextBox.Size = new Size(100, 23);
            QRCodeTextBox.TabIndex = 601;
            QRCodeTextBox.KeyDown += QRCodeTextBox_KeyDown;
            // 
            // QRCodeCheckBox
            // 
            QRCodeCheckBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            QRCodeCheckBox.Checked = true;
            QRCodeCheckBox.CheckState = CheckState.Checked;
            QRCodeCheckBox.Location = new Point(61, 325);
            QRCodeCheckBox.Name = "QRCodeCheckBox";
            QRCodeCheckBox.Size = new Size(69, 19);
            QRCodeCheckBox.TabIndex = 501;
            QRCodeCheckBox.Text = "QRコード";
            QRCodeCheckBox.UseVisualStyleBackColor = true;
            // 
            // MainWindow
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(684, 381);
            Controls.Add(QRCodeCheckBox);
            Controls.Add(QRCodePanel);
            Controls.Add(FontSizePanel);
            Controls.Add(RegisterButton);
            Controls.Add(HistoryButton);
            Controls.Add(CategoryListBox3);
            Controls.Add(CategoryListBox2);
            Controls.Add(CategoryListBox1);
            Controls.Add(CategoryRadioButton4);
            Controls.Add(CategoryRadioButton3);
            Controls.Add(CategoryRadioButton2);
            Controls.Add(CategoryRadioButton1);
            Controls.Add(MainWindowMenuStrip);
            Font = new Font("Meiryo UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = MainWindowMenuStrip;
            MaximizeBox = false;
            Name = "MainWindow";
            Text = "ProductDataBase";
            Load += MainWindow_Load;
            MainWindowMenuStrip.ResumeLayout(false);
            MainWindowMenuStrip.PerformLayout();
            FontSizePanel.ResumeLayout(false);
            QRCodePanel.ResumeLayout(false);
            QRCodePanel.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
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
        private CheckBox QRCodeCheckBox;
        private Button QRCodeButton;
    }
}