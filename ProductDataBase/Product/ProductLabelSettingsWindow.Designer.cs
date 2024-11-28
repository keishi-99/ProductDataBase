namespace ProductDatabase {
    partial class ProductLabelSettingsWindow {
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
            this.OKButton = new Button();
            this.PrintTextGroupBox = new GroupBox();
            this.Label41 = new Label();
            this.Label40 = new Label();
            this.PrintTextFormatTextBox = new TextBox();
            this.Label39 = new Label();
            this.PrintTextFontButton = new Button();
            this.PrintTextFontTextBox = new TextBox();
            this.Label36 = new Label();
            this.PrintTextQuantityTextBox = new TextBox();
            this.Label35 = new Label();
            this.TextCenterCheckBox = new CheckBox();
            this.PrintTextPostionYTextBox = new TextBox();
            this.Label33 = new Label();
            this.PrintTextPostionXTextBox = new TextBox();
            this.Label32 = new Label();
            this.WhiteSpaceGroupBox = new GroupBox();
            this.Label11 = new Label();
            this.Label10 = new Label();
            this.PageOffsetYTextBox = new TextBox();
            this.PageOffsetXTextBox = new TextBox();
            this.Label9 = new Label();
            this.Label8 = new Label();
            this.QuantityGroupBox = new GroupBox();
            this.Label7 = new Label();
            this.QuantityYTextBox = new TextBox();
            this.QuantityXTextBox = new TextBox();
            this.Label6 = new Label();
            this.Label5 = new Label();
            this.LabelSizeGroupBox = new GroupBox();
            this.Label4 = new Label();
            this.Label3 = new Label();
            this.LabelHeightTextBox = new TextBox();
            this.LabelWidthTextBox = new TextBox();
            this.Label2 = new Label();
            this.Label1 = new Label();
            this.HeaderFooterGroupBox = new GroupBox();
            this.Label30 = new Label();
            this.Label29 = new Label();
            this.HeaderStringTextBox = new TextBox();
            this.Label22 = new Label();
            this.Label21 = new Label();
            this.Label20 = new Label();
            this.HeaderPostionYTextBox = new TextBox();
            this.HeaderPostionXTextBox = new TextBox();
            this.Label19 = new Label();
            this.Label18 = new Label();
            this.Label17 = new Label();
            this.HeaderFooterFontButton = new Button();
            this.HeaderFooterFontTextBox = new TextBox();
            this.Label16 = new Label();
            this.LabelIntervalGroupBox = new GroupBox();
            this.Label15 = new Label();
            this.Label14 = new Label();
            this.LabelIntervalYTextBox = new TextBox();
            this.LabelIntervalXTextBox = new TextBox();
            this.Label13 = new Label();
            this.Label12 = new Label();
            this.CloseButton = new Button();
            this.HeaderFontDialog = new FontDialog();
            this.TextFontDialog = new FontDialog();
            this.PrintTextGroupBox.SuspendLayout();
            this.WhiteSpaceGroupBox.SuspendLayout();
            this.QuantityGroupBox.SuspendLayout();
            this.LabelSizeGroupBox.SuspendLayout();
            this.HeaderFooterGroupBox.SuspendLayout();
            this.LabelIntervalGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // OKButton
            // 
            this.OKButton.Location = new Point(471, 524);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new Size(75, 25);
            this.OKButton.TabIndex = 24;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Click += this.BtnOK_Click;
            // 
            // PrintTextGroupBox
            // 
            this.PrintTextGroupBox.Controls.Add(this.Label41);
            this.PrintTextGroupBox.Controls.Add(this.Label40);
            this.PrintTextGroupBox.Controls.Add(this.PrintTextFormatTextBox);
            this.PrintTextGroupBox.Controls.Add(this.Label39);
            this.PrintTextGroupBox.Controls.Add(this.PrintTextFontButton);
            this.PrintTextGroupBox.Controls.Add(this.PrintTextFontTextBox);
            this.PrintTextGroupBox.Controls.Add(this.Label36);
            this.PrintTextGroupBox.Controls.Add(this.PrintTextQuantityTextBox);
            this.PrintTextGroupBox.Controls.Add(this.Label35);
            this.PrintTextGroupBox.Controls.Add(this.TextCenterCheckBox);
            this.PrintTextGroupBox.Controls.Add(this.PrintTextPostionYTextBox);
            this.PrintTextGroupBox.Controls.Add(this.Label33);
            this.PrintTextGroupBox.Controls.Add(this.PrintTextPostionXTextBox);
            this.PrintTextGroupBox.Controls.Add(this.Label32);
            this.PrintTextGroupBox.Location = new Point(27, 309);
            this.PrintTextGroupBox.Name = "PrintTextGroupBox";
            this.PrintTextGroupBox.Size = new Size(600, 170);
            this.PrintTextGroupBox.TabIndex = 22;
            this.PrintTextGroupBox.TabStop = false;
            this.PrintTextGroupBox.Text = "印刷文字設定";
            // 
            // Label41
            // 
            this.Label41.AutoSize = true;
            this.Label41.Location = new Point(318, 124);
            this.Label41.Name = "Label41";
            this.Label41.Size = new Size(281, 15);
            this.Label41.TabIndex = 30;
            this.Label41.Text = "%Y: 製造年 %MM: 製造月(2桁) %M: 製造月(1桁)";
            // 
            // Label40
            // 
            this.Label40.AutoSize = true;
            this.Label40.Location = new Point(318, 109);
            this.Label40.Name = "Label40";
            this.Label40.Size = new Size(201, 15);
            this.Label40.TabIndex = 24;
            this.Label40.Text = "%T: 型式 %R: リビジョン %S: シリアル";
            // 
            // PrintTextFormatTextBox
            // 
            this.PrintTextFormatTextBox.Location = new Point(379, 74);
            this.PrintTextFormatTextBox.MaxLength = 50;
            this.PrintTextFormatTextBox.Name = "PrintTextFormatTextBox";
            this.PrintTextFormatTextBox.Size = new Size(180, 23);
            this.PrintTextFormatTextBox.TabIndex = 24;
            // 
            // Label39
            // 
            this.Label39.AutoSize = true;
            this.Label39.Location = new Point(318, 77);
            this.Label39.Name = "Label39";
            this.Label39.Size = new Size(57, 15);
            this.Label39.TabIndex = 24;
            this.Label39.Text = "フォーマット";
            // 
            // PrintTextFontButton
            // 
            this.PrintTextFontButton.Location = new Point(479, 29);
            this.PrintTextFontButton.Name = "PrintTextFontButton";
            this.PrintTextFontButton.Size = new Size(96, 25);
            this.PrintTextFontButton.TabIndex = 24;
            this.PrintTextFontButton.Text = "フォント選択";
            this.PrintTextFontButton.UseVisualStyleBackColor = true;
            this.PrintTextFontButton.Click += this.PrintTextFontButton_Click;
            // 
            // PrintTextFontTextBox
            // 
            this.PrintTextFontTextBox.Location = new Point(320, 30);
            this.PrintTextFontTextBox.MaxLength = 50;
            this.PrintTextFontTextBox.Name = "PrintTextFontTextBox";
            this.PrintTextFontTextBox.Size = new Size(135, 23);
            this.PrintTextFontTextBox.TabIndex = 24;
            // 
            // Label36
            // 
            this.Label36.AutoSize = true;
            this.Label36.Location = new Point(318, 15);
            this.Label36.Name = "Label36";
            this.Label36.Size = new Size(40, 15);
            this.Label36.TabIndex = 24;
            this.Label36.Text = "フォント";
            // 
            // PrintTextQuantityTextBox
            // 
            this.PrintTextQuantityTextBox.Location = new Point(22, 92);
            this.PrintTextQuantityTextBox.MaxLength = 5;
            this.PrintTextQuantityTextBox.Name = "PrintTextQuantityTextBox";
            this.PrintTextQuantityTextBox.Size = new Size(52, 23);
            this.PrintTextQuantityTextBox.TabIndex = 10;
            // 
            // Label35
            // 
            this.Label35.AutoSize = true;
            this.Label35.Location = new Point(20, 77);
            this.Label35.Name = "Label35";
            this.Label35.Size = new Size(55, 15);
            this.Label35.TabIndex = 9;
            this.Label35.Text = "発行枚数";
            // 
            // TextCenterCheckBox
            // 
            this.TextCenterCheckBox.AutoSize = true;
            this.TextCenterCheckBox.Location = new Point(22, 55);
            this.TextCenterCheckBox.Name = "TextCenterCheckBox";
            this.TextCenterCheckBox.Size = new Size(145, 19);
            this.TextCenterCheckBox.TabIndex = 6;
            this.TextCenterCheckBox.Text = "横位置を中心に合わせる";
            this.TextCenterCheckBox.UseVisualStyleBackColor = true;
            this.TextCenterCheckBox.CheckedChanged += this.TextCenterCheckBox_CheckedChanged;
            // 
            // PrintTextPostionYTextBox
            // 
            this.PrintTextPostionYTextBox.Location = new Point(112, 30);
            this.PrintTextPostionYTextBox.MaxLength = 5;
            this.PrintTextPostionYTextBox.Name = "PrintTextPostionYTextBox";
            this.PrintTextPostionYTextBox.Size = new Size(68, 23);
            this.PrintTextPostionYTextBox.TabIndex = 5;
            // 
            // Label33
            // 
            this.Label33.AutoSize = true;
            this.Label33.Location = new Point(110, 15);
            this.Label33.Name = "Label33";
            this.Label33.Size = new Size(43, 15);
            this.Label33.TabIndex = 4;
            this.Label33.Text = "縦位置";
            // 
            // PrintTextPostionXTextBox
            // 
            this.PrintTextPostionXTextBox.Location = new Point(22, 30);
            this.PrintTextPostionXTextBox.MaxLength = 5;
            this.PrintTextPostionXTextBox.Name = "PrintTextPostionXTextBox";
            this.PrintTextPostionXTextBox.Size = new Size(68, 23);
            this.PrintTextPostionXTextBox.TabIndex = 3;
            // 
            // Label32
            // 
            this.Label32.AutoSize = true;
            this.Label32.Location = new Point(20, 15);
            this.Label32.Name = "Label32";
            this.Label32.Size = new Size(43, 15);
            this.Label32.TabIndex = 2;
            this.Label32.Text = "横位置";
            // 
            // WhiteSpaceGroupBox
            // 
            this.WhiteSpaceGroupBox.Controls.Add(this.Label11);
            this.WhiteSpaceGroupBox.Controls.Add(this.Label10);
            this.WhiteSpaceGroupBox.Controls.Add(this.PageOffsetYTextBox);
            this.WhiteSpaceGroupBox.Controls.Add(this.PageOffsetXTextBox);
            this.WhiteSpaceGroupBox.Controls.Add(this.Label9);
            this.WhiteSpaceGroupBox.Controls.Add(this.Label8);
            this.WhiteSpaceGroupBox.Location = new Point(27, 158);
            this.WhiteSpaceGroupBox.Name = "WhiteSpaceGroupBox";
            this.WhiteSpaceGroupBox.Size = new Size(311, 67);
            this.WhiteSpaceGroupBox.TabIndex = 19;
            this.WhiteSpaceGroupBox.TabStop = false;
            this.WhiteSpaceGroupBox.Text = "用紙余白";
            // 
            // Label11
            // 
            this.Label11.AutoSize = true;
            this.Label11.Location = new Point(261, 39);
            this.Label11.Name = "Label11";
            this.Label11.Size = new Size(31, 15);
            this.Label11.TabIndex = 5;
            this.Label11.Text = "mm";
            // 
            // Label10
            // 
            this.Label10.AutoSize = true;
            this.Label10.Location = new Point(126, 39);
            this.Label10.Name = "Label10";
            this.Label10.Size = new Size(31, 15);
            this.Label10.TabIndex = 4;
            this.Label10.Text = "mm";
            // 
            // PageOffsetYTextBox
            // 
            this.PageOffsetYTextBox.Location = new Point(158, 36);
            this.PageOffsetYTextBox.MaxLength = 5;
            this.PageOffsetYTextBox.Name = "PageOffsetYTextBox";
            this.PageOffsetYTextBox.Size = new Size(100, 23);
            this.PageOffsetYTextBox.TabIndex = 3;
            // 
            // PageOffsetXTextBox
            // 
            this.PageOffsetXTextBox.Location = new Point(20, 36);
            this.PageOffsetXTextBox.MaxLength = 5;
            this.PageOffsetXTextBox.Name = "PageOffsetXTextBox";
            this.PageOffsetXTextBox.Size = new Size(100, 23);
            this.PageOffsetXTextBox.TabIndex = 1;
            // 
            // Label9
            // 
            this.Label9.AutoSize = true;
            this.Label9.Location = new Point(156, 21);
            this.Label9.Name = "Label9";
            this.Label9.Size = new Size(43, 15);
            this.Label9.TabIndex = 2;
            this.Label9.Text = "上余白";
            // 
            // Label8
            // 
            this.Label8.AutoSize = true;
            this.Label8.Location = new Point(18, 21);
            this.Label8.Name = "Label8";
            this.Label8.Size = new Size(43, 15);
            this.Label8.TabIndex = 1;
            this.Label8.Text = "左余白";
            // 
            // QuantityGroupBox
            // 
            this.QuantityGroupBox.Controls.Add(this.Label7);
            this.QuantityGroupBox.Controls.Add(this.QuantityYTextBox);
            this.QuantityGroupBox.Controls.Add(this.QuantityXTextBox);
            this.QuantityGroupBox.Controls.Add(this.Label6);
            this.QuantityGroupBox.Controls.Add(this.Label5);
            this.QuantityGroupBox.Location = new Point(27, 85);
            this.QuantityGroupBox.Name = "QuantityGroupBox";
            this.QuantityGroupBox.Size = new Size(311, 67);
            this.QuantityGroupBox.TabIndex = 20;
            this.QuantityGroupBox.TabStop = false;
            this.QuantityGroupBox.Text = "配置数";
            // 
            // Label7
            // 
            this.Label7.AutoSize = true;
            this.Label7.Location = new Point(132, 38);
            this.Label7.Name = "Label7";
            this.Label7.Size = new Size(17, 15);
            this.Label7.TabIndex = 4;
            this.Label7.Text = "×";
            // 
            // QuantityYTextBox
            // 
            this.QuantityYTextBox.Location = new Point(158, 36);
            this.QuantityYTextBox.MaxLength = 5;
            this.QuantityYTextBox.Name = "QuantityYTextBox";
            this.QuantityYTextBox.Size = new Size(100, 23);
            this.QuantityYTextBox.TabIndex = 3;
            // 
            // QuantityXTextBox
            // 
            this.QuantityXTextBox.Location = new Point(20, 36);
            this.QuantityXTextBox.MaxLength = 5;
            this.QuantityXTextBox.Name = "QuantityXTextBox";
            this.QuantityXTextBox.Size = new Size(100, 23);
            this.QuantityXTextBox.TabIndex = 1;
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new Point(156, 21);
            this.Label6.Name = "Label6";
            this.Label6.Size = new Size(19, 15);
            this.Label6.TabIndex = 2;
            this.Label6.Text = "縦";
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new Point(18, 21);
            this.Label5.Name = "Label5";
            this.Label5.Size = new Size(19, 15);
            this.Label5.TabIndex = 1;
            this.Label5.Text = "横";
            // 
            // LabelSizeGroupBox
            // 
            this.LabelSizeGroupBox.Controls.Add(this.Label4);
            this.LabelSizeGroupBox.Controls.Add(this.Label3);
            this.LabelSizeGroupBox.Controls.Add(this.LabelHeightTextBox);
            this.LabelSizeGroupBox.Controls.Add(this.LabelWidthTextBox);
            this.LabelSizeGroupBox.Controls.Add(this.Label2);
            this.LabelSizeGroupBox.Controls.Add(this.Label1);
            this.LabelSizeGroupBox.Location = new Point(27, 12);
            this.LabelSizeGroupBox.Name = "LabelSizeGroupBox";
            this.LabelSizeGroupBox.Size = new Size(311, 67);
            this.LabelSizeGroupBox.TabIndex = 18;
            this.LabelSizeGroupBox.TabStop = false;
            this.LabelSizeGroupBox.Text = "ラベルサイズ";
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new Point(261, 39);
            this.Label4.Name = "Label4";
            this.Label4.Size = new Size(31, 15);
            this.Label4.TabIndex = 5;
            this.Label4.Text = "mm";
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new Point(126, 39);
            this.Label3.Name = "Label3";
            this.Label3.Size = new Size(31, 15);
            this.Label3.TabIndex = 4;
            this.Label3.Text = "mm";
            // 
            // LabelHeightTextBox
            // 
            this.LabelHeightTextBox.Location = new Point(158, 36);
            this.LabelHeightTextBox.MaxLength = 5;
            this.LabelHeightTextBox.Name = "LabelHeightTextBox";
            this.LabelHeightTextBox.Size = new Size(100, 23);
            this.LabelHeightTextBox.TabIndex = 3;
            // 
            // LabelWidthTextBox
            // 
            this.LabelWidthTextBox.Location = new Point(20, 36);
            this.LabelWidthTextBox.MaxLength = 5;
            this.LabelWidthTextBox.Name = "LabelWidthTextBox";
            this.LabelWidthTextBox.Size = new Size(100, 23);
            this.LabelWidthTextBox.TabIndex = 1;
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new Point(156, 21);
            this.Label2.Name = "Label2";
            this.Label2.Size = new Size(19, 15);
            this.Label2.TabIndex = 2;
            this.Label2.Text = "縦";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new Point(18, 21);
            this.Label1.Name = "Label1";
            this.Label1.Size = new Size(19, 15);
            this.Label1.TabIndex = 1;
            this.Label1.Text = "横";
            // 
            // HeaderFooterGroupBox
            // 
            this.HeaderFooterGroupBox.Controls.Add(this.Label30);
            this.HeaderFooterGroupBox.Controls.Add(this.Label29);
            this.HeaderFooterGroupBox.Controls.Add(this.HeaderStringTextBox);
            this.HeaderFooterGroupBox.Controls.Add(this.Label22);
            this.HeaderFooterGroupBox.Controls.Add(this.Label21);
            this.HeaderFooterGroupBox.Controls.Add(this.Label20);
            this.HeaderFooterGroupBox.Controls.Add(this.HeaderPostionYTextBox);
            this.HeaderFooterGroupBox.Controls.Add(this.HeaderPostionXTextBox);
            this.HeaderFooterGroupBox.Controls.Add(this.Label19);
            this.HeaderFooterGroupBox.Controls.Add(this.Label18);
            this.HeaderFooterGroupBox.Controls.Add(this.Label17);
            this.HeaderFooterGroupBox.Controls.Add(this.HeaderFooterFontButton);
            this.HeaderFooterGroupBox.Controls.Add(this.HeaderFooterFontTextBox);
            this.HeaderFooterGroupBox.Controls.Add(this.Label16);
            this.HeaderFooterGroupBox.Location = new Point(357, 12);
            this.HeaderFooterGroupBox.Name = "HeaderFooterGroupBox";
            this.HeaderFooterGroupBox.Size = new Size(270, 213);
            this.HeaderFooterGroupBox.TabIndex = 21;
            this.HeaderFooterGroupBox.TabStop = false;
            this.HeaderFooterGroupBox.Text = "ヘッダ・フッタ";
            // 
            // Label30
            // 
            this.Label30.AutoSize = true;
            this.Label30.Location = new Point(6, 186);
            this.Label30.Name = "Label30";
            this.Label30.Size = new Size(188, 15);
            this.Label30.TabIndex = 23;
            this.Label30.Text = "%P 製品名 %N 台数  %U 担当者";
            // 
            // Label29
            // 
            this.Label29.AutoSize = true;
            this.Label29.Location = new Point(6, 171);
            this.Label29.Name = "Label29";
            this.Label29.Size = new Size(227, 15);
            this.Label29.TabIndex = 22;
            this.Label29.Text = "%D 印刷日 %M 製番 %O 注番 %T 型式";
            // 
            // HeaderStringTextBox
            // 
            this.HeaderStringTextBox.Location = new Point(79, 125);
            this.HeaderStringTextBox.MaxLength = 50;
            this.HeaderStringTextBox.Name = "HeaderStringTextBox";
            this.HeaderStringTextBox.Size = new Size(185, 23);
            this.HeaderStringTextBox.TabIndex = 12;
            // 
            // Label22
            // 
            this.Label22.AutoSize = true;
            this.Label22.Location = new Point(6, 128);
            this.Label22.Name = "Label22";
            this.Label22.Size = new Size(69, 15);
            this.Label22.TabIndex = 11;
            this.Label22.Text = "ヘッダ文字列";
            // 
            // Label21
            // 
            this.Label21.AutoSize = true;
            this.Label21.Location = new Point(238, 102);
            this.Label21.Name = "Label21";
            this.Label21.Size = new Size(31, 15);
            this.Label21.TabIndex = 10;
            this.Label21.Text = "mm";
            // 
            // Label20
            // 
            this.Label20.AutoSize = true;
            this.Label20.Location = new Point(106, 102);
            this.Label20.Name = "Label20";
            this.Label20.Size = new Size(31, 15);
            this.Label20.TabIndex = 9;
            this.Label20.Text = "mm";
            // 
            // HeaderPostionYTextBox
            // 
            this.HeaderPostionYTextBox.Location = new Point(168, 99);
            this.HeaderPostionYTextBox.MaxLength = 50;
            this.HeaderPostionYTextBox.Name = "HeaderPostionYTextBox";
            this.HeaderPostionYTextBox.Size = new Size(64, 23);
            this.HeaderPostionYTextBox.TabIndex = 8;
            // 
            // HeaderPostionXTextBox
            // 
            this.HeaderPostionXTextBox.Location = new Point(36, 99);
            this.HeaderPostionXTextBox.MaxLength = 50;
            this.HeaderPostionXTextBox.Name = "HeaderPostionXTextBox";
            this.HeaderPostionXTextBox.Size = new Size(64, 23);
            this.HeaderPostionXTextBox.TabIndex = 7;
            // 
            // Label19
            // 
            this.Label19.AutoSize = true;
            this.Label19.Location = new Point(166, 84);
            this.Label19.Name = "Label19";
            this.Label19.Size = new Size(19, 15);
            this.Label19.TabIndex = 6;
            this.Label19.Text = "縦";
            // 
            // Label18
            // 
            this.Label18.AutoSize = true;
            this.Label18.Location = new Point(34, 84);
            this.Label18.Name = "Label18";
            this.Label18.Size = new Size(19, 15);
            this.Label18.TabIndex = 5;
            this.Label18.Text = "横";
            // 
            // Label17
            // 
            this.Label17.AutoSize = true;
            this.Label17.Location = new Point(23, 64);
            this.Label17.Name = "Label17";
            this.Label17.Size = new Size(57, 15);
            this.Label17.TabIndex = 4;
            this.Label17.Text = "ヘッダ位置";
            // 
            // HeaderFooterFontButton
            // 
            this.HeaderFooterFontButton.Location = new Point(168, 35);
            this.HeaderFooterFontButton.Name = "HeaderFooterFontButton";
            this.HeaderFooterFontButton.Size = new Size(96, 25);
            this.HeaderFooterFontButton.TabIndex = 3;
            this.HeaderFooterFontButton.Text = "フォント選択";
            this.HeaderFooterFontButton.UseVisualStyleBackColor = true;
            this.HeaderFooterFontButton.Click += this.BtnHeaderFooterFont_Click;
            // 
            // HeaderFooterFontTextBox
            // 
            this.HeaderFooterFontTextBox.Location = new Point(27, 36);
            this.HeaderFooterFontTextBox.MaxLength = 50;
            this.HeaderFooterFontTextBox.Name = "HeaderFooterFontTextBox";
            this.HeaderFooterFontTextBox.Size = new Size(135, 23);
            this.HeaderFooterFontTextBox.TabIndex = 2;
            // 
            // Label16
            // 
            this.Label16.AutoSize = true;
            this.Label16.Location = new Point(25, 21);
            this.Label16.Name = "Label16";
            this.Label16.Size = new Size(40, 15);
            this.Label16.TabIndex = 1;
            this.Label16.Text = "フォント";
            // 
            // LabelIntervalGroupBox
            // 
            this.LabelIntervalGroupBox.Controls.Add(this.Label15);
            this.LabelIntervalGroupBox.Controls.Add(this.Label14);
            this.LabelIntervalGroupBox.Controls.Add(this.LabelIntervalYTextBox);
            this.LabelIntervalGroupBox.Controls.Add(this.LabelIntervalXTextBox);
            this.LabelIntervalGroupBox.Controls.Add(this.Label13);
            this.LabelIntervalGroupBox.Controls.Add(this.Label12);
            this.LabelIntervalGroupBox.Location = new Point(27, 231);
            this.LabelIntervalGroupBox.Name = "LabelIntervalGroupBox";
            this.LabelIntervalGroupBox.Size = new Size(311, 67);
            this.LabelIntervalGroupBox.TabIndex = 23;
            this.LabelIntervalGroupBox.TabStop = false;
            this.LabelIntervalGroupBox.Text = "ラベル間隔";
            // 
            // Label15
            // 
            this.Label15.AutoSize = true;
            this.Label15.Location = new Point(261, 39);
            this.Label15.Name = "Label15";
            this.Label15.Size = new Size(31, 15);
            this.Label15.TabIndex = 5;
            this.Label15.Text = "mm";
            // 
            // Label14
            // 
            this.Label14.AutoSize = true;
            this.Label14.Location = new Point(126, 39);
            this.Label14.Name = "Label14";
            this.Label14.Size = new Size(31, 15);
            this.Label14.TabIndex = 4;
            this.Label14.Text = "mm";
            // 
            // LabelIntervalYTextBox
            // 
            this.LabelIntervalYTextBox.Location = new Point(158, 36);
            this.LabelIntervalYTextBox.MaxLength = 5;
            this.LabelIntervalYTextBox.Name = "LabelIntervalYTextBox";
            this.LabelIntervalYTextBox.Size = new Size(100, 23);
            this.LabelIntervalYTextBox.TabIndex = 3;
            // 
            // LabelIntervalXTextBox
            // 
            this.LabelIntervalXTextBox.Location = new Point(20, 36);
            this.LabelIntervalXTextBox.MaxLength = 5;
            this.LabelIntervalXTextBox.Name = "LabelIntervalXTextBox";
            this.LabelIntervalXTextBox.Size = new Size(100, 23);
            this.LabelIntervalXTextBox.TabIndex = 1;
            // 
            // Label13
            // 
            this.Label13.AutoSize = true;
            this.Label13.Location = new Point(156, 21);
            this.Label13.Name = "Label13";
            this.Label13.Size = new Size(55, 15);
            this.Label13.TabIndex = 2;
            this.Label13.Text = "上下間隔";
            // 
            // Label12
            // 
            this.Label12.AutoSize = true;
            this.Label12.Location = new Point(18, 21);
            this.Label12.Name = "Label12";
            this.Label12.Size = new Size(55, 15);
            this.Label12.TabIndex = 1;
            this.Label12.Text = "左右間隔";
            // 
            // CloseButton
            // 
            this.CloseButton.Location = new Point(552, 524);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new Size(75, 25);
            this.CloseButton.TabIndex = 25;
            this.CloseButton.Text = "キャンセル";
            this.CloseButton.UseVisualStyleBackColor = true;
            this.CloseButton.Click += this.CloseButton_Click;
            // 
            // ProductLabelSettingsWindow
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.CancelButton = this.CloseButton;
            this.ClientSize = new Size(654, 561);
            this.Controls.Add(this.CloseButton);
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.PrintTextGroupBox);
            this.Controls.Add(this.WhiteSpaceGroupBox);
            this.Controls.Add(this.QuantityGroupBox);
            this.Controls.Add(this.LabelSizeGroupBox);
            this.Controls.Add(this.HeaderFooterGroupBox);
            this.Controls.Add(this.LabelIntervalGroupBox);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProductLabelSettingsWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "製品ラベル設定";
            this.Load += this.ProductPrintSetting_Load;
            this.PrintTextGroupBox.ResumeLayout(false);
            this.PrintTextGroupBox.PerformLayout();
            this.WhiteSpaceGroupBox.ResumeLayout(false);
            this.WhiteSpaceGroupBox.PerformLayout();
            this.QuantityGroupBox.ResumeLayout(false);
            this.QuantityGroupBox.PerformLayout();
            this.LabelSizeGroupBox.ResumeLayout(false);
            this.LabelSizeGroupBox.PerformLayout();
            this.HeaderFooterGroupBox.ResumeLayout(false);
            this.HeaderFooterGroupBox.PerformLayout();
            this.LabelIntervalGroupBox.ResumeLayout(false);
            this.LabelIntervalGroupBox.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        private Button OKButton;
        private GroupBox PrintTextGroupBox;
        private Label Label41;
        private Label Label40;
        private TextBox PrintTextFormatTextBox;
        private Label Label39;
        private Button PrintTextFontButton;
        private TextBox PrintTextFontTextBox;
        private Label Label36;
        private TextBox PrintTextQuantityTextBox;
        private Label Label35;
        private CheckBox TextCenterCheckBox;
        private TextBox PrintTextPostionYTextBox;
        private Label Label33;
        private TextBox PrintTextPostionXTextBox;
        private Label Label32;
        private GroupBox WhiteSpaceGroupBox;
        private Label Label11;
        private Label Label10;
        private TextBox PageOffsetYTextBox;
        private TextBox PageOffsetXTextBox;
        private Label Label9;
        private Label Label8;
        private GroupBox QuantityGroupBox;
        private Label Label7;
        private TextBox QuantityYTextBox;
        private TextBox QuantityXTextBox;
        private Label Label6;
        private Label Label5;
        private GroupBox LabelSizeGroupBox;
        private Label Label4;
        private Label Label3;
        private TextBox LabelHeightTextBox;
        private TextBox LabelWidthTextBox;
        private Label Label2;
        private Label Label1;
        private GroupBox HeaderFooterGroupBox;
        private Label Label30;
        private Label Label29;
        private TextBox HeaderStringTextBox;
        private Label Label22;
        private Label Label21;
        private Label Label20;
        private TextBox HeaderPostionYTextBox;
        private TextBox HeaderPostionXTextBox;
        private Label Label19;
        private Label Label18;
        private Label Label17;
        private Button HeaderFooterFontButton;
        private TextBox HeaderFooterFontTextBox;
        private Label Label16;
        private GroupBox LabelIntervalGroupBox;
        private Label Label15;
        private Label Label14;
        private TextBox LabelIntervalYTextBox;
        private TextBox LabelIntervalXTextBox;
        private Label Label13;
        private Label Label12;
        private Button CloseButton;
        private FontDialog HeaderFontDialog;
        private FontDialog TextFontDialog;
    }
}