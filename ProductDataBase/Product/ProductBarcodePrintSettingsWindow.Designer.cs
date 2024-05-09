namespace ProductDatabase {
    partial class ProductBarcodePrintSettingsWindow {
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
            OKButton = new Button();
            BarcodeGroupBox = new GroupBox();
            FontCenterCheckBox = new CheckBox();
            FontPostionYTextBox = new TextBox();
            FontPostionXTextBox = new TextBox();
            Label37 = new Label();
            Label28 = new Label();
            Label41 = new Label();
            Label40 = new Label();
            BarcodeFormatTextBox = new TextBox();
            Label39 = new Label();
            BarcodeFontButton = new Button();
            BarcodeFontTextBox = new TextBox();
            Label36 = new Label();
            BarcodeQuantityTextBox = new TextBox();
            Label35 = new Label();
            BarcodeMagnitudeTextBox = new TextBox();
            Label34 = new Label();
            BarcodeCenterCheckBox = new CheckBox();
            BarcodePostionYTextBox = new TextBox();
            Label33 = new Label();
            BarcodePostionXTextBox = new TextBox();
            Label32 = new Label();
            BarcodeHeightTextBox = new TextBox();
            Label31 = new Label();
            WhiteSpaceGroupBox = new GroupBox();
            Label11 = new Label();
            Label10 = new Label();
            BarcodePageOffsetYTextBox = new TextBox();
            BarcodePageOffsetXTextBox = new TextBox();
            Label9 = new Label();
            Label8 = new Label();
            QuantityGroupBox = new GroupBox();
            Label7 = new Label();
            BarcodeQuantityYTextBox = new TextBox();
            BarcodeQuantityXTextBox = new TextBox();
            Label6 = new Label();
            Label5 = new Label();
            BarcodeSizeGroupBox = new GroupBox();
            Label4 = new Label();
            Label3 = new Label();
            BarcodeLabelHeightTextBox = new TextBox();
            BarcodeLabelWidthTextBox = new TextBox();
            Label2 = new Label();
            Label1 = new Label();
            HeaderFooterGroupBox = new GroupBox();
            BarcodeFooterStringTextBox = new TextBox();
            Label23 = new Label();
            Label24 = new Label();
            BarcodeFooterPostionYTextBox = new TextBox();
            BarcodeFooterPostionXTextBox = new TextBox();
            Label25 = new Label();
            Label26 = new Label();
            Label27 = new Label();
            Label30 = new Label();
            Label29 = new Label();
            BarcodeHeaderStringTextBox = new TextBox();
            Label22 = new Label();
            Label21 = new Label();
            Label20 = new Label();
            BarcodeHeaderPostionYTextBox = new TextBox();
            BarcodeHeaderPostionXTextBox = new TextBox();
            Label19 = new Label();
            Label18 = new Label();
            Label17 = new Label();
            BarcodeHeaderFooterFontButton = new Button();
            BarcodeHeaderFooterFontTextBox = new TextBox();
            Label16 = new Label();
            LabelIntervalGroupBox = new GroupBox();
            Label15 = new Label();
            Label14 = new Label();
            BarcodeLabelIntervalYTextBox = new TextBox();
            BarcodeLabelIntervalXTextBox = new TextBox();
            Label13 = new Label();
            Label12 = new Label();
            CloseButton = new Button();
            BarcodeHeaderFontDialog = new FontDialog();
            BarcodeFontDialog = new FontDialog();
            BarcodeGroupBox.SuspendLayout();
            WhiteSpaceGroupBox.SuspendLayout();
            QuantityGroupBox.SuspendLayout();
            BarcodeSizeGroupBox.SuspendLayout();
            HeaderFooterGroupBox.SuspendLayout();
            LabelIntervalGroupBox.SuspendLayout();
            SuspendLayout();
            // 
            // OKButton
            // 
            OKButton.Location = new Point(471, 524);
            OKButton.Name = "OKButton";
            OKButton.Size = new Size(75, 25);
            OKButton.TabIndex = 32;
            OKButton.Text = "OK";
            OKButton.UseVisualStyleBackColor = true;
            OKButton.Click += BtnOK_Click;
            // 
            // BarcodeGroupBox
            // 
            BarcodeGroupBox.Controls.Add(FontCenterCheckBox);
            BarcodeGroupBox.Controls.Add(FontPostionYTextBox);
            BarcodeGroupBox.Controls.Add(FontPostionXTextBox);
            BarcodeGroupBox.Controls.Add(Label37);
            BarcodeGroupBox.Controls.Add(Label28);
            BarcodeGroupBox.Controls.Add(Label41);
            BarcodeGroupBox.Controls.Add(Label40);
            BarcodeGroupBox.Controls.Add(BarcodeFormatTextBox);
            BarcodeGroupBox.Controls.Add(Label39);
            BarcodeGroupBox.Controls.Add(BarcodeFontButton);
            BarcodeGroupBox.Controls.Add(BarcodeFontTextBox);
            BarcodeGroupBox.Controls.Add(Label36);
            BarcodeGroupBox.Controls.Add(BarcodeQuantityTextBox);
            BarcodeGroupBox.Controls.Add(Label35);
            BarcodeGroupBox.Controls.Add(BarcodeMagnitudeTextBox);
            BarcodeGroupBox.Controls.Add(Label34);
            BarcodeGroupBox.Controls.Add(BarcodeCenterCheckBox);
            BarcodeGroupBox.Controls.Add(BarcodePostionYTextBox);
            BarcodeGroupBox.Controls.Add(Label33);
            BarcodeGroupBox.Controls.Add(BarcodePostionXTextBox);
            BarcodeGroupBox.Controls.Add(Label32);
            BarcodeGroupBox.Controls.Add(BarcodeHeightTextBox);
            BarcodeGroupBox.Controls.Add(Label31);
            BarcodeGroupBox.Location = new Point(27, 309);
            BarcodeGroupBox.Name = "BarcodeGroupBox";
            BarcodeGroupBox.Size = new Size(600, 210);
            BarcodeGroupBox.TabIndex = 30;
            BarcodeGroupBox.TabStop = false;
            BarcodeGroupBox.Text = "バーコード";
            // 
            // FontCenterCheckBox
            // 
            FontCenterCheckBox.AutoSize = true;
            FontCenterCheckBox.Location = new Point(320, 104);
            FontCenterCheckBox.Name = "FontCenterCheckBox";
            FontCenterCheckBox.Size = new Size(145, 19);
            FontCenterCheckBox.TabIndex = 35;
            FontCenterCheckBox.Text = "横位置を中心に合わせる";
            FontCenterCheckBox.UseVisualStyleBackColor = true;
            FontCenterCheckBox.CheckedChanged += FontCenterCheckBox_CheckedChanged;
            // 
            // FontPostionYTextBox
            // 
            FontPostionYTextBox.Location = new Point(410, 79);
            FontPostionYTextBox.MaxLength = 5;
            FontPostionYTextBox.Name = "FontPostionYTextBox";
            FontPostionYTextBox.Size = new Size(68, 23);
            FontPostionYTextBox.TabIndex = 34;
            // 
            // FontPostionXTextBox
            // 
            FontPostionXTextBox.Location = new Point(320, 79);
            FontPostionXTextBox.MaxLength = 5;
            FontPostionXTextBox.Name = "FontPostionXTextBox";
            FontPostionXTextBox.Size = new Size(68, 23);
            FontPostionXTextBox.TabIndex = 33;
            // 
            // Label37
            // 
            Label37.AutoSize = true;
            Label37.Location = new Point(408, 64);
            Label37.Name = "Label37";
            Label37.Size = new Size(43, 15);
            Label37.TabIndex = 32;
            Label37.Text = "縦位置";
            // 
            // Label28
            // 
            Label28.AutoSize = true;
            Label28.Location = new Point(318, 64);
            Label28.Name = "Label28";
            Label28.Size = new Size(43, 15);
            Label28.TabIndex = 31;
            Label28.Text = "横位置";
            // 
            // Label41
            // 
            Label41.AutoSize = true;
            Label41.Location = new Point(318, 187);
            Label41.Name = "Label41";
            Label41.Size = new Size(281, 15);
            Label41.TabIndex = 30;
            Label41.Text = "%Y: 製造年 %MM: 製造月(2桁) %M: 製造月(1桁)";
            // 
            // Label40
            // 
            Label40.AutoSize = true;
            Label40.Location = new Point(318, 172);
            Label40.Name = "Label40";
            Label40.Size = new Size(201, 15);
            Label40.TabIndex = 24;
            Label40.Text = "%T: 型式 %R: リビジョン %S: シリアル";
            // 
            // BarcodeFormatTextBox
            // 
            BarcodeFormatTextBox.Location = new Point(379, 137);
            BarcodeFormatTextBox.MaxLength = 50;
            BarcodeFormatTextBox.Name = "BarcodeFormatTextBox";
            BarcodeFormatTextBox.Size = new Size(180, 23);
            BarcodeFormatTextBox.TabIndex = 24;
            // 
            // Label39
            // 
            Label39.AutoSize = true;
            Label39.Location = new Point(318, 140);
            Label39.Name = "Label39";
            Label39.Size = new Size(57, 15);
            Label39.TabIndex = 24;
            Label39.Text = "フォーマット";
            // 
            // BarcodeFontButton
            // 
            BarcodeFontButton.Location = new Point(479, 29);
            BarcodeFontButton.Name = "BarcodeFontButton";
            BarcodeFontButton.Size = new Size(96, 25);
            BarcodeFontButton.TabIndex = 24;
            BarcodeFontButton.Text = "フォント選択";
            BarcodeFontButton.UseVisualStyleBackColor = true;
            BarcodeFontButton.Click += BtnBarcodeFont_Click;
            // 
            // BarcodeFontTextBox
            // 
            BarcodeFontTextBox.Location = new Point(320, 30);
            BarcodeFontTextBox.MaxLength = 50;
            BarcodeFontTextBox.Name = "BarcodeFontTextBox";
            BarcodeFontTextBox.Size = new Size(135, 23);
            BarcodeFontTextBox.TabIndex = 24;
            // 
            // Label36
            // 
            Label36.AutoSize = true;
            Label36.Location = new Point(318, 15);
            Label36.Name = "Label36";
            Label36.Size = new Size(40, 15);
            Label36.TabIndex = 24;
            Label36.Text = "フォント";
            // 
            // BarcodeQuantityTextBox
            // 
            BarcodeQuantityTextBox.Location = new Point(104, 92);
            BarcodeQuantityTextBox.MaxLength = 5;
            BarcodeQuantityTextBox.Name = "BarcodeQuantityTextBox";
            BarcodeQuantityTextBox.Size = new Size(52, 23);
            BarcodeQuantityTextBox.TabIndex = 10;
            // 
            // Label35
            // 
            Label35.AutoSize = true;
            Label35.Location = new Point(101, 77);
            Label35.Name = "Label35";
            Label35.Size = new Size(55, 15);
            Label35.TabIndex = 9;
            Label35.Text = "発行枚数";
            // 
            // BarcodeMagnitudeTextBox
            // 
            BarcodeMagnitudeTextBox.Location = new Point(22, 92);
            BarcodeMagnitudeTextBox.MaxLength = 5;
            BarcodeMagnitudeTextBox.Name = "BarcodeMagnitudeTextBox";
            BarcodeMagnitudeTextBox.Size = new Size(52, 23);
            BarcodeMagnitudeTextBox.TabIndex = 8;
            // 
            // Label34
            // 
            Label34.AutoSize = true;
            Label34.Location = new Point(20, 77);
            Label34.Name = "Label34";
            Label34.Size = new Size(31, 15);
            Label34.TabIndex = 7;
            Label34.Text = "倍率";
            // 
            // BarcodeCenterCheckBox
            // 
            BarcodeCenterCheckBox.AutoSize = true;
            BarcodeCenterCheckBox.Location = new Point(22, 55);
            BarcodeCenterCheckBox.Name = "BarcodeCenterCheckBox";
            BarcodeCenterCheckBox.Size = new Size(145, 19);
            BarcodeCenterCheckBox.TabIndex = 6;
            BarcodeCenterCheckBox.Text = "横位置を中心に合わせる";
            BarcodeCenterCheckBox.UseVisualStyleBackColor = true;
            BarcodeCenterCheckBox.CheckedChanged += BarcodeCenterCheckBox_CheckedChanged;
            // 
            // BarcodePostionYTextBox
            // 
            BarcodePostionYTextBox.Location = new Point(194, 30);
            BarcodePostionYTextBox.MaxLength = 5;
            BarcodePostionYTextBox.Name = "BarcodePostionYTextBox";
            BarcodePostionYTextBox.Size = new Size(68, 23);
            BarcodePostionYTextBox.TabIndex = 5;
            // 
            // Label33
            // 
            Label33.AutoSize = true;
            Label33.Location = new Point(191, 15);
            Label33.Name = "Label33";
            Label33.Size = new Size(43, 15);
            Label33.TabIndex = 4;
            Label33.Text = "縦位置";
            // 
            // BarcodePostionXTextBox
            // 
            BarcodePostionXTextBox.Location = new Point(104, 30);
            BarcodePostionXTextBox.MaxLength = 5;
            BarcodePostionXTextBox.Name = "BarcodePostionXTextBox";
            BarcodePostionXTextBox.Size = new Size(68, 23);
            BarcodePostionXTextBox.TabIndex = 3;
            // 
            // Label32
            // 
            Label32.AutoSize = true;
            Label32.Location = new Point(101, 15);
            Label32.Name = "Label32";
            Label32.Size = new Size(43, 15);
            Label32.TabIndex = 2;
            Label32.Text = "横位置";
            // 
            // BarcodeHeightTextBox
            // 
            BarcodeHeightTextBox.Location = new Point(22, 30);
            BarcodeHeightTextBox.MaxLength = 5;
            BarcodeHeightTextBox.Name = "BarcodeHeightTextBox";
            BarcodeHeightTextBox.Size = new Size(52, 23);
            BarcodeHeightTextBox.TabIndex = 1;
            // 
            // Label31
            // 
            Label31.AutoSize = true;
            Label31.Location = new Point(20, 15);
            Label31.Name = "Label31";
            Label31.Size = new Size(27, 15);
            Label31.TabIndex = 1;
            Label31.Text = "高さ";
            // 
            // WhiteSpaceGroupBox
            // 
            WhiteSpaceGroupBox.Controls.Add(Label11);
            WhiteSpaceGroupBox.Controls.Add(Label10);
            WhiteSpaceGroupBox.Controls.Add(BarcodePageOffsetYTextBox);
            WhiteSpaceGroupBox.Controls.Add(BarcodePageOffsetXTextBox);
            WhiteSpaceGroupBox.Controls.Add(Label9);
            WhiteSpaceGroupBox.Controls.Add(Label8);
            WhiteSpaceGroupBox.Location = new Point(27, 158);
            WhiteSpaceGroupBox.Name = "WhiteSpaceGroupBox";
            WhiteSpaceGroupBox.Size = new Size(311, 67);
            WhiteSpaceGroupBox.TabIndex = 27;
            WhiteSpaceGroupBox.TabStop = false;
            WhiteSpaceGroupBox.Text = "用紙余白";
            // 
            // Label11
            // 
            Label11.AutoSize = true;
            Label11.Location = new Point(261, 39);
            Label11.Name = "Label11";
            Label11.Size = new Size(31, 15);
            Label11.TabIndex = 5;
            Label11.Text = "mm";
            // 
            // Label10
            // 
            Label10.AutoSize = true;
            Label10.Location = new Point(126, 39);
            Label10.Name = "Label10";
            Label10.Size = new Size(31, 15);
            Label10.TabIndex = 4;
            Label10.Text = "mm";
            // 
            // BarcodePageOffsetYTextBox
            // 
            BarcodePageOffsetYTextBox.Location = new Point(158, 36);
            BarcodePageOffsetYTextBox.MaxLength = 5;
            BarcodePageOffsetYTextBox.Name = "BarcodePageOffsetYTextBox";
            BarcodePageOffsetYTextBox.Size = new Size(100, 23);
            BarcodePageOffsetYTextBox.TabIndex = 3;
            // 
            // BarcodePageOffsetXTextBox
            // 
            BarcodePageOffsetXTextBox.Location = new Point(20, 36);
            BarcodePageOffsetXTextBox.MaxLength = 5;
            BarcodePageOffsetXTextBox.Name = "BarcodePageOffsetXTextBox";
            BarcodePageOffsetXTextBox.Size = new Size(100, 23);
            BarcodePageOffsetXTextBox.TabIndex = 1;
            // 
            // Label9
            // 
            Label9.AutoSize = true;
            Label9.Location = new Point(156, 21);
            Label9.Name = "Label9";
            Label9.Size = new Size(43, 15);
            Label9.TabIndex = 2;
            Label9.Text = "上余白";
            // 
            // Label8
            // 
            Label8.AutoSize = true;
            Label8.Location = new Point(18, 21);
            Label8.Name = "Label8";
            Label8.Size = new Size(43, 15);
            Label8.TabIndex = 1;
            Label8.Text = "左余白";
            // 
            // QuantityGroupBox
            // 
            QuantityGroupBox.Controls.Add(Label7);
            QuantityGroupBox.Controls.Add(BarcodeQuantityYTextBox);
            QuantityGroupBox.Controls.Add(BarcodeQuantityXTextBox);
            QuantityGroupBox.Controls.Add(Label6);
            QuantityGroupBox.Controls.Add(Label5);
            QuantityGroupBox.Location = new Point(27, 85);
            QuantityGroupBox.Name = "QuantityGroupBox";
            QuantityGroupBox.Size = new Size(311, 67);
            QuantityGroupBox.TabIndex = 28;
            QuantityGroupBox.TabStop = false;
            QuantityGroupBox.Text = "配置数";
            // 
            // Label7
            // 
            Label7.AutoSize = true;
            Label7.Location = new Point(132, 38);
            Label7.Name = "Label7";
            Label7.Size = new Size(17, 15);
            Label7.TabIndex = 4;
            Label7.Text = "×";
            // 
            // BarcodeQuantityYTextBox
            // 
            BarcodeQuantityYTextBox.Location = new Point(158, 36);
            BarcodeQuantityYTextBox.MaxLength = 5;
            BarcodeQuantityYTextBox.Name = "BarcodeQuantityYTextBox";
            BarcodeQuantityYTextBox.Size = new Size(100, 23);
            BarcodeQuantityYTextBox.TabIndex = 3;
            // 
            // BarcodeQuantityXTextBox
            // 
            BarcodeQuantityXTextBox.Location = new Point(20, 36);
            BarcodeQuantityXTextBox.MaxLength = 5;
            BarcodeQuantityXTextBox.Name = "BarcodeQuantityXTextBox";
            BarcodeQuantityXTextBox.Size = new Size(100, 23);
            BarcodeQuantityXTextBox.TabIndex = 1;
            // 
            // Label6
            // 
            Label6.AutoSize = true;
            Label6.Location = new Point(156, 21);
            Label6.Name = "Label6";
            Label6.Size = new Size(19, 15);
            Label6.TabIndex = 2;
            Label6.Text = "縦";
            // 
            // Label5
            // 
            Label5.AutoSize = true;
            Label5.Location = new Point(18, 21);
            Label5.Name = "Label5";
            Label5.Size = new Size(19, 15);
            Label5.TabIndex = 1;
            Label5.Text = "横";
            // 
            // BarcodeSizeGroupBox
            // 
            BarcodeSizeGroupBox.Controls.Add(Label4);
            BarcodeSizeGroupBox.Controls.Add(Label3);
            BarcodeSizeGroupBox.Controls.Add(BarcodeLabelHeightTextBox);
            BarcodeSizeGroupBox.Controls.Add(BarcodeLabelWidthTextBox);
            BarcodeSizeGroupBox.Controls.Add(Label2);
            BarcodeSizeGroupBox.Controls.Add(Label1);
            BarcodeSizeGroupBox.Location = new Point(27, 12);
            BarcodeSizeGroupBox.Name = "BarcodeSizeGroupBox";
            BarcodeSizeGroupBox.Size = new Size(311, 67);
            BarcodeSizeGroupBox.TabIndex = 26;
            BarcodeSizeGroupBox.TabStop = false;
            BarcodeSizeGroupBox.Text = "ラベルサイズ";
            // 
            // Label4
            // 
            Label4.AutoSize = true;
            Label4.Location = new Point(261, 39);
            Label4.Name = "Label4";
            Label4.Size = new Size(31, 15);
            Label4.TabIndex = 5;
            Label4.Text = "mm";
            // 
            // Label3
            // 
            Label3.AutoSize = true;
            Label3.Location = new Point(126, 39);
            Label3.Name = "Label3";
            Label3.Size = new Size(31, 15);
            Label3.TabIndex = 4;
            Label3.Text = "mm";
            // 
            // BarcodeLabelHeightTextBox
            // 
            BarcodeLabelHeightTextBox.Location = new Point(158, 36);
            BarcodeLabelHeightTextBox.MaxLength = 5;
            BarcodeLabelHeightTextBox.Name = "BarcodeLabelHeightTextBox";
            BarcodeLabelHeightTextBox.Size = new Size(100, 23);
            BarcodeLabelHeightTextBox.TabIndex = 3;
            // 
            // BarcodeLabelWidthTextBox
            // 
            BarcodeLabelWidthTextBox.Location = new Point(20, 36);
            BarcodeLabelWidthTextBox.MaxLength = 5;
            BarcodeLabelWidthTextBox.Name = "BarcodeLabelWidthTextBox";
            BarcodeLabelWidthTextBox.Size = new Size(100, 23);
            BarcodeLabelWidthTextBox.TabIndex = 1;
            // 
            // Label2
            // 
            Label2.AutoSize = true;
            Label2.Location = new Point(156, 21);
            Label2.Name = "Label2";
            Label2.Size = new Size(19, 15);
            Label2.TabIndex = 2;
            Label2.Text = "縦";
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Location = new Point(18, 21);
            Label1.Name = "Label1";
            Label1.Size = new Size(19, 15);
            Label1.TabIndex = 1;
            Label1.Text = "横";
            // 
            // HeaderFooterGroupBox
            // 
            HeaderFooterGroupBox.Controls.Add(BarcodeFooterStringTextBox);
            HeaderFooterGroupBox.Controls.Add(Label23);
            HeaderFooterGroupBox.Controls.Add(Label24);
            HeaderFooterGroupBox.Controls.Add(BarcodeFooterPostionYTextBox);
            HeaderFooterGroupBox.Controls.Add(BarcodeFooterPostionXTextBox);
            HeaderFooterGroupBox.Controls.Add(Label25);
            HeaderFooterGroupBox.Controls.Add(Label26);
            HeaderFooterGroupBox.Controls.Add(Label27);
            HeaderFooterGroupBox.Controls.Add(Label30);
            HeaderFooterGroupBox.Controls.Add(Label29);
            HeaderFooterGroupBox.Controls.Add(BarcodeHeaderStringTextBox);
            HeaderFooterGroupBox.Controls.Add(Label22);
            HeaderFooterGroupBox.Controls.Add(Label21);
            HeaderFooterGroupBox.Controls.Add(Label20);
            HeaderFooterGroupBox.Controls.Add(BarcodeHeaderPostionYTextBox);
            HeaderFooterGroupBox.Controls.Add(BarcodeHeaderPostionXTextBox);
            HeaderFooterGroupBox.Controls.Add(Label19);
            HeaderFooterGroupBox.Controls.Add(Label18);
            HeaderFooterGroupBox.Controls.Add(Label17);
            HeaderFooterGroupBox.Controls.Add(BarcodeHeaderFooterFontButton);
            HeaderFooterGroupBox.Controls.Add(BarcodeHeaderFooterFontTextBox);
            HeaderFooterGroupBox.Controls.Add(Label16);
            HeaderFooterGroupBox.Location = new Point(357, 12);
            HeaderFooterGroupBox.Name = "HeaderFooterGroupBox";
            HeaderFooterGroupBox.Size = new Size(270, 291);
            HeaderFooterGroupBox.TabIndex = 29;
            HeaderFooterGroupBox.TabStop = false;
            HeaderFooterGroupBox.Text = "ヘッダ・フッタ";
            // 
            // BarcodeFooterStringTextBox
            // 
            BarcodeFooterStringTextBox.Location = new Point(79, 215);
            BarcodeFooterStringTextBox.MaxLength = 50;
            BarcodeFooterStringTextBox.Name = "BarcodeFooterStringTextBox";
            BarcodeFooterStringTextBox.Size = new Size(185, 23);
            BarcodeFooterStringTextBox.TabIndex = 31;
            // 
            // Label23
            // 
            Label23.AutoSize = true;
            Label23.Location = new Point(6, 218);
            Label23.Name = "Label23";
            Label23.Size = new Size(69, 15);
            Label23.TabIndex = 30;
            Label23.Text = "ヘッダ文字列";
            // 
            // Label24
            // 
            Label24.AutoSize = true;
            Label24.Location = new Point(106, 192);
            Label24.Name = "Label24";
            Label24.Size = new Size(31, 15);
            Label24.TabIndex = 29;
            Label24.Text = "mm";
            // 
            // BarcodeFooterPostionYTextBox
            // 
            BarcodeFooterPostionYTextBox.Location = new Point(168, 189);
            BarcodeFooterPostionYTextBox.MaxLength = 50;
            BarcodeFooterPostionYTextBox.Name = "BarcodeFooterPostionYTextBox";
            BarcodeFooterPostionYTextBox.Size = new Size(64, 23);
            BarcodeFooterPostionYTextBox.TabIndex = 28;
            // 
            // BarcodeFooterPostionXTextBox
            // 
            BarcodeFooterPostionXTextBox.Location = new Point(36, 189);
            BarcodeFooterPostionXTextBox.MaxLength = 50;
            BarcodeFooterPostionXTextBox.Name = "BarcodeFooterPostionXTextBox";
            BarcodeFooterPostionXTextBox.Size = new Size(64, 23);
            BarcodeFooterPostionXTextBox.TabIndex = 27;
            // 
            // Label25
            // 
            Label25.AutoSize = true;
            Label25.Location = new Point(166, 174);
            Label25.Name = "Label25";
            Label25.Size = new Size(19, 15);
            Label25.TabIndex = 26;
            Label25.Text = "縦";
            // 
            // Label26
            // 
            Label26.AutoSize = true;
            Label26.Location = new Point(34, 174);
            Label26.Name = "Label26";
            Label26.Size = new Size(19, 15);
            Label26.TabIndex = 25;
            Label26.Text = "横";
            // 
            // Label27
            // 
            Label27.AutoSize = true;
            Label27.Location = new Point(23, 154);
            Label27.Name = "Label27";
            Label27.Size = new Size(55, 15);
            Label27.TabIndex = 24;
            Label27.Text = "フッダ位置";
            // 
            // Label30
            // 
            Label30.AutoSize = true;
            Label30.Location = new Point(7, 265);
            Label30.Name = "Label30";
            Label30.Size = new Size(188, 15);
            Label30.TabIndex = 23;
            Label30.Text = "%P 製品名 %N 台数  %U 担当者";
            // 
            // Label29
            // 
            Label29.AutoSize = true;
            Label29.Location = new Point(7, 250);
            Label29.Name = "Label29";
            Label29.Size = new Size(227, 15);
            Label29.TabIndex = 22;
            Label29.Text = "%D 印刷日 %M 製番 %O 注番 %T 型式";
            // 
            // BarcodeHeaderStringTextBox
            // 
            BarcodeHeaderStringTextBox.Location = new Point(79, 125);
            BarcodeHeaderStringTextBox.MaxLength = 50;
            BarcodeHeaderStringTextBox.Name = "BarcodeHeaderStringTextBox";
            BarcodeHeaderStringTextBox.Size = new Size(185, 23);
            BarcodeHeaderStringTextBox.TabIndex = 12;
            // 
            // Label22
            // 
            Label22.AutoSize = true;
            Label22.Location = new Point(6, 128);
            Label22.Name = "Label22";
            Label22.Size = new Size(69, 15);
            Label22.TabIndex = 11;
            Label22.Text = "ヘッダ文字列";
            // 
            // Label21
            // 
            Label21.AutoSize = true;
            Label21.Location = new Point(238, 102);
            Label21.Name = "Label21";
            Label21.Size = new Size(31, 15);
            Label21.TabIndex = 10;
            Label21.Text = "mm";
            // 
            // Label20
            // 
            Label20.AutoSize = true;
            Label20.Location = new Point(106, 102);
            Label20.Name = "Label20";
            Label20.Size = new Size(31, 15);
            Label20.TabIndex = 9;
            Label20.Text = "mm";
            // 
            // BarcodeHeaderPostionYTextBox
            // 
            BarcodeHeaderPostionYTextBox.Location = new Point(168, 99);
            BarcodeHeaderPostionYTextBox.MaxLength = 50;
            BarcodeHeaderPostionYTextBox.Name = "BarcodeHeaderPostionYTextBox";
            BarcodeHeaderPostionYTextBox.Size = new Size(64, 23);
            BarcodeHeaderPostionYTextBox.TabIndex = 8;
            // 
            // BarcodeHeaderPostionXTextBox
            // 
            BarcodeHeaderPostionXTextBox.Location = new Point(36, 99);
            BarcodeHeaderPostionXTextBox.MaxLength = 50;
            BarcodeHeaderPostionXTextBox.Name = "BarcodeHeaderPostionXTextBox";
            BarcodeHeaderPostionXTextBox.Size = new Size(64, 23);
            BarcodeHeaderPostionXTextBox.TabIndex = 7;
            // 
            // Label19
            // 
            Label19.AutoSize = true;
            Label19.Location = new Point(166, 84);
            Label19.Name = "Label19";
            Label19.Size = new Size(19, 15);
            Label19.TabIndex = 6;
            Label19.Text = "縦";
            // 
            // Label18
            // 
            Label18.AutoSize = true;
            Label18.Location = new Point(34, 84);
            Label18.Name = "Label18";
            Label18.Size = new Size(19, 15);
            Label18.TabIndex = 5;
            Label18.Text = "横";
            // 
            // Label17
            // 
            Label17.AutoSize = true;
            Label17.Location = new Point(23, 64);
            Label17.Name = "Label17";
            Label17.Size = new Size(57, 15);
            Label17.TabIndex = 4;
            Label17.Text = "ヘッダ位置";
            // 
            // BarcodeHeaderFooterFontButton
            // 
            BarcodeHeaderFooterFontButton.Location = new Point(168, 35);
            BarcodeHeaderFooterFontButton.Name = "BarcodeHeaderFooterFontButton";
            BarcodeHeaderFooterFontButton.Size = new Size(96, 25);
            BarcodeHeaderFooterFontButton.TabIndex = 3;
            BarcodeHeaderFooterFontButton.Text = "フォント選択";
            BarcodeHeaderFooterFontButton.UseVisualStyleBackColor = true;
            BarcodeHeaderFooterFontButton.Click += BtnHeaderFooterFont_Click;
            // 
            // BarcodeHeaderFooterFontTextBox
            // 
            BarcodeHeaderFooterFontTextBox.Location = new Point(27, 36);
            BarcodeHeaderFooterFontTextBox.MaxLength = 50;
            BarcodeHeaderFooterFontTextBox.Name = "BarcodeHeaderFooterFontTextBox";
            BarcodeHeaderFooterFontTextBox.Size = new Size(135, 23);
            BarcodeHeaderFooterFontTextBox.TabIndex = 2;
            // 
            // Label16
            // 
            Label16.AutoSize = true;
            Label16.Location = new Point(25, 21);
            Label16.Name = "Label16";
            Label16.Size = new Size(40, 15);
            Label16.TabIndex = 1;
            Label16.Text = "フォント";
            // 
            // LabelIntervalGroupBox
            // 
            LabelIntervalGroupBox.Controls.Add(Label15);
            LabelIntervalGroupBox.Controls.Add(Label14);
            LabelIntervalGroupBox.Controls.Add(BarcodeLabelIntervalYTextBox);
            LabelIntervalGroupBox.Controls.Add(BarcodeLabelIntervalXTextBox);
            LabelIntervalGroupBox.Controls.Add(Label13);
            LabelIntervalGroupBox.Controls.Add(Label12);
            LabelIntervalGroupBox.Location = new Point(27, 231);
            LabelIntervalGroupBox.Name = "LabelIntervalGroupBox";
            LabelIntervalGroupBox.Size = new Size(311, 67);
            LabelIntervalGroupBox.TabIndex = 31;
            LabelIntervalGroupBox.TabStop = false;
            LabelIntervalGroupBox.Text = "ラベル間隔";
            // 
            // Label15
            // 
            Label15.AutoSize = true;
            Label15.Location = new Point(261, 39);
            Label15.Name = "Label15";
            Label15.Size = new Size(31, 15);
            Label15.TabIndex = 5;
            Label15.Text = "mm";
            // 
            // Label14
            // 
            Label14.AutoSize = true;
            Label14.Location = new Point(126, 39);
            Label14.Name = "Label14";
            Label14.Size = new Size(31, 15);
            Label14.TabIndex = 4;
            Label14.Text = "mm";
            // 
            // BarcodeLabelIntervalYTextBox
            // 
            BarcodeLabelIntervalYTextBox.Location = new Point(158, 36);
            BarcodeLabelIntervalYTextBox.MaxLength = 5;
            BarcodeLabelIntervalYTextBox.Name = "BarcodeLabelIntervalYTextBox";
            BarcodeLabelIntervalYTextBox.Size = new Size(100, 23);
            BarcodeLabelIntervalYTextBox.TabIndex = 3;
            // 
            // BarcodeLabelIntervalXTextBox
            // 
            BarcodeLabelIntervalXTextBox.Location = new Point(20, 36);
            BarcodeLabelIntervalXTextBox.MaxLength = 5;
            BarcodeLabelIntervalXTextBox.Name = "BarcodeLabelIntervalXTextBox";
            BarcodeLabelIntervalXTextBox.Size = new Size(100, 23);
            BarcodeLabelIntervalXTextBox.TabIndex = 1;
            // 
            // Label13
            // 
            Label13.AutoSize = true;
            Label13.Location = new Point(156, 21);
            Label13.Name = "Label13";
            Label13.Size = new Size(55, 15);
            Label13.TabIndex = 2;
            Label13.Text = "上下間隔";
            // 
            // Label12
            // 
            Label12.AutoSize = true;
            Label12.Location = new Point(18, 21);
            Label12.Name = "Label12";
            Label12.Size = new Size(55, 15);
            Label12.TabIndex = 1;
            Label12.Text = "左右間隔";
            // 
            // CloseButton
            // 
            CloseButton.Location = new Point(552, 524);
            CloseButton.Name = "CloseButton";
            CloseButton.Size = new Size(75, 25);
            CloseButton.TabIndex = 33;
            CloseButton.Text = "キャンセル";
            CloseButton.UseVisualStyleBackColor = true;
            CloseButton.Click += CloseButton_Click;
            // 
            // ProductBarcodePrintSettingsWindow
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            CancelButton = CloseButton;
            ClientSize = new Size(654, 561);
            Controls.Add(OKButton);
            Controls.Add(BarcodeGroupBox);
            Controls.Add(CloseButton);
            Controls.Add(WhiteSpaceGroupBox);
            Controls.Add(QuantityGroupBox);
            Controls.Add(BarcodeSizeGroupBox);
            Controls.Add(HeaderFooterGroupBox);
            Controls.Add(LabelIntervalGroupBox);
            Font = new Font("Meiryo UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "ProductBarcodePrintSettingsWindow";
            StartPosition = FormStartPosition.CenterParent;
            Text = "製品バーコード設定";
            Load += ProductBarcodePrintSetting_Load;
            BarcodeGroupBox.ResumeLayout(false);
            BarcodeGroupBox.PerformLayout();
            WhiteSpaceGroupBox.ResumeLayout(false);
            WhiteSpaceGroupBox.PerformLayout();
            QuantityGroupBox.ResumeLayout(false);
            QuantityGroupBox.PerformLayout();
            BarcodeSizeGroupBox.ResumeLayout(false);
            BarcodeSizeGroupBox.PerformLayout();
            HeaderFooterGroupBox.ResumeLayout(false);
            HeaderFooterGroupBox.PerformLayout();
            LabelIntervalGroupBox.ResumeLayout(false);
            LabelIntervalGroupBox.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Button OKButton;
        private GroupBox BarcodeGroupBox;
        private Label Label41;
        private Label Label40;
        private TextBox BarcodeFormatTextBox;
        private Label Label39;
        private Button BarcodeFontButton;
        private TextBox BarcodeFontTextBox;
        private Label Label36;
        private TextBox BarcodeQuantityTextBox;
        private Label Label35;
        private TextBox BarcodeMagnitudeTextBox;
        private Label Label34;
        private CheckBox BarcodeCenterCheckBox;
        private TextBox BarcodePostionYTextBox;
        private Label Label33;
        private TextBox BarcodePostionXTextBox;
        private Label Label32;
        private TextBox BarcodeHeightTextBox;
        private Label Label31;
        private GroupBox WhiteSpaceGroupBox;
        private Label Label11;
        private Label Label10;
        private TextBox BarcodePageOffsetYTextBox;
        private TextBox BarcodePageOffsetXTextBox;
        private Label Label9;
        private Label Label8;
        private GroupBox QuantityGroupBox;
        private Label Label7;
        private TextBox BarcodeQuantityYTextBox;
        private TextBox BarcodeQuantityXTextBox;
        private Label Label6;
        private Label Label5;
        private GroupBox BarcodeSizeGroupBox;
        private Label Label4;
        private Label Label3;
        private TextBox BarcodeLabelHeightTextBox;
        private TextBox BarcodeLabelWidthTextBox;
        private Label Label2;
        private Label Label1;
        private GroupBox HeaderFooterGroupBox;
        private Label Label30;
        private Label Label29;
        private TextBox BarcodeHeaderStringTextBox;
        private Label Label22;
        private Label Label21;
        private Label Label20;
        private TextBox BarcodeHeaderPostionYTextBox;
        private TextBox BarcodeHeaderPostionXTextBox;
        private Label Label19;
        private Label Label18;
        private Label Label17;
        private Button BarcodeHeaderFooterFontButton;
        private TextBox BarcodeHeaderFooterFontTextBox;
        private Label Label16;
        private GroupBox LabelIntervalGroupBox;
        private Label Label15;
        private Label Label14;
        private TextBox BarcodeLabelIntervalYTextBox;
        private TextBox BarcodeLabelIntervalXTextBox;
        private Label Label13;
        private Label Label12;
        private TextBox BarcodeFooterStringTextBox;
        private Label Label23;
        private Label Label24;
        private TextBox BarcodeFooterPostionYTextBox;
        private TextBox BarcodeFooterPostionXTextBox;
        private Label Label25;
        private Label Label26;
        private Label Label27;
        private CheckBox FontCenterCheckBox;
        private TextBox FontPostionYTextBox;
        private TextBox FontPostionXTextBox;
        private Label Label37;
        private Label Label28;
        private Button CloseButton;
        private FontDialog BarcodeHeaderFontDialog;
        private FontDialog BarcodeFontDialog;
    }
}