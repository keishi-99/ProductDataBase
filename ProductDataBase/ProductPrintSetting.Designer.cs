namespace ProductDataBase {
    partial class ProductPrintSetting {
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
            PrintTextGroupBox = new GroupBox();
            Label41 = new Label();
            Label40 = new Label();
            PrintTextFormatTextBox = new TextBox();
            Label39 = new Label();
            PrintTextFontButton = new Button();
            PrintTextFontTextBox = new TextBox();
            Label36 = new Label();
            PrintTextQuantityTextBox = new TextBox();
            Label35 = new Label();
            PrintTextMagnitudeTextBox = new TextBox();
            Label34 = new Label();
            BarcodeCenterCheckBox = new CheckBox();
            PrintTextPostionYTextBox = new TextBox();
            Label33 = new Label();
            PrintTextPostionXTextBox = new TextBox();
            Label32 = new Label();
            PrintTextHeightTextBox = new TextBox();
            Label31 = new Label();
            WhiteSpaceGroupBox = new GroupBox();
            Label11 = new Label();
            Label10 = new Label();
            PageOffsetYTextBox = new TextBox();
            PageOffsetXTextBox = new TextBox();
            Label9 = new Label();
            Label8 = new Label();
            QuantityGroupBox = new GroupBox();
            Label7 = new Label();
            QuantityYTextBox = new TextBox();
            QuantityXTextBox = new TextBox();
            Label6 = new Label();
            Label5 = new Label();
            LabelSizeGroupBox = new GroupBox();
            Label4 = new Label();
            Label3 = new Label();
            LabelHeightTextBox = new TextBox();
            LabelWidthTextBox = new TextBox();
            Label2 = new Label();
            Label1 = new Label();
            HeaderFooterGroupBox = new GroupBox();
            Label30 = new Label();
            Label29 = new Label();
            HeaderStringTextBox = new TextBox();
            Label22 = new Label();
            Label21 = new Label();
            Label20 = new Label();
            HeaderPostionYTextBox = new TextBox();
            HeaderPostionXTextBox = new TextBox();
            Label19 = new Label();
            Label18 = new Label();
            Label17 = new Label();
            HeaderFooterFontButton = new Button();
            HeaderFooterFontTextBox = new TextBox();
            Label16 = new Label();
            LabelIntervalGroupBox = new GroupBox();
            Label15 = new Label();
            Label14 = new Label();
            LabelIntervalYTextBox = new TextBox();
            LabelIntervalXTextBox = new TextBox();
            Label13 = new Label();
            Label12 = new Label();
            CloseButton = new Button();
            HeaderFontDialog = new FontDialog();
            TextFontDialog = new FontDialog();
            PrintTextGroupBox.SuspendLayout();
            WhiteSpaceGroupBox.SuspendLayout();
            QuantityGroupBox.SuspendLayout();
            LabelSizeGroupBox.SuspendLayout();
            HeaderFooterGroupBox.SuspendLayout();
            LabelIntervalGroupBox.SuspendLayout();
            SuspendLayout();
            // 
            // OKButton
            // 
            OKButton.Location = new Point(471, 524);
            OKButton.Name = "OKButton";
            OKButton.Size = new Size(75, 25);
            OKButton.TabIndex = 24;
            OKButton.Text = "OK";
            OKButton.UseVisualStyleBackColor = true;
            OKButton.Click += BtnOK_Click;
            // 
            // PrintTextGroupBox
            // 
            PrintTextGroupBox.Controls.Add(Label41);
            PrintTextGroupBox.Controls.Add(Label40);
            PrintTextGroupBox.Controls.Add(PrintTextFormatTextBox);
            PrintTextGroupBox.Controls.Add(Label39);
            PrintTextGroupBox.Controls.Add(PrintTextFontButton);
            PrintTextGroupBox.Controls.Add(PrintTextFontTextBox);
            PrintTextGroupBox.Controls.Add(Label36);
            PrintTextGroupBox.Controls.Add(PrintTextQuantityTextBox);
            PrintTextGroupBox.Controls.Add(Label35);
            PrintTextGroupBox.Controls.Add(PrintTextMagnitudeTextBox);
            PrintTextGroupBox.Controls.Add(Label34);
            PrintTextGroupBox.Controls.Add(BarcodeCenterCheckBox);
            PrintTextGroupBox.Controls.Add(PrintTextPostionYTextBox);
            PrintTextGroupBox.Controls.Add(Label33);
            PrintTextGroupBox.Controls.Add(PrintTextPostionXTextBox);
            PrintTextGroupBox.Controls.Add(Label32);
            PrintTextGroupBox.Controls.Add(PrintTextHeightTextBox);
            PrintTextGroupBox.Controls.Add(Label31);
            PrintTextGroupBox.Location = new Point(27, 309);
            PrintTextGroupBox.Name = "PrintTextGroupBox";
            PrintTextGroupBox.Size = new Size(600, 170);
            PrintTextGroupBox.TabIndex = 22;
            PrintTextGroupBox.TabStop = false;
            PrintTextGroupBox.Text = "印刷文字設定";
            // 
            // Label41
            // 
            Label41.AutoSize = true;
            Label41.Location = new Point(318, 124);
            Label41.Name = "Label41";
            Label41.Size = new Size(281, 15);
            Label41.TabIndex = 30;
            Label41.Text = "%Y: 製造年 %MM: 製造月(2桁) %M: 製造月(1桁)";
            // 
            // Label40
            // 
            Label40.AutoSize = true;
            Label40.Location = new Point(318, 109);
            Label40.Name = "Label40";
            Label40.Size = new Size(201, 15);
            Label40.TabIndex = 24;
            Label40.Text = "%T: 型式 %R: リビジョン %S: シリアル";
            // 
            // PrintTextFormatTextBox
            // 
            PrintTextFormatTextBox.Location = new Point(379, 74);
            PrintTextFormatTextBox.MaxLength = 50;
            PrintTextFormatTextBox.Name = "PrintTextFormatTextBox";
            PrintTextFormatTextBox.Size = new Size(180, 23);
            PrintTextFormatTextBox.TabIndex = 24;
            // 
            // Label39
            // 
            Label39.AutoSize = true;
            Label39.Location = new Point(318, 77);
            Label39.Name = "Label39";
            Label39.Size = new Size(57, 15);
            Label39.TabIndex = 24;
            Label39.Text = "フォーマット";
            // 
            // PrintTextFontButton
            // 
            PrintTextFontButton.Location = new Point(479, 29);
            PrintTextFontButton.Name = "PrintTextFontButton";
            PrintTextFontButton.Size = new Size(96, 25);
            PrintTextFontButton.TabIndex = 24;
            PrintTextFontButton.Text = "フォント選択";
            PrintTextFontButton.UseVisualStyleBackColor = true;
            PrintTextFontButton.Click += BtnBarcodeFont_Click;
            // 
            // PrintTextFontTextBox
            // 
            PrintTextFontTextBox.Location = new Point(320, 30);
            PrintTextFontTextBox.MaxLength = 50;
            PrintTextFontTextBox.Name = "PrintTextFontTextBox";
            PrintTextFontTextBox.Size = new Size(135, 23);
            PrintTextFontTextBox.TabIndex = 24;
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
            // PrintTextQuantityTextBox
            // 
            PrintTextQuantityTextBox.Location = new Point(104, 92);
            PrintTextQuantityTextBox.MaxLength = 5;
            PrintTextQuantityTextBox.Name = "PrintTextQuantityTextBox";
            PrintTextQuantityTextBox.Size = new Size(52, 23);
            PrintTextQuantityTextBox.TabIndex = 10;
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
            // PrintTextMagnitudeTextBox
            // 
            PrintTextMagnitudeTextBox.Location = new Point(22, 92);
            PrintTextMagnitudeTextBox.MaxLength = 5;
            PrintTextMagnitudeTextBox.Name = "PrintTextMagnitudeTextBox";
            PrintTextMagnitudeTextBox.Size = new Size(52, 23);
            PrintTextMagnitudeTextBox.TabIndex = 8;
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
            // PrintTextPostionYTextBox
            // 
            PrintTextPostionYTextBox.Location = new Point(194, 30);
            PrintTextPostionYTextBox.MaxLength = 5;
            PrintTextPostionYTextBox.Name = "PrintTextPostionYTextBox";
            PrintTextPostionYTextBox.Size = new Size(68, 23);
            PrintTextPostionYTextBox.TabIndex = 5;
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
            // PrintTextPostionXTextBox
            // 
            PrintTextPostionXTextBox.Location = new Point(104, 30);
            PrintTextPostionXTextBox.MaxLength = 5;
            PrintTextPostionXTextBox.Name = "PrintTextPostionXTextBox";
            PrintTextPostionXTextBox.Size = new Size(68, 23);
            PrintTextPostionXTextBox.TabIndex = 3;
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
            // PrintTextHeightTextBox
            // 
            PrintTextHeightTextBox.Location = new Point(22, 30);
            PrintTextHeightTextBox.MaxLength = 5;
            PrintTextHeightTextBox.Name = "PrintTextHeightTextBox";
            PrintTextHeightTextBox.Size = new Size(52, 23);
            PrintTextHeightTextBox.TabIndex = 1;
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
            WhiteSpaceGroupBox.Controls.Add(PageOffsetYTextBox);
            WhiteSpaceGroupBox.Controls.Add(PageOffsetXTextBox);
            WhiteSpaceGroupBox.Controls.Add(Label9);
            WhiteSpaceGroupBox.Controls.Add(Label8);
            WhiteSpaceGroupBox.Location = new Point(27, 158);
            WhiteSpaceGroupBox.Name = "WhiteSpaceGroupBox";
            WhiteSpaceGroupBox.Size = new Size(311, 67);
            WhiteSpaceGroupBox.TabIndex = 19;
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
            // PageOffsetYTextBox
            // 
            PageOffsetYTextBox.Location = new Point(158, 36);
            PageOffsetYTextBox.MaxLength = 5;
            PageOffsetYTextBox.Name = "PageOffsetYTextBox";
            PageOffsetYTextBox.Size = new Size(100, 23);
            PageOffsetYTextBox.TabIndex = 3;
            // 
            // PageOffsetXTextBox
            // 
            PageOffsetXTextBox.Location = new Point(20, 36);
            PageOffsetXTextBox.MaxLength = 5;
            PageOffsetXTextBox.Name = "PageOffsetXTextBox";
            PageOffsetXTextBox.Size = new Size(100, 23);
            PageOffsetXTextBox.TabIndex = 1;
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
            QuantityGroupBox.Controls.Add(QuantityYTextBox);
            QuantityGroupBox.Controls.Add(QuantityXTextBox);
            QuantityGroupBox.Controls.Add(Label6);
            QuantityGroupBox.Controls.Add(Label5);
            QuantityGroupBox.Location = new Point(27, 85);
            QuantityGroupBox.Name = "QuantityGroupBox";
            QuantityGroupBox.Size = new Size(311, 67);
            QuantityGroupBox.TabIndex = 20;
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
            // QuantityYTextBox
            // 
            QuantityYTextBox.Location = new Point(158, 36);
            QuantityYTextBox.MaxLength = 5;
            QuantityYTextBox.Name = "QuantityYTextBox";
            QuantityYTextBox.Size = new Size(100, 23);
            QuantityYTextBox.TabIndex = 3;
            // 
            // QuantityXTextBox
            // 
            QuantityXTextBox.Location = new Point(20, 36);
            QuantityXTextBox.MaxLength = 5;
            QuantityXTextBox.Name = "QuantityXTextBox";
            QuantityXTextBox.Size = new Size(100, 23);
            QuantityXTextBox.TabIndex = 1;
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
            // LabelSizeGroupBox
            // 
            LabelSizeGroupBox.Controls.Add(Label4);
            LabelSizeGroupBox.Controls.Add(Label3);
            LabelSizeGroupBox.Controls.Add(LabelHeightTextBox);
            LabelSizeGroupBox.Controls.Add(LabelWidthTextBox);
            LabelSizeGroupBox.Controls.Add(Label2);
            LabelSizeGroupBox.Controls.Add(Label1);
            LabelSizeGroupBox.Location = new Point(27, 12);
            LabelSizeGroupBox.Name = "LabelSizeGroupBox";
            LabelSizeGroupBox.Size = new Size(311, 67);
            LabelSizeGroupBox.TabIndex = 18;
            LabelSizeGroupBox.TabStop = false;
            LabelSizeGroupBox.Text = "ラベルサイズ";
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
            // LabelHeightTextBox
            // 
            LabelHeightTextBox.Location = new Point(158, 36);
            LabelHeightTextBox.MaxLength = 5;
            LabelHeightTextBox.Name = "LabelHeightTextBox";
            LabelHeightTextBox.Size = new Size(100, 23);
            LabelHeightTextBox.TabIndex = 3;
            // 
            // LabelWidthTextBox
            // 
            LabelWidthTextBox.Location = new Point(20, 36);
            LabelWidthTextBox.MaxLength = 5;
            LabelWidthTextBox.Name = "LabelWidthTextBox";
            LabelWidthTextBox.Size = new Size(100, 23);
            LabelWidthTextBox.TabIndex = 1;
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
            HeaderFooterGroupBox.Controls.Add(Label30);
            HeaderFooterGroupBox.Controls.Add(Label29);
            HeaderFooterGroupBox.Controls.Add(HeaderStringTextBox);
            HeaderFooterGroupBox.Controls.Add(Label22);
            HeaderFooterGroupBox.Controls.Add(Label21);
            HeaderFooterGroupBox.Controls.Add(Label20);
            HeaderFooterGroupBox.Controls.Add(HeaderPostionYTextBox);
            HeaderFooterGroupBox.Controls.Add(HeaderPostionXTextBox);
            HeaderFooterGroupBox.Controls.Add(Label19);
            HeaderFooterGroupBox.Controls.Add(Label18);
            HeaderFooterGroupBox.Controls.Add(Label17);
            HeaderFooterGroupBox.Controls.Add(HeaderFooterFontButton);
            HeaderFooterGroupBox.Controls.Add(HeaderFooterFontTextBox);
            HeaderFooterGroupBox.Controls.Add(Label16);
            HeaderFooterGroupBox.Location = new Point(357, 12);
            HeaderFooterGroupBox.Name = "HeaderFooterGroupBox";
            HeaderFooterGroupBox.Size = new Size(270, 213);
            HeaderFooterGroupBox.TabIndex = 21;
            HeaderFooterGroupBox.TabStop = false;
            HeaderFooterGroupBox.Text = "ヘッダ・フッタ";
            // 
            // Label30
            // 
            Label30.AutoSize = true;
            Label30.Location = new Point(6, 186);
            Label30.Name = "Label30";
            Label30.Size = new Size(188, 15);
            Label30.TabIndex = 23;
            Label30.Text = "%P 製品名 %N 台数  %U 担当者";
            // 
            // Label29
            // 
            Label29.AutoSize = true;
            Label29.Location = new Point(6, 171);
            Label29.Name = "Label29";
            Label29.Size = new Size(227, 15);
            Label29.TabIndex = 22;
            Label29.Text = "%D 印刷日 %M 製番 %O 注番 %T 型式";
            // 
            // HeaderStringTextBox
            // 
            HeaderStringTextBox.Location = new Point(79, 125);
            HeaderStringTextBox.MaxLength = 50;
            HeaderStringTextBox.Name = "HeaderStringTextBox";
            HeaderStringTextBox.Size = new Size(185, 23);
            HeaderStringTextBox.TabIndex = 12;
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
            // HeaderPostionYTextBox
            // 
            HeaderPostionYTextBox.Location = new Point(168, 99);
            HeaderPostionYTextBox.MaxLength = 50;
            HeaderPostionYTextBox.Name = "HeaderPostionYTextBox";
            HeaderPostionYTextBox.Size = new Size(64, 23);
            HeaderPostionYTextBox.TabIndex = 8;
            // 
            // HeaderPostionXTextBox
            // 
            HeaderPostionXTextBox.Location = new Point(36, 99);
            HeaderPostionXTextBox.MaxLength = 50;
            HeaderPostionXTextBox.Name = "HeaderPostionXTextBox";
            HeaderPostionXTextBox.Size = new Size(64, 23);
            HeaderPostionXTextBox.TabIndex = 7;
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
            // HeaderFooterFontButton
            // 
            HeaderFooterFontButton.Location = new Point(168, 35);
            HeaderFooterFontButton.Name = "HeaderFooterFontButton";
            HeaderFooterFontButton.Size = new Size(96, 25);
            HeaderFooterFontButton.TabIndex = 3;
            HeaderFooterFontButton.Text = "フォント選択";
            HeaderFooterFontButton.UseVisualStyleBackColor = true;
            HeaderFooterFontButton.Click += BtnHeaderFooterFont_Click;
            // 
            // HeaderFooterFontTextBox
            // 
            HeaderFooterFontTextBox.Location = new Point(27, 36);
            HeaderFooterFontTextBox.MaxLength = 50;
            HeaderFooterFontTextBox.Name = "HeaderFooterFontTextBox";
            HeaderFooterFontTextBox.Size = new Size(135, 23);
            HeaderFooterFontTextBox.TabIndex = 2;
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
            LabelIntervalGroupBox.Controls.Add(LabelIntervalYTextBox);
            LabelIntervalGroupBox.Controls.Add(LabelIntervalXTextBox);
            LabelIntervalGroupBox.Controls.Add(Label13);
            LabelIntervalGroupBox.Controls.Add(Label12);
            LabelIntervalGroupBox.Location = new Point(27, 231);
            LabelIntervalGroupBox.Name = "LabelIntervalGroupBox";
            LabelIntervalGroupBox.Size = new Size(311, 67);
            LabelIntervalGroupBox.TabIndex = 23;
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
            // LabelIntervalYTextBox
            // 
            LabelIntervalYTextBox.Location = new Point(158, 36);
            LabelIntervalYTextBox.MaxLength = 5;
            LabelIntervalYTextBox.Name = "LabelIntervalYTextBox";
            LabelIntervalYTextBox.Size = new Size(100, 23);
            LabelIntervalYTextBox.TabIndex = 3;
            // 
            // LabelIntervalXTextBox
            // 
            LabelIntervalXTextBox.Location = new Point(20, 36);
            LabelIntervalXTextBox.MaxLength = 5;
            LabelIntervalXTextBox.Name = "LabelIntervalXTextBox";
            LabelIntervalXTextBox.Size = new Size(100, 23);
            LabelIntervalXTextBox.TabIndex = 1;
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
            CloseButton.TabIndex = 25;
            CloseButton.Text = "Close";
            CloseButton.UseVisualStyleBackColor = true;
            CloseButton.Click += CloseButton_Click;
            // 
            // ProductPrintSetting
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(654, 561);
            Controls.Add(CloseButton);
            Controls.Add(OKButton);
            Controls.Add(PrintTextGroupBox);
            Controls.Add(WhiteSpaceGroupBox);
            Controls.Add(QuantityGroupBox);
            Controls.Add(LabelSizeGroupBox);
            Controls.Add(HeaderFooterGroupBox);
            Controls.Add(LabelIntervalGroupBox);
            Font = new Font("Meiryo UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "ProductPrintSetting";
            ShowIcon = false;
            ShowInTaskbar = false;
            Text = "製品ラベル設定";
            Load += ProductPrintSetting_Load;
            PrintTextGroupBox.ResumeLayout(false);
            PrintTextGroupBox.PerformLayout();
            WhiteSpaceGroupBox.ResumeLayout(false);
            WhiteSpaceGroupBox.PerformLayout();
            QuantityGroupBox.ResumeLayout(false);
            QuantityGroupBox.PerformLayout();
            LabelSizeGroupBox.ResumeLayout(false);
            LabelSizeGroupBox.PerformLayout();
            HeaderFooterGroupBox.ResumeLayout(false);
            HeaderFooterGroupBox.PerformLayout();
            LabelIntervalGroupBox.ResumeLayout(false);
            LabelIntervalGroupBox.PerformLayout();
            ResumeLayout(false);
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
        private TextBox PrintTextMagnitudeTextBox;
        private Label Label34;
        private CheckBox BarcodeCenterCheckBox;
        private TextBox PrintTextPostionYTextBox;
        private Label Label33;
        private TextBox PrintTextPostionXTextBox;
        private Label Label32;
        private TextBox PrintTextHeightTextBox;
        private Label Label31;
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