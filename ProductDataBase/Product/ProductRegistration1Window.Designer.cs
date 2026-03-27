namespace ProductDatabase {
    partial class ProductRegistration1Window {
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
            this.RegisterButton = new Button();
            this.TemplateButton = new Button();
            this.CommentComboBox = new ComboBox();
            this.CommentTextBox = new TextBox();
            this.CommentCheckBox = new CheckBox();
            this.PersonComboBox = new ComboBox();
            this.PersonCheckBox = new CheckBox();
            this.RegistrationDateCheckBox = new CheckBox();
            this.FirstSerialNumberTextBox = new TextBox();
            this.FirstSerialNumberCheckBox = new CheckBox();
            this.ExtraTextBox3 = new TextBox();
            this.ExtraCheckBox3 = new CheckBox();
            this.OLesNumberTextBox = new TextBox();
            this.OLesNumberCheckBox = new CheckBox();
            this.RevisionTextBox = new TextBox();
            this.RevisionCheckBox = new CheckBox();
            this.ExtraTextBox1 = new TextBox();
            this.ExtraCheckBox1 = new CheckBox();
            this.QuantityTextBox = new TextBox();
            this.QuantityCheckBox = new CheckBox();
            this.ManufacturingNumberMaskedTextBox = new MaskedTextBox();
            this.ManufacturingNumberCheckBox = new CheckBox();
            this.OrderNumberTextBox = new TextBox();
            this.OrderNumberCheckBox = new CheckBox();
            this.SubstrateModelLabel2 = new Label();
            this.SubstrateModelLabel = new Label();
            this.ファイルToolStripMenuItem = new ToolStripMenuItem();
            this.終了ToolStripMenuItem = new ToolStripMenuItem();
            this.ProductNameLabel2 = new Label();
            this.ヘルプToolStripMenuItem = new ToolStripMenuItem();
            this.メッセージ設定ToolStripMenuItem = new ToolStripMenuItem();
            this.取得情報ToolStripMenuItem = new ToolStripMenuItem();
            this.ProductNameLabel = new Label();
            this.ProductRegistration1MenuStrip = new MenuStrip();
            this.ProductTypeLabel2 = new Label();
            this.ProductTypeLabel = new Label();
            this.panelCommentTextBox = new Panel();
            this.QrCodeTextBox = new TextBox();
            this.QrCodeButton = new Button();
            this.textToUpperCheckBox = new CheckBox();
            this.RevisionChangeButton = new Button();
            this.RegistrationDateTimePicker = new DateTimePicker();
            this.RNumberCheckBox = new CheckBox();
            this.MessageTextBox = new RichTextBox();
            this.ErrorMessageLabel = new Label();
            this.groupBox1 = new GroupBox();
            this.ProductRegistration1MenuStrip.SuspendLayout();
            this.panelCommentTextBox.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // RegisterButton
            // 
            this.RegisterButton.Location = new Point(278, 407);
            this.RegisterButton.Margin = new Padding(0);
            this.RegisterButton.Name = "RegisterButton";
            this.RegisterButton.Size = new Size(75, 25);
            this.RegisterButton.TabIndex = 25;
            this.RegisterButton.Text = "OK";
            this.RegisterButton.UseVisualStyleBackColor = true;
            this.RegisterButton.Click += this.RegisterButton_Click;
            // 
            // TemplateButton
            // 
            this.TemplateButton.Enabled = false;
            this.TemplateButton.Location = new Point(543, 340);
            this.TemplateButton.Name = "TemplateButton";
            this.TemplateButton.Size = new Size(75, 25);
            this.TemplateButton.TabIndex = 24;
            this.TemplateButton.Text = "定型文";
            this.TemplateButton.UseVisualStyleBackColor = true;
            this.TemplateButton.Click += this.TemplateButton_Click;
            // 
            // CommentComboBox
            // 
            this.CommentComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.CommentComboBox.DropDownWidth = 150;
            this.CommentComboBox.Enabled = false;
            this.CommentComboBox.FormattingEnabled = true;
            this.CommentComboBox.Location = new Point(313, 341);
            this.CommentComboBox.Name = "CommentComboBox";
            this.CommentComboBox.Size = new Size(224, 23);
            this.CommentComboBox.TabIndex = 23;
            // 
            // CommentTextBox
            // 
            this.CommentTextBox.Dock = DockStyle.Fill;
            this.CommentTextBox.Enabled = false;
            this.CommentTextBox.Location = new Point(0, 0);
            this.CommentTextBox.Margin = new Padding(0);
            this.CommentTextBox.MaxLength = 500;
            this.CommentTextBox.Multiline = true;
            this.CommentTextBox.Name = "CommentTextBox";
            this.CommentTextBox.ScrollBars = ScrollBars.Vertical;
            this.CommentTextBox.Size = new Size(305, 193);
            this.CommentTextBox.TabIndex = 22;
            // 
            // CommentCheckBox
            // 
            this.CommentCheckBox.Location = new Point(313, 118);
            this.CommentCheckBox.Margin = new Padding(0);
            this.CommentCheckBox.Name = "CommentCheckBox";
            this.CommentCheckBox.Size = new Size(59, 19);
            this.CommentCheckBox.TabIndex = 21;
            this.CommentCheckBox.TabStop = false;
            this.CommentCheckBox.Text = "コメント";
            this.CommentCheckBox.UseVisualStyleBackColor = true;
            this.CommentCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // PersonComboBox
            // 
            this.PersonComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.PersonComboBox.Enabled = false;
            this.PersonComboBox.FormattingEnabled = true;
            this.PersonComboBox.Location = new Point(166, 341);
            this.PersonComboBox.Name = "PersonComboBox";
            this.PersonComboBox.Size = new Size(121, 23);
            this.PersonComboBox.TabIndex = 20;
            // 
            // PersonCheckBox
            // 
            this.PersonCheckBox.Location = new Point(166, 318);
            this.PersonCheckBox.Margin = new Padding(0);
            this.PersonCheckBox.Name = "PersonCheckBox";
            this.PersonCheckBox.Size = new Size(62, 19);
            this.PersonCheckBox.TabIndex = 19;
            this.PersonCheckBox.TabStop = false;
            this.PersonCheckBox.Text = "担当者";
            this.PersonCheckBox.UseVisualStyleBackColor = true;
            this.PersonCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // RegistrationDateCheckBox
            // 
            this.RegistrationDateCheckBox.Location = new Point(166, 268);
            this.RegistrationDateCheckBox.Margin = new Padding(0);
            this.RegistrationDateCheckBox.Name = "RegistrationDateCheckBox";
            this.RegistrationDateCheckBox.Size = new Size(62, 19);
            this.RegistrationDateCheckBox.TabIndex = 17;
            this.RegistrationDateCheckBox.TabStop = false;
            this.RegistrationDateCheckBox.Text = "登録日";
            this.RegistrationDateCheckBox.UseVisualStyleBackColor = true;
            this.RegistrationDateCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // FirstSerialNumberTextBox
            // 
            this.FirstSerialNumberTextBox.Enabled = false;
            this.FirstSerialNumberTextBox.Location = new Point(19, 291);
            this.FirstSerialNumberTextBox.Margin = new Padding(0);
            this.FirstSerialNumberTextBox.MaxLength = 20;
            this.FirstSerialNumberTextBox.Name = "FirstSerialNumberTextBox";
            this.FirstSerialNumberTextBox.Size = new Size(120, 23);
            this.FirstSerialNumberTextBox.TabIndex = 8;
            this.FirstSerialNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // FirstSerialNumberCheckBox
            // 
            this.FirstSerialNumberCheckBox.Location = new Point(19, 268);
            this.FirstSerialNumberCheckBox.Margin = new Padding(0);
            this.FirstSerialNumberCheckBox.Name = "FirstSerialNumberCheckBox";
            this.FirstSerialNumberCheckBox.Size = new Size(110, 19);
            this.FirstSerialNumberCheckBox.TabIndex = 7;
            this.FirstSerialNumberCheckBox.TabStop = false;
            this.FirstSerialNumberCheckBox.Text = "先頭シリアル番号";
            this.FirstSerialNumberCheckBox.UseVisualStyleBackColor = true;
            this.FirstSerialNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox3
            // 
            this.ExtraTextBox3.Enabled = false;
            this.ExtraTextBox3.Location = new Point(166, 241);
            this.ExtraTextBox3.Margin = new Padding(0);
            this.ExtraTextBox3.MaxLength = 20;
            this.ExtraTextBox3.Name = "ExtraTextBox3";
            this.ExtraTextBox3.Size = new Size(120, 23);
            this.ExtraTextBox3.TabIndex = 16;
            this.ExtraTextBox3.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox3
            // 
            this.ExtraCheckBox3.Location = new Point(166, 218);
            this.ExtraCheckBox3.Margin = new Padding(0);
            this.ExtraCheckBox3.Name = "ExtraCheckBox3";
            this.ExtraCheckBox3.Size = new Size(50, 19);
            this.ExtraCheckBox3.TabIndex = 15;
            this.ExtraCheckBox3.TabStop = false;
            this.ExtraCheckBox3.Text = "予備";
            this.ExtraCheckBox3.UseVisualStyleBackColor = true;
            this.ExtraCheckBox3.CheckedChanged += this.CheckBoxChecked;
            // 
            // OLesNumberTextBox
            // 
            this.OLesNumberTextBox.Enabled = false;
            this.OLesNumberTextBox.Location = new Point(166, 191);
            this.OLesNumberTextBox.Margin = new Padding(0);
            this.OLesNumberTextBox.MaxLength = 20;
            this.OLesNumberTextBox.Name = "OLesNumberTextBox";
            this.OLesNumberTextBox.Size = new Size(120, 23);
            this.OLesNumberTextBox.TabIndex = 14;
            this.OLesNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // OLesNumberCheckBox
            // 
            this.OLesNumberCheckBox.Location = new Point(166, 168);
            this.OLesNumberCheckBox.Margin = new Padding(0);
            this.OLesNumberCheckBox.Name = "OLesNumberCheckBox";
            this.OLesNumberCheckBox.Size = new Size(79, 19);
            this.OLesNumberCheckBox.TabIndex = 13;
            this.OLesNumberCheckBox.TabStop = false;
            this.OLesNumberCheckBox.Text = "OLes番号";
            this.OLesNumberCheckBox.UseVisualStyleBackColor = true;
            this.OLesNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // RevisionTextBox
            // 
            this.RevisionTextBox.Enabled = false;
            this.RevisionTextBox.Location = new Point(19, 341);
            this.RevisionTextBox.Margin = new Padding(0);
            this.RevisionTextBox.MaxLength = 4;
            this.RevisionTextBox.Name = "RevisionTextBox";
            this.RevisionTextBox.Size = new Size(120, 23);
            this.RevisionTextBox.TabIndex = 10;
            this.RevisionTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // RevisionCheckBox
            // 
            this.RevisionCheckBox.Location = new Point(19, 318);
            this.RevisionCheckBox.Margin = new Padding(0);
            this.RevisionCheckBox.Name = "RevisionCheckBox";
            this.RevisionCheckBox.Size = new Size(68, 19);
            this.RevisionCheckBox.TabIndex = 9;
            this.RevisionCheckBox.TabStop = false;
            this.RevisionCheckBox.Text = "リビジョン";
            this.RevisionCheckBox.UseVisualStyleBackColor = true;
            this.RevisionCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox1
            // 
            this.ExtraTextBox1.Enabled = false;
            this.ExtraTextBox1.Location = new Point(166, 141);
            this.ExtraTextBox1.Margin = new Padding(0);
            this.ExtraTextBox1.MaxLength = 4;
            this.ExtraTextBox1.Name = "ExtraTextBox1";
            this.ExtraTextBox1.Size = new Size(120, 23);
            this.ExtraTextBox1.TabIndex = 12;
            this.ExtraTextBox1.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox1
            // 
            this.ExtraCheckBox1.Location = new Point(166, 118);
            this.ExtraCheckBox1.Margin = new Padding(0);
            this.ExtraCheckBox1.Name = "ExtraCheckBox1";
            this.ExtraCheckBox1.Size = new Size(62, 19);
            this.ExtraCheckBox1.TabIndex = 11;
            this.ExtraCheckBox1.TabStop = false;
            this.ExtraCheckBox1.Text = "予備";
            this.ExtraCheckBox1.UseVisualStyleBackColor = true;
            this.ExtraCheckBox1.CheckedChanged += this.CheckBoxChecked;
            // 
            // QuantityTextBox
            // 
            this.QuantityTextBox.Enabled = false;
            this.QuantityTextBox.Location = new Point(19, 241);
            this.QuantityTextBox.Margin = new Padding(0);
            this.QuantityTextBox.MaxLength = 4;
            this.QuantityTextBox.Name = "QuantityTextBox";
            this.QuantityTextBox.Size = new Size(120, 23);
            this.QuantityTextBox.TabIndex = 6;
            this.QuantityTextBox.TextAlign = HorizontalAlignment.Right;
            this.QuantityTextBox.KeyPress += this.QuantityTextBox_KeyPress;
            // 
            // QuantityCheckBox
            // 
            this.QuantityCheckBox.Location = new Point(19, 218);
            this.QuantityCheckBox.Margin = new Padding(0);
            this.QuantityCheckBox.Name = "QuantityCheckBox";
            this.QuantityCheckBox.Size = new Size(50, 19);
            this.QuantityCheckBox.TabIndex = 5;
            this.QuantityCheckBox.TabStop = false;
            this.QuantityCheckBox.Text = "数量";
            this.QuantityCheckBox.UseVisualStyleBackColor = true;
            this.QuantityCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ManufacturingNumberMaskedTextBox
            // 
            this.ManufacturingNumberMaskedTextBox.Enabled = false;
            this.ManufacturingNumberMaskedTextBox.Location = new Point(19, 191);
            this.ManufacturingNumberMaskedTextBox.Mask = ">LA00A00000-0000";
            this.ManufacturingNumberMaskedTextBox.Name = "ManufacturingNumberMaskedTextBox";
            this.ManufacturingNumberMaskedTextBox.PromptChar = '*';
            this.ManufacturingNumberMaskedTextBox.Size = new Size(120, 23);
            this.ManufacturingNumberMaskedTextBox.TabIndex = 4;
            this.ManufacturingNumberMaskedTextBox.Text = "H";
            this.ManufacturingNumberMaskedTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // ManufacturingNumberCheckBox
            // 
            this.ManufacturingNumberCheckBox.Location = new Point(19, 168);
            this.ManufacturingNumberCheckBox.Margin = new Padding(0);
            this.ManufacturingNumberCheckBox.Name = "ManufacturingNumberCheckBox";
            this.ManufacturingNumberCheckBox.Size = new Size(74, 19);
            this.ManufacturingNumberCheckBox.TabIndex = 2;
            this.ManufacturingNumberCheckBox.TabStop = false;
            this.ManufacturingNumberCheckBox.Text = "製造番号";
            this.ManufacturingNumberCheckBox.UseVisualStyleBackColor = true;
            this.ManufacturingNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // OrderNumberTextBox
            // 
            this.OrderNumberTextBox.Enabled = false;
            this.OrderNumberTextBox.Location = new Point(19, 141);
            this.OrderNumberTextBox.Margin = new Padding(0);
            this.OrderNumberTextBox.MaxLength = 20;
            this.OrderNumberTextBox.Name = "OrderNumberTextBox";
            this.OrderNumberTextBox.Size = new Size(120, 23);
            this.OrderNumberTextBox.TabIndex = 1;
            this.OrderNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // OrderNumberCheckBox
            // 
            this.OrderNumberCheckBox.Location = new Point(19, 118);
            this.OrderNumberCheckBox.Margin = new Padding(0);
            this.OrderNumberCheckBox.Name = "OrderNumberCheckBox";
            this.OrderNumberCheckBox.Size = new Size(74, 19);
            this.OrderNumberCheckBox.TabIndex = 0;
            this.OrderNumberCheckBox.TabStop = false;
            this.OrderNumberCheckBox.Text = "注文番号";
            this.OrderNumberCheckBox.UseVisualStyleBackColor = true;
            this.OrderNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // SubstrateModelLabel2
            // 
            this.SubstrateModelLabel2.AutoSize = true;
            this.SubstrateModelLabel2.BorderStyle = BorderStyle.Fixed3D;
            this.SubstrateModelLabel2.Location = new Point(90, 85);
            this.SubstrateModelLabel2.Margin = new Padding(4, 0, 4, 0);
            this.SubstrateModelLabel2.Name = "SubstrateModelLabel2";
            this.SubstrateModelLabel2.Size = new Size(24, 17);
            this.SubstrateModelLabel2.TabIndex = 999;
            this.SubstrateModelLabel2.Text = "---";
            this.SubstrateModelLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // SubstrateModelLabel
            // 
            this.SubstrateModelLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.SubstrateModelLabel.AutoSize = true;
            this.SubstrateModelLabel.BorderStyle = BorderStyle.Fixed3D;
            this.SubstrateModelLabel.Location = new Point(16, 85);
            this.SubstrateModelLabel.Margin = new Padding(4, 0, 4, 0);
            this.SubstrateModelLabel.Name = "SubstrateModelLabel";
            this.SubstrateModelLabel.Size = new Size(66, 17);
            this.SubstrateModelLabel.TabIndex = 999;
            this.SubstrateModelLabel.Text = "製品型式 :";
            this.SubstrateModelLabel.TextAlign = ContentAlignment.MiddleRight;
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
            // 
            // ProductNameLabel2
            // 
            this.ProductNameLabel2.AutoSize = true;
            this.ProductNameLabel2.BorderStyle = BorderStyle.Fixed3D;
            this.ProductNameLabel2.Location = new Point(90, 39);
            this.ProductNameLabel2.Margin = new Padding(4, 0, 4, 0);
            this.ProductNameLabel2.Name = "ProductNameLabel2";
            this.ProductNameLabel2.Size = new Size(24, 17);
            this.ProductNameLabel2.TabIndex = 999;
            this.ProductNameLabel2.Text = "---";
            this.ProductNameLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // ヘルプToolStripMenuItem
            // 
            this.ヘルプToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { this.メッセージ設定ToolStripMenuItem, this.取得情報ToolStripMenuItem });
            this.ヘルプToolStripMenuItem.Name = "ヘルプToolStripMenuItem";
            this.ヘルプToolStripMenuItem.Size = new Size(48, 20);
            this.ヘルプToolStripMenuItem.Text = "ヘルプ";
            // 
            // メッセージ設定ToolStripMenuItem
            // 
            this.メッセージ設定ToolStripMenuItem.Name = "メッセージ設定ToolStripMenuItem";
            this.メッセージ設定ToolStripMenuItem.Size = new Size(143, 22);
            this.メッセージ設定ToolStripMenuItem.Text = "メッセージ設定";
            this.メッセージ設定ToolStripMenuItem.Click += this.メッセージ設定ToolStripMenuItem_Click;
            // 
            // 取得情報ToolStripMenuItem
            // 
            this.取得情報ToolStripMenuItem.Name = "取得情報ToolStripMenuItem";
            this.取得情報ToolStripMenuItem.Size = new Size(143, 22);
            this.取得情報ToolStripMenuItem.Text = "取得情報";
            this.取得情報ToolStripMenuItem.Click += this.取得情報ToolStripMenuItem_Click;
            // 
            // ProductNameLabel
            // 
            this.ProductNameLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.ProductNameLabel.AutoSize = true;
            this.ProductNameLabel.BorderStyle = BorderStyle.Fixed3D;
            this.ProductNameLabel.Location = new Point(28, 39);
            this.ProductNameLabel.Margin = new Padding(4, 0, 4, 0);
            this.ProductNameLabel.Name = "ProductNameLabel";
            this.ProductNameLabel.Size = new Size(54, 17);
            this.ProductNameLabel.TabIndex = 999;
            this.ProductNameLabel.Text = "製品名 :";
            this.ProductNameLabel.TextAlign = ContentAlignment.MiddleRight;
            // 
            // ProductRegistration1MenuStrip
            // 
            this.ProductRegistration1MenuStrip.Font = new Font("Meiryo UI", 9F);
            this.ProductRegistration1MenuStrip.Items.AddRange(new ToolStripItem[] { this.ファイルToolStripMenuItem, this.ヘルプToolStripMenuItem });
            this.ProductRegistration1MenuStrip.Location = new Point(0, 0);
            this.ProductRegistration1MenuStrip.Name = "ProductRegistration1MenuStrip";
            this.ProductRegistration1MenuStrip.Size = new Size(630, 24);
            this.ProductRegistration1MenuStrip.TabIndex = 999;
            this.ProductRegistration1MenuStrip.Text = "menuStrip";
            // 
            // ProductTypeLabel2
            // 
            this.ProductTypeLabel2.AutoSize = true;
            this.ProductTypeLabel2.BorderStyle = BorderStyle.Fixed3D;
            this.ProductTypeLabel2.Location = new Point(90, 62);
            this.ProductTypeLabel2.Margin = new Padding(4, 0, 4, 0);
            this.ProductTypeLabel2.Name = "ProductTypeLabel2";
            this.ProductTypeLabel2.Size = new Size(24, 17);
            this.ProductTypeLabel2.TabIndex = 999;
            this.ProductTypeLabel2.Text = "---";
            this.ProductTypeLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // ProductTypeLabel
            // 
            this.ProductTypeLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.ProductTypeLabel.AutoSize = true;
            this.ProductTypeLabel.BorderStyle = BorderStyle.Fixed3D;
            this.ProductTypeLabel.Location = new Point(38, 62);
            this.ProductTypeLabel.Margin = new Padding(4, 0, 4, 0);
            this.ProductTypeLabel.Name = "ProductTypeLabel";
            this.ProductTypeLabel.Size = new Size(44, 17);
            this.ProductTypeLabel.TabIndex = 999;
            this.ProductTypeLabel.Text = "タイプ :";
            this.ProductTypeLabel.TextAlign = ContentAlignment.MiddleRight;
            // 
            // panelCommentTextBox
            // 
            this.panelCommentTextBox.Controls.Add(this.CommentTextBox);
            this.panelCommentTextBox.Location = new Point(313, 141);
            this.panelCommentTextBox.Name = "panelCommentTextBox";
            this.panelCommentTextBox.Size = new Size(305, 193);
            this.panelCommentTextBox.TabIndex = 22;
            // 
            // QrCodeTextBox
            // 
            this.QrCodeTextBox.Location = new Point(9, 22);
            this.QrCodeTextBox.MaxLength = 100;
            this.QrCodeTextBox.Name = "QrCodeTextBox";
            this.QrCodeTextBox.Size = new Size(305, 23);
            this.QrCodeTextBox.TabIndex = 999;
            this.QrCodeTextBox.Enter += this.QrCodeTextBox_Enter;
            // 
            // QrCodeButton
            // 
            this.QrCodeButton.Location = new Point(239, 51);
            this.QrCodeButton.Name = "QrCodeButton";
            this.QrCodeButton.Size = new Size(75, 23);
            this.QrCodeButton.TabIndex = 999;
            this.QrCodeButton.Text = "入力";
            this.QrCodeButton.UseVisualStyleBackColor = true;
            this.QrCodeButton.Click += this.QrCodeButton_Click;
            // 
            // textToUpperCheckBox
            // 
            this.textToUpperCheckBox.AutoSize = true;
            this.textToUpperCheckBox.Checked = true;
            this.textToUpperCheckBox.CheckState = CheckState.Checked;
            this.textToUpperCheckBox.Location = new Point(93, 54);
            this.textToUpperCheckBox.Name = "textToUpperCheckBox";
            this.textToUpperCheckBox.Size = new Size(140, 19);
            this.textToUpperCheckBox.TabIndex = 999;
            this.textToUpperCheckBox.Text = "小文字を大文字に変換";
            this.textToUpperCheckBox.UseVisualStyleBackColor = true;
            // 
            // RevisionChangeButton
            // 
            this.RevisionChangeButton.Location = new Point(19, 375);
            this.RevisionChangeButton.Name = "RevisionChangeButton";
            this.RevisionChangeButton.Size = new Size(75, 25);
            this.RevisionChangeButton.TabIndex = 999;
            this.RevisionChangeButton.Text = "Rev変更";
            this.RevisionChangeButton.UseVisualStyleBackColor = true;
            this.RevisionChangeButton.Click += this.RevisionChangeButton_Click;
            // 
            // RegistrationDateTimePicker
            // 
            this.RegistrationDateTimePicker.Format = DateTimePickerFormat.Short;
            this.RegistrationDateTimePicker.Location = new Point(166, 291);
            this.RegistrationDateTimePicker.Name = "RegistrationDateTimePicker";
            this.RegistrationDateTimePicker.Size = new Size(120, 23);
            this.RegistrationDateTimePicker.TabIndex = 18;
            // 
            // RNumberCheckBox
            // 
            this.RNumberCheckBox.AutoSize = true;
            this.RNumberCheckBox.Location = new Point(96, 168);
            this.RNumberCheckBox.Name = "RNumberCheckBox";
            this.RNumberCheckBox.Size = new Size(46, 19);
            this.RNumberCheckBox.TabIndex = 3;
            this.RNumberCheckBox.Text = "R番";
            this.RNumberCheckBox.UseVisualStyleBackColor = true;
            this.RNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // MessageTextBox
            // 
            this.MessageTextBox.Enabled = false;
            this.MessageTextBox.ForeColor = Color.Red;
            this.MessageTextBox.Location = new Point(168, 377);
            this.MessageTextBox.MaxLength = 100;
            this.MessageTextBox.Multiline = false;
            this.MessageTextBox.Name = "MessageTextBox";
            this.MessageTextBox.ReadOnly = true;
            this.MessageTextBox.ScrollBars = RichTextBoxScrollBars.None;
            this.MessageTextBox.Size = new Size(450, 25);
            this.MessageTextBox.TabIndex = 999;
            this.MessageTextBox.TabStop = false;
            this.MessageTextBox.Text = "";
            this.MessageTextBox.WordWrap = false;
            // 
            // ErrorMessageLabel
            // 
            this.ErrorMessageLabel.Location = new Point(115, 432);
            this.ErrorMessageLabel.Name = "ErrorMessageLabel";
            this.ErrorMessageLabel.Size = new Size(400, 15);
            this.ErrorMessageLabel.TabIndex = 1001;
            this.ErrorMessageLabel.Text = "----------";
            this.ErrorMessageLabel.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.QrCodeTextBox);
            this.groupBox1.Controls.Add(this.QrCodeButton);
            this.groupBox1.Controls.Add(this.textToUpperCheckBox);
            this.groupBox1.Location = new Point(298, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new Size(320, 80);
            this.groupBox1.TabIndex = 1003;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "QRコード入力用";
            // 
            // ProductRegistration1Window
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(630, 455);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.ErrorMessageLabel);
            this.Controls.Add(this.MessageTextBox);
            this.Controls.Add(this.RNumberCheckBox);
            this.Controls.Add(this.RegistrationDateTimePicker);
            this.Controls.Add(this.RevisionChangeButton);
            this.Controls.Add(this.ProductTypeLabel2);
            this.Controls.Add(this.ProductTypeLabel);
            this.Controls.Add(this.RegisterButton);
            this.Controls.Add(this.TemplateButton);
            this.Controls.Add(this.CommentComboBox);
            this.Controls.Add(this.CommentCheckBox);
            this.Controls.Add(this.PersonComboBox);
            this.Controls.Add(this.PersonCheckBox);
            this.Controls.Add(this.RegistrationDateCheckBox);
            this.Controls.Add(this.FirstSerialNumberTextBox);
            this.Controls.Add(this.FirstSerialNumberCheckBox);
            this.Controls.Add(this.ExtraTextBox3);
            this.Controls.Add(this.ExtraCheckBox3);
            this.Controls.Add(this.OLesNumberTextBox);
            this.Controls.Add(this.OLesNumberCheckBox);
            this.Controls.Add(this.RevisionTextBox);
            this.Controls.Add(this.RevisionCheckBox);
            this.Controls.Add(this.ExtraTextBox1);
            this.Controls.Add(this.ExtraCheckBox1);
            this.Controls.Add(this.QuantityTextBox);
            this.Controls.Add(this.QuantityCheckBox);
            this.Controls.Add(this.ManufacturingNumberMaskedTextBox);
            this.Controls.Add(this.ManufacturingNumberCheckBox);
            this.Controls.Add(this.OrderNumberTextBox);
            this.Controls.Add(this.OrderNumberCheckBox);
            this.Controls.Add(this.SubstrateModelLabel2);
            this.Controls.Add(this.SubstrateModelLabel);
            this.Controls.Add(this.ProductNameLabel2);
            this.Controls.Add(this.ProductNameLabel);
            this.Controls.Add(this.ProductRegistration1MenuStrip);
            this.Controls.Add(this.panelCommentTextBox);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MainMenuStrip = this.ProductRegistration1MenuStrip;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProductRegistration1Window";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "製品登録";
            this.Load += this.ProductRegistration1Window_Load;
            this.ProductRegistration1MenuStrip.ResumeLayout(false);
            this.ProductRegistration1MenuStrip.PerformLayout();
            this.panelCommentTextBox.ResumeLayout(false);
            this.panelCommentTextBox.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private Button RegisterButton;
        private Button TemplateButton;
        private ComboBox CommentComboBox;
        private TextBox CommentTextBox;
        private CheckBox CommentCheckBox;
        private ComboBox PersonComboBox;
        private CheckBox PersonCheckBox;
        private CheckBox RegistrationDateCheckBox;
        private TextBox FirstSerialNumberTextBox;
        private CheckBox FirstSerialNumberCheckBox;
        private TextBox ExtraTextBox3;
        private CheckBox ExtraCheckBox3;
        private TextBox OLesNumberTextBox;
        private CheckBox OLesNumberCheckBox;
        private TextBox RevisionTextBox;
        private CheckBox RevisionCheckBox;
        private TextBox ExtraTextBox1;
        private CheckBox ExtraCheckBox1;
        private TextBox QuantityTextBox;
        private CheckBox QuantityCheckBox;
        private MaskedTextBox ManufacturingNumberMaskedTextBox;
        private CheckBox ManufacturingNumberCheckBox;
        private TextBox OrderNumberTextBox;
        private CheckBox OrderNumberCheckBox;
        private Label SubstrateModelLabel2;
        private Label SubstrateModelLabel;
        private ToolStripMenuItem ファイルToolStripMenuItem;
        private ToolStripMenuItem 終了ToolStripMenuItem;
        private Label ProductNameLabel2;
        private ToolStripMenuItem ヘルプToolStripMenuItem;
        private ToolStripMenuItem 取得情報ToolStripMenuItem;
        private Label ProductNameLabel;
        private MenuStrip ProductRegistration1MenuStrip;
        private Label ProductTypeLabel2;
        private Label ProductTypeLabel;
        private Panel panelCommentTextBox;
        private TextBox QrCodeTextBox;
        private Button QrCodeButton;
        private CheckBox textToUpperCheckBox;
        private Button RevisionChangeButton;
        private DateTimePicker RegistrationDateTimePicker;
        private CheckBox RNumberCheckBox;
        private ToolStripMenuItem メッセージ設定ToolStripMenuItem;
        private RichTextBox MessageTextBox;
        private Label ErrorMessageLabel;
        private GroupBox groupBox1;
    }
}