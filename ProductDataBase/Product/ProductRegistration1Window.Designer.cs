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
            this.ExtraTextBox6 = new TextBox();
            this.ExtraCheckBox6 = new CheckBox();
            this.ExtraTextBox5 = new TextBox();
            this.ExtraCheckBox5 = new CheckBox();
            this.ExtraTextBox4 = new TextBox();
            this.ExtraCheckBox4 = new CheckBox();
            this.PersonComboBox = new ComboBox();
            this.PersonCheckBox = new CheckBox();
            this.RegistrationDateMaskedTextBox = new MaskedTextBox();
            this.RegistrationDateCheckBox = new CheckBox();
            this.FirstSerialNumberTextBox = new TextBox();
            this.FirstSerialNumberCheckBox = new CheckBox();
            this.ExtraTextBox3 = new TextBox();
            this.ExtraCheckBox3 = new CheckBox();
            this.ExtraTextBox2 = new TextBox();
            this.ExtraCheckBox2 = new CheckBox();
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
            this.SubstrateModelLabel1 = new Label();
            this.ファイルToolStripMenuItem = new ToolStripMenuItem();
            this.終了ToolStripMenuItem = new ToolStripMenuItem();
            this.ProductNameLabel2 = new Label();
            this.ヘルプToolStripMenuItem = new ToolStripMenuItem();
            this.取得情報ToolStripMenuItem = new ToolStripMenuItem();
            this.ProductNameLabel1 = new Label();
            this.ProductRegistration1MenuStrip = new MenuStrip();
            this.ProductTypeLabel2 = new Label();
            this.ProductTypeLabel1 = new Label();
            this.ProductRegistration1MenuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // RegisterButton
            // 
            this.RegisterButton.Location = new Point(313, 411);
            this.RegisterButton.Margin = new Padding(0);
            this.RegisterButton.Name = "RegisterButton";
            this.RegisterButton.Size = new Size(75, 25);
            this.RegisterButton.TabIndex = 81;
            this.RegisterButton.Text = "OK";
            this.RegisterButton.UseVisualStyleBackColor = true;
            this.RegisterButton.Click += this.RegisterButton_Click;
            // 
            // TemplateButton
            // 
            this.TemplateButton.Enabled = false;
            this.TemplateButton.Location = new Point(539, 347);
            this.TemplateButton.Name = "TemplateButton";
            this.TemplateButton.Size = new Size(75, 25);
            this.TemplateButton.TabIndex = 78;
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
            this.CommentComboBox.Location = new Point(464, 320);
            this.CommentComboBox.Name = "CommentComboBox";
            this.CommentComboBox.Size = new Size(150, 23);
            this.CommentComboBox.TabIndex = 77;
            // 
            // CommentTextBox
            // 
            this.CommentTextBox.Enabled = false;
            this.CommentTextBox.Location = new Point(464, 141);
            this.CommentTextBox.Margin = new Padding(0);
            this.CommentTextBox.MaxLength = 500;
            this.CommentTextBox.Multiline = true;
            this.CommentTextBox.Name = "CommentTextBox";
            this.CommentTextBox.ScrollBars = ScrollBars.Vertical;
            this.CommentTextBox.Size = new Size(150, 175);
            this.CommentTextBox.TabIndex = 76;
            // 
            // CommentCheckBox
            // 
            this.CommentCheckBox.Location = new Point(464, 118);
            this.CommentCheckBox.Margin = new Padding(0);
            this.CommentCheckBox.Name = "CommentCheckBox";
            this.CommentCheckBox.Size = new Size(59, 19);
            this.CommentCheckBox.TabIndex = 75;
            this.CommentCheckBox.TabStop = false;
            this.CommentCheckBox.Text = "コメント";
            this.CommentCheckBox.UseVisualStyleBackColor = true;
            this.CommentCheckBox.CheckedChanged += this.NumberCheckBox_CheckedChanged;
            // 
            // ExtraTextBox6
            // 
            this.ExtraTextBox6.Enabled = false;
            this.ExtraTextBox6.Location = new Point(313, 241);
            this.ExtraTextBox6.Margin = new Padding(0);
            this.ExtraTextBox6.MaxLength = 20;
            this.ExtraTextBox6.Name = "ExtraTextBox6";
            this.ExtraTextBox6.Size = new Size(120, 23);
            this.ExtraTextBox6.TabIndex = 74;
            this.ExtraTextBox6.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox6
            // 
            this.ExtraCheckBox6.Location = new Point(313, 218);
            this.ExtraCheckBox6.Margin = new Padding(0);
            this.ExtraCheckBox6.Name = "ExtraCheckBox6";
            this.ExtraCheckBox6.Size = new Size(50, 19);
            this.ExtraCheckBox6.TabIndex = 73;
            this.ExtraCheckBox6.TabStop = false;
            this.ExtraCheckBox6.Text = "予備";
            this.ExtraCheckBox6.UseVisualStyleBackColor = true;
            this.ExtraCheckBox6.Visible = false;
            this.ExtraCheckBox6.CheckedChanged += this.NumberCheckBox_CheckedChanged;
            // 
            // ExtraTextBox5
            // 
            this.ExtraTextBox5.Enabled = false;
            this.ExtraTextBox5.Location = new Point(313, 191);
            this.ExtraTextBox5.Margin = new Padding(0);
            this.ExtraTextBox5.MaxLength = 20;
            this.ExtraTextBox5.Name = "ExtraTextBox5";
            this.ExtraTextBox5.Size = new Size(120, 23);
            this.ExtraTextBox5.TabIndex = 72;
            this.ExtraTextBox5.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox5
            // 
            this.ExtraCheckBox5.Location = new Point(313, 168);
            this.ExtraCheckBox5.Margin = new Padding(0);
            this.ExtraCheckBox5.Name = "ExtraCheckBox5";
            this.ExtraCheckBox5.Size = new Size(50, 19);
            this.ExtraCheckBox5.TabIndex = 71;
            this.ExtraCheckBox5.TabStop = false;
            this.ExtraCheckBox5.Text = "予備";
            this.ExtraCheckBox5.UseVisualStyleBackColor = true;
            this.ExtraCheckBox5.Visible = false;
            this.ExtraCheckBox5.CheckedChanged += this.NumberCheckBox_CheckedChanged;
            // 
            // ExtraTextBox4
            // 
            this.ExtraTextBox4.Enabled = false;
            this.ExtraTextBox4.Location = new Point(313, 141);
            this.ExtraTextBox4.Margin = new Padding(0);
            this.ExtraTextBox4.MaxLength = 20;
            this.ExtraTextBox4.Name = "ExtraTextBox4";
            this.ExtraTextBox4.Size = new Size(120, 23);
            this.ExtraTextBox4.TabIndex = 70;
            this.ExtraTextBox4.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox4
            // 
            this.ExtraCheckBox4.Location = new Point(313, 118);
            this.ExtraCheckBox4.Margin = new Padding(0);
            this.ExtraCheckBox4.Name = "ExtraCheckBox4";
            this.ExtraCheckBox4.Size = new Size(50, 19);
            this.ExtraCheckBox4.TabIndex = 69;
            this.ExtraCheckBox4.TabStop = false;
            this.ExtraCheckBox4.Text = "予備";
            this.ExtraCheckBox4.UseVisualStyleBackColor = true;
            this.ExtraCheckBox4.Visible = false;
            this.ExtraCheckBox4.CheckedChanged += this.NumberCheckBox_CheckedChanged;
            // 
            // PersonComboBox
            // 
            this.PersonComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.PersonComboBox.Enabled = false;
            this.PersonComboBox.FormattingEnabled = true;
            this.PersonComboBox.Location = new Point(166, 341);
            this.PersonComboBox.Name = "PersonComboBox";
            this.PersonComboBox.Size = new Size(121, 23);
            this.PersonComboBox.TabIndex = 68;
            // 
            // PersonCheckBox
            // 
            this.PersonCheckBox.Location = new Point(166, 318);
            this.PersonCheckBox.Margin = new Padding(0);
            this.PersonCheckBox.Name = "PersonCheckBox";
            this.PersonCheckBox.Size = new Size(62, 19);
            this.PersonCheckBox.TabIndex = 67;
            this.PersonCheckBox.TabStop = false;
            this.PersonCheckBox.Text = "担当者";
            this.PersonCheckBox.UseVisualStyleBackColor = true;
            this.PersonCheckBox.CheckedChanged += this.NumberCheckBox_CheckedChanged;
            // 
            // RegistrationDateMaskedTextBox
            // 
            this.RegistrationDateMaskedTextBox.Enabled = false;
            this.RegistrationDateMaskedTextBox.Location = new Point(166, 291);
            this.RegistrationDateMaskedTextBox.Mask = "0000/00/00";
            this.RegistrationDateMaskedTextBox.Name = "RegistrationDateMaskedTextBox";
            this.RegistrationDateMaskedTextBox.PromptChar = '*';
            this.RegistrationDateMaskedTextBox.Size = new Size(120, 23);
            this.RegistrationDateMaskedTextBox.TabIndex = 66;
            this.RegistrationDateMaskedTextBox.TextAlign = HorizontalAlignment.Right;
            this.RegistrationDateMaskedTextBox.TypeValidationCompleted += this.RegistrationDateMaskedTextBox_TypeValidationCompleted;
            // 
            // RegistrationDateCheckBox
            // 
            this.RegistrationDateCheckBox.Location = new Point(166, 268);
            this.RegistrationDateCheckBox.Margin = new Padding(0);
            this.RegistrationDateCheckBox.Name = "RegistrationDateCheckBox";
            this.RegistrationDateCheckBox.Size = new Size(62, 19);
            this.RegistrationDateCheckBox.TabIndex = 65;
            this.RegistrationDateCheckBox.TabStop = false;
            this.RegistrationDateCheckBox.Text = "登録日";
            this.RegistrationDateCheckBox.UseVisualStyleBackColor = true;
            this.RegistrationDateCheckBox.CheckedChanged += this.NumberCheckBox_CheckedChanged;
            // 
            // FirstSerialNumberTextBox
            // 
            this.FirstSerialNumberTextBox.Enabled = false;
            this.FirstSerialNumberTextBox.Location = new Point(166, 241);
            this.FirstSerialNumberTextBox.Margin = new Padding(0);
            this.FirstSerialNumberTextBox.MaxLength = 20;
            this.FirstSerialNumberTextBox.Name = "FirstSerialNumberTextBox";
            this.FirstSerialNumberTextBox.Size = new Size(120, 23);
            this.FirstSerialNumberTextBox.TabIndex = 64;
            this.FirstSerialNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // FirstSerialNumberCheckBox
            // 
            this.FirstSerialNumberCheckBox.Location = new Point(166, 218);
            this.FirstSerialNumberCheckBox.Margin = new Padding(0);
            this.FirstSerialNumberCheckBox.Name = "FirstSerialNumberCheckBox";
            this.FirstSerialNumberCheckBox.Size = new Size(110, 19);
            this.FirstSerialNumberCheckBox.TabIndex = 63;
            this.FirstSerialNumberCheckBox.TabStop = false;
            this.FirstSerialNumberCheckBox.Text = "先頭シリアル番号";
            this.FirstSerialNumberCheckBox.UseVisualStyleBackColor = true;
            this.FirstSerialNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox3
            // 
            this.ExtraTextBox3.Enabled = false;
            this.ExtraTextBox3.Location = new Point(166, 191);
            this.ExtraTextBox3.Margin = new Padding(0);
            this.ExtraTextBox3.MaxLength = 20;
            this.ExtraTextBox3.Name = "ExtraTextBox3";
            this.ExtraTextBox3.Size = new Size(120, 23);
            this.ExtraTextBox3.TabIndex = 62;
            this.ExtraTextBox3.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox3
            // 
            this.ExtraCheckBox3.Location = new Point(166, 168);
            this.ExtraCheckBox3.Margin = new Padding(0);
            this.ExtraCheckBox3.Name = "ExtraCheckBox3";
            this.ExtraCheckBox3.Size = new Size(50, 19);
            this.ExtraCheckBox3.TabIndex = 61;
            this.ExtraCheckBox3.TabStop = false;
            this.ExtraCheckBox3.Text = "予備";
            this.ExtraCheckBox3.UseVisualStyleBackColor = true;
            this.ExtraCheckBox3.Visible = false;
            this.ExtraCheckBox3.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox2
            // 
            this.ExtraTextBox2.Enabled = false;
            this.ExtraTextBox2.Location = new Point(166, 141);
            this.ExtraTextBox2.Margin = new Padding(0);
            this.ExtraTextBox2.MaxLength = 20;
            this.ExtraTextBox2.Name = "ExtraTextBox2";
            this.ExtraTextBox2.Size = new Size(120, 23);
            this.ExtraTextBox2.TabIndex = 60;
            this.ExtraTextBox2.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox2
            // 
            this.ExtraCheckBox2.Location = new Point(166, 118);
            this.ExtraCheckBox2.Margin = new Padding(0);
            this.ExtraCheckBox2.Name = "ExtraCheckBox2";
            this.ExtraCheckBox2.Size = new Size(50, 19);
            this.ExtraCheckBox2.TabIndex = 59;
            this.ExtraCheckBox2.TabStop = false;
            this.ExtraCheckBox2.Text = "予備";
            this.ExtraCheckBox2.UseVisualStyleBackColor = true;
            this.ExtraCheckBox2.Visible = false;
            this.ExtraCheckBox2.CheckedChanged += this.CheckBoxChecked;
            // 
            // RevisionTextBox
            // 
            this.RevisionTextBox.Enabled = false;
            this.RevisionTextBox.Location = new Point(19, 341);
            this.RevisionTextBox.Margin = new Padding(0);
            this.RevisionTextBox.MaxLength = 4;
            this.RevisionTextBox.Name = "RevisionTextBox";
            this.RevisionTextBox.Size = new Size(120, 23);
            this.RevisionTextBox.TabIndex = 58;
            this.RevisionTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // RevisionCheckBox
            // 
            this.RevisionCheckBox.Location = new Point(19, 318);
            this.RevisionCheckBox.Margin = new Padding(0);
            this.RevisionCheckBox.Name = "RevisionCheckBox";
            this.RevisionCheckBox.Size = new Size(68, 19);
            this.RevisionCheckBox.TabIndex = 57;
            this.RevisionCheckBox.TabStop = false;
            this.RevisionCheckBox.Text = "レビジョン";
            this.RevisionCheckBox.UseVisualStyleBackColor = true;
            this.RevisionCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox1
            // 
            this.ExtraTextBox1.Enabled = false;
            this.ExtraTextBox1.Location = new Point(19, 291);
            this.ExtraTextBox1.Margin = new Padding(0);
            this.ExtraTextBox1.MaxLength = 4;
            this.ExtraTextBox1.Name = "ExtraTextBox1";
            this.ExtraTextBox1.Size = new Size(120, 23);
            this.ExtraTextBox1.TabIndex = 56;
            this.ExtraTextBox1.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox1
            // 
            this.ExtraCheckBox1.Location = new Point(19, 268);
            this.ExtraCheckBox1.Margin = new Padding(0);
            this.ExtraCheckBox1.Name = "ExtraCheckBox1";
            this.ExtraCheckBox1.Size = new Size(62, 19);
            this.ExtraCheckBox1.TabIndex = 55;
            this.ExtraCheckBox1.TabStop = false;
            this.ExtraCheckBox1.Text = "予備";
            this.ExtraCheckBox1.UseVisualStyleBackColor = true;
            this.ExtraCheckBox1.Visible = false;
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
            this.QuantityTextBox.TabIndex = 54;
            this.QuantityTextBox.TextAlign = HorizontalAlignment.Right;
            this.QuantityTextBox.KeyPress += this.QuantityTextBox_KeyPress;
            // 
            // QuantityCheckBox
            // 
            this.QuantityCheckBox.Location = new Point(19, 218);
            this.QuantityCheckBox.Margin = new Padding(0);
            this.QuantityCheckBox.Name = "QuantityCheckBox";
            this.QuantityCheckBox.Size = new Size(50, 19);
            this.QuantityCheckBox.TabIndex = 53;
            this.QuantityCheckBox.TabStop = false;
            this.QuantityCheckBox.Text = "数量";
            this.QuantityCheckBox.UseVisualStyleBackColor = true;
            this.QuantityCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ManufacturingNumberMaskedTextBox
            // 
            this.ManufacturingNumberMaskedTextBox.Enabled = false;
            this.ManufacturingNumberMaskedTextBox.Location = new Point(19, 191);
            this.ManufacturingNumberMaskedTextBox.Mask = "LA00L00000-0000";
            this.ManufacturingNumberMaskedTextBox.Name = "ManufacturingNumberMaskedTextBox";
            this.ManufacturingNumberMaskedTextBox.PromptChar = '*';
            this.ManufacturingNumberMaskedTextBox.Size = new Size(120, 23);
            this.ManufacturingNumberMaskedTextBox.TabIndex = 52;
            this.ManufacturingNumberMaskedTextBox.Text = "H";
            this.ManufacturingNumberMaskedTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // ManufacturingNumberCheckBox
            // 
            this.ManufacturingNumberCheckBox.Location = new Point(19, 168);
            this.ManufacturingNumberCheckBox.Margin = new Padding(0);
            this.ManufacturingNumberCheckBox.Name = "ManufacturingNumberCheckBox";
            this.ManufacturingNumberCheckBox.Size = new Size(74, 19);
            this.ManufacturingNumberCheckBox.TabIndex = 51;
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
            this.OrderNumberTextBox.TabIndex = 50;
            this.OrderNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // OrderNumberCheckBox
            // 
            this.OrderNumberCheckBox.Location = new Point(19, 118);
            this.OrderNumberCheckBox.Margin = new Padding(0);
            this.OrderNumberCheckBox.Name = "OrderNumberCheckBox";
            this.OrderNumberCheckBox.Size = new Size(74, 19);
            this.OrderNumberCheckBox.TabIndex = 49;
            this.OrderNumberCheckBox.TabStop = false;
            this.OrderNumberCheckBox.Text = "注文番号";
            this.OrderNumberCheckBox.UseVisualStyleBackColor = true;
            this.OrderNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // SubstrateModelLabel2
            // 
            this.SubstrateModelLabel2.AutoSize = true;
            this.SubstrateModelLabel2.BorderStyle = BorderStyle.Fixed3D;
            this.SubstrateModelLabel2.Location = new Point(84, 81);
            this.SubstrateModelLabel2.Margin = new Padding(4, 0, 4, 0);
            this.SubstrateModelLabel2.Name = "SubstrateModelLabel2";
            this.SubstrateModelLabel2.Size = new Size(24, 17);
            this.SubstrateModelLabel2.TabIndex = 46;
            this.SubstrateModelLabel2.Text = "---";
            this.SubstrateModelLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // SubstrateModelLabel1
            // 
            this.SubstrateModelLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.SubstrateModelLabel1.Location = new Point(16, 82);
            this.SubstrateModelLabel1.Margin = new Padding(4, 0, 4, 0);
            this.SubstrateModelLabel1.Name = "SubstrateModelLabel1";
            this.SubstrateModelLabel1.Size = new Size(60, 15);
            this.SubstrateModelLabel1.TabIndex = 45;
            this.SubstrateModelLabel1.Text = "基盤型式:";
            this.SubstrateModelLabel1.TextAlign = ContentAlignment.MiddleRight;
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
            this.ProductNameLabel2.Location = new Point(84, 37);
            this.ProductNameLabel2.Margin = new Padding(4, 0, 4, 0);
            this.ProductNameLabel2.Name = "ProductNameLabel2";
            this.ProductNameLabel2.Size = new Size(24, 17);
            this.ProductNameLabel2.TabIndex = 44;
            this.ProductNameLabel2.Text = "---";
            this.ProductNameLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // ヘルプToolStripMenuItem
            // 
            this.ヘルプToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { this.取得情報ToolStripMenuItem });
            this.ヘルプToolStripMenuItem.Name = "ヘルプToolStripMenuItem";
            this.ヘルプToolStripMenuItem.Size = new Size(48, 20);
            this.ヘルプToolStripMenuItem.Text = "ヘルプ";
            // 
            // 取得情報ToolStripMenuItem
            // 
            this.取得情報ToolStripMenuItem.Name = "取得情報ToolStripMenuItem";
            this.取得情報ToolStripMenuItem.Size = new Size(122, 22);
            this.取得情報ToolStripMenuItem.Text = "取得情報";
            this.取得情報ToolStripMenuItem.Click += this.取得情報ToolStripMenuItem_Click;
            // 
            // ProductNameLabel1
            // 
            this.ProductNameLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.ProductNameLabel1.Location = new Point(28, 39);
            this.ProductNameLabel1.Margin = new Padding(4, 0, 4, 0);
            this.ProductNameLabel1.Name = "ProductNameLabel1";
            this.ProductNameLabel1.Size = new Size(48, 15);
            this.ProductNameLabel1.TabIndex = 43;
            this.ProductNameLabel1.Text = "製品名:";
            this.ProductNameLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // ProductRegistration1MenuStrip
            // 
            this.ProductRegistration1MenuStrip.Font = new Font("Meiryo UI", 9F);
            this.ProductRegistration1MenuStrip.Items.AddRange(new ToolStripItem[] { this.ファイルToolStripMenuItem, this.ヘルプToolStripMenuItem });
            this.ProductRegistration1MenuStrip.Location = new Point(0, 0);
            this.ProductRegistration1MenuStrip.Name = "ProductRegistration1MenuStrip";
            this.ProductRegistration1MenuStrip.Size = new Size(630, 24);
            this.ProductRegistration1MenuStrip.TabIndex = 42;
            this.ProductRegistration1MenuStrip.Text = "menuStrip";
            // 
            // ProductTypeLabel2
            // 
            this.ProductTypeLabel2.AutoSize = true;
            this.ProductTypeLabel2.BorderStyle = BorderStyle.Fixed3D;
            this.ProductTypeLabel2.Location = new Point(84, 59);
            this.ProductTypeLabel2.Margin = new Padding(4, 0, 4, 0);
            this.ProductTypeLabel2.Name = "ProductTypeLabel2";
            this.ProductTypeLabel2.Size = new Size(24, 17);
            this.ProductTypeLabel2.TabIndex = 83;
            this.ProductTypeLabel2.Text = "---";
            this.ProductTypeLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // ProductTypeLabel1
            // 
            this.ProductTypeLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.ProductTypeLabel1.Location = new Point(16, 61);
            this.ProductTypeLabel1.Margin = new Padding(4, 0, 4, 0);
            this.ProductTypeLabel1.Name = "ProductTypeLabel1";
            this.ProductTypeLabel1.Size = new Size(60, 15);
            this.ProductTypeLabel1.TabIndex = 82;
            this.ProductTypeLabel1.Text = "タイプ:";
            this.ProductTypeLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // ProductRegistration1Window
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(630, 441);
            this.Controls.Add(this.ProductTypeLabel2);
            this.Controls.Add(this.ProductTypeLabel1);
            this.Controls.Add(this.RegisterButton);
            this.Controls.Add(this.TemplateButton);
            this.Controls.Add(this.CommentComboBox);
            this.Controls.Add(this.CommentTextBox);
            this.Controls.Add(this.CommentCheckBox);
            this.Controls.Add(this.ExtraTextBox6);
            this.Controls.Add(this.ExtraCheckBox6);
            this.Controls.Add(this.ExtraTextBox5);
            this.Controls.Add(this.ExtraCheckBox5);
            this.Controls.Add(this.ExtraTextBox4);
            this.Controls.Add(this.ExtraCheckBox4);
            this.Controls.Add(this.PersonComboBox);
            this.Controls.Add(this.PersonCheckBox);
            this.Controls.Add(this.RegistrationDateMaskedTextBox);
            this.Controls.Add(this.RegistrationDateCheckBox);
            this.Controls.Add(this.FirstSerialNumberTextBox);
            this.Controls.Add(this.FirstSerialNumberCheckBox);
            this.Controls.Add(this.ExtraTextBox3);
            this.Controls.Add(this.ExtraCheckBox3);
            this.Controls.Add(this.ExtraTextBox2);
            this.Controls.Add(this.ExtraCheckBox2);
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
            this.Controls.Add(this.SubstrateModelLabel1);
            this.Controls.Add(this.ProductNameLabel2);
            this.Controls.Add(this.ProductNameLabel1);
            this.Controls.Add(this.ProductRegistration1MenuStrip);
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
            this.FormClosing += this.ProductRegistration1Window_FormClosing;
            this.Load += this.ProductRegistration1Window_Load;
            this.ProductRegistration1MenuStrip.ResumeLayout(false);
            this.ProductRegistration1MenuStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private Button RegisterButton;
        private Button TemplateButton;
        private ComboBox CommentComboBox;
        private TextBox CommentTextBox;
        private CheckBox CommentCheckBox;
        private TextBox ExtraTextBox6;
        private CheckBox ExtraCheckBox6;
        private TextBox ExtraTextBox5;
        private CheckBox ExtraCheckBox5;
        private TextBox ExtraTextBox4;
        private CheckBox ExtraCheckBox4;
        private ComboBox PersonComboBox;
        private CheckBox PersonCheckBox;
        private MaskedTextBox RegistrationDateMaskedTextBox;
        private CheckBox RegistrationDateCheckBox;
        private TextBox FirstSerialNumberTextBox;
        private CheckBox FirstSerialNumberCheckBox;
        private TextBox ExtraTextBox3;
        private CheckBox ExtraCheckBox3;
        private TextBox ExtraTextBox2;
        private CheckBox ExtraCheckBox2;
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
        private Label SubstrateModelLabel1;
        private ToolStripMenuItem ファイルToolStripMenuItem;
        private ToolStripMenuItem 終了ToolStripMenuItem;
        private Label ProductNameLabel2;
        private ToolStripMenuItem ヘルプToolStripMenuItem;
        private ToolStripMenuItem 取得情報ToolStripMenuItem;
        private Label ProductNameLabel1;
        private MenuStrip ProductRegistration1MenuStrip;
        private Label ProductTypeLabel2;
        private Label ProductTypeLabel1;
    }
}