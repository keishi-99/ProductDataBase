namespace ProductDatabase.Other {
    partial class InputDialog1 {
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
            this.label1 = new Label();
            this.label2 = new Label();
            this.TemperatureTextBox = new TextBox();
            this.HumidityTextBox = new TextBox();
            this.SuspendLayout();
            // 
            // OKButton
            // 
            this.OKButton.Location = new Point(55, 76);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new Size(75, 23);
            this.OKButton.TabIndex = 0;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Click += this.OKButton_Click;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new Point(47, 21);
            this.label1.Name = "label1";
            this.label1.Size = new Size(31, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "温度";
            this.label1.TextAlign = ContentAlignment.TopRight;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new Point(125, 21);
            this.label2.Name = "label2";
            this.label2.Size = new Size(31, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "湿度";
            this.label2.TextAlign = ContentAlignment.TopRight;
            // 
            // TemperatureTextBox
            // 
            this.TemperatureTextBox.Location = new Point(28, 39);
            this.TemperatureTextBox.Name = "TemperatureTextBox";
            this.TemperatureTextBox.Size = new Size(50, 23);
            this.TemperatureTextBox.TabIndex = 3;
            this.TemperatureTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // HumidityTextBox
            // 
            this.HumidityTextBox.Location = new Point(106, 39);
            this.HumidityTextBox.Name = "HumidityTextBox";
            this.HumidityTextBox.Size = new Size(50, 23);
            this.HumidityTextBox.TabIndex = 4;
            this.HumidityTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // InputDialog1
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(184, 111);
            this.Controls.Add(this.HumidityTextBox);
            this.Controls.Add(this.TemperatureTextBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.OKButton);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InputDialog1";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Load += this.InputDialog1_Load;
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private Button OKButton;
        private Label label1;
        private Label label2;
        private TextBox TemperatureTextBox;
        private TextBox HumidityTextBox;
    }
}