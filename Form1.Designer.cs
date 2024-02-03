namespace WinFormsReadExcelTavetPhuKien
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private class ComboboxItem
        {
            public string Text { get; set; }
            public object Value { get; set; }

            public override string ToString()
            {
                return Text;
            }
        }

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            button1 = new Button();
            richTextBox1 = new RichTextBox();
            comboBox1 = new ComboBox();
            SuspendLayout();
            // 
            // button1
            // 
            button1.AutoSize = true;
            button1.BackColor = Color.Fuchsia;
            button1.FlatStyle = FlatStyle.Flat;
            button1.Font = new Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point);
            button1.Location = new Point(544, 127);
            button1.Name = "button1";
            button1.Size = new Size(186, 35);
            button1.TabIndex = 0;
            button1.Text = "Chọn File Excel";
            button1.UseVisualStyleBackColor = false;
            button1.Click += button1_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.BackColor = SystemColors.ScrollBar;
            richTextBox1.Location = new Point(70, 243);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(1178, 151);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "";
            // 
            // comboBox1
            // 
            comboBox1.BackColor = Color.Fuchsia;
            comboBox1.DropDownHeight = 156;
            comboBox1.FlatStyle = FlatStyle.Flat;
            comboBox1.ForeColor = Color.White;
            comboBox1.FormattingEnabled = true;
            comboBox1.IntegralHeight = false;
            comboBox1.ItemHeight = 15;
            comboBox1.Location = new Point(544, 75);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(186, 23);
            comboBox1.TabIndex = 2;
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Purple;
            ClientSize = new Size(1326, 450);
            Controls.Add(comboBox1);
            Controls.Add(richTextBox1);
            Controls.Add(button1);
            ForeColor = Color.White;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Tool Xuất Tà vẹt Phụ kiện";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private RichTextBox richTextBox1;
        private ComboBox comboBox1;
    }
}
