namespace FirstExcelAddin
{
    partial class ShowRangeForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.rangeListShowRange = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.close_button = new System.Windows.Forms.Button();
            this.rangeCellListLabel = new System.Windows.Forms.Label();
            this.rangeCellListBox = new System.Windows.Forms.TextBox();
            this.GD_Label = new System.Windows.Forms.Label();
            this.GD_TextBox = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(152, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Range List";
            // 
            // rangeListShowRange
            // 
            this.rangeListShowRange.Location = new System.Drawing.Point(30, 44);
            this.rangeListShowRange.Multiline = true;
            this.rangeListShowRange.Name = "rangeListShowRange";
            this.rangeListShowRange.ReadOnly = true;
            this.rangeListShowRange.Size = new System.Drawing.Size(292, 90);
            this.rangeListShowRange.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.close_button);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 434);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(352, 30);
            this.panel1.TabIndex = 2;
            // 
            // close_button
            // 
            this.close_button.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.close_button.Location = new System.Drawing.Point(269, 4);
            this.close_button.Name = "close_button";
            this.close_button.Size = new System.Drawing.Size(75, 23);
            this.close_button.TabIndex = 0;
            this.close_button.Text = "CLOSE";
            this.close_button.UseVisualStyleBackColor = true;
            // 
            // rangeCellListLabel
            // 
            this.rangeCellListLabel.AutoSize = true;
            this.rangeCellListLabel.Location = new System.Drawing.Point(146, 153);
            this.rangeCellListLabel.Name = "rangeCellListLabel";
            this.rangeCellListLabel.Size = new System.Drawing.Size(64, 13);
            this.rangeCellListLabel.TabIndex = 3;
            this.rangeCellListLabel.Text = "Range Cells";
            // 
            // rangeCellListBox
            // 
            this.rangeCellListBox.Location = new System.Drawing.Point(30, 184);
            this.rangeCellListBox.Multiline = true;
            this.rangeCellListBox.Name = "rangeCellListBox";
            this.rangeCellListBox.ReadOnly = true;
            this.rangeCellListBox.Size = new System.Drawing.Size(292, 90);
            this.rangeCellListBox.TabIndex = 4;
            // 
            // GD_Label
            // 
            this.GD_Label.AutoSize = true;
            this.GD_Label.Location = new System.Drawing.Point(132, 286);
            this.GD_Label.Name = "GD_Label";
            this.GD_Label.Size = new System.Drawing.Size(100, 13);
            this.GD_Label.TabIndex = 5;
            this.GD_Label.Text = "General Description";
            // 
            // GD_TextBox
            // 
            this.GD_TextBox.Location = new System.Drawing.Point(30, 311);
            this.GD_TextBox.Multiline = true;
            this.GD_TextBox.Name = "GD_TextBox";
            this.GD_TextBox.ReadOnly = true;
            this.GD_TextBox.Size = new System.Drawing.Size(292, 102);
            this.GD_TextBox.TabIndex = 6;
            // 
            // ShowRangeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(352, 464);
            this.Controls.Add(this.GD_TextBox);
            this.Controls.Add(this.GD_Label);
            this.Controls.Add(this.rangeCellListBox);
            this.Controls.Add(this.rangeCellListLabel);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.rangeListShowRange);
            this.Controls.Add(this.label1);
            this.Name = "ShowRangeForm";
            this.Text = "ShowRangeForm";
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button close_button;
        private System.Windows.Forms.Label rangeCellListLabel;
        private System.Windows.Forms.Label GD_Label;
        public System.Windows.Forms.TextBox rangeListShowRange;
        public System.Windows.Forms.TextBox rangeCellListBox;
        public System.Windows.Forms.TextBox GD_TextBox;
    }
}