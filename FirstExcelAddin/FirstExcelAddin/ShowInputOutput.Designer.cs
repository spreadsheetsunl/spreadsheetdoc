namespace FirstExcelAddin
{
    partial class ShowInputOutput
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
            this.cancelInOut = new System.Windows.Forms.Button();
            this.ok_showInOut = new System.Windows.Forms.Button();
            this.showInOutBox = new System.Windows.Forms.TextBox();
            this.showInOutLabel = new System.Windows.Forms.Label();
            this.backgroundInOut = new System.Windows.Forms.CheckBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.General_DescriptionINOUT = new System.Windows.Forms.Label();
            this.textBoxShowInput = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cancelInOut
            // 
            this.cancelInOut.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelInOut.Location = new System.Drawing.Point(269, 4);
            this.cancelInOut.Name = "cancelInOut";
            this.cancelInOut.Size = new System.Drawing.Size(75, 23);
            this.cancelInOut.TabIndex = 0;
            this.cancelInOut.Text = "CANCEL";
            this.cancelInOut.UseVisualStyleBackColor = true;
            this.cancelInOut.Click += new System.EventHandler(this.cancelInOut_Click);
            // 
            // ok_showInOut
            // 
            this.ok_showInOut.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.ok_showInOut.Location = new System.Drawing.Point(188, 4);
            this.ok_showInOut.Name = "ok_showInOut";
            this.ok_showInOut.Size = new System.Drawing.Size(75, 23);
            this.ok_showInOut.TabIndex = 1;
            this.ok_showInOut.Text = "OK";
            this.ok_showInOut.UseVisualStyleBackColor = true;
            this.ok_showInOut.Click += new System.EventHandler(this.ok_showInOut_Click);
            // 
            // showInOutBox
            // 
            this.showInOutBox.Location = new System.Drawing.Point(36, 44);
            this.showInOutBox.Multiline = true;
            this.showInOutBox.Name = "showInOutBox";
            this.showInOutBox.ReadOnly = true;
            this.showInOutBox.Size = new System.Drawing.Size(292, 90);
            this.showInOutBox.TabIndex = 2;
            // 
            // showInOutLabel
            // 
            this.showInOutLabel.AutoSize = true;
            this.showInOutLabel.Location = new System.Drawing.Point(162, 18);
            this.showInOutLabel.Name = "showInOutLabel";
            this.showInOutLabel.Size = new System.Drawing.Size(35, 13);
            this.showInOutLabel.TabIndex = 3;
            this.showInOutLabel.Text = "label1";
            // 
            // backgroundInOut
            // 
            this.backgroundInOut.AutoSize = true;
            this.backgroundInOut.Location = new System.Drawing.Point(12, 283);
            this.backgroundInOut.Name = "backgroundInOut";
            this.backgroundInOut.Size = new System.Drawing.Size(80, 17);
            this.backgroundInOut.TabIndex = 4;
            this.backgroundInOut.Text = "checkBox1";
            this.backgroundInOut.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.cancelInOut);
            this.panel1.Controls.Add(this.ok_showInOut);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 306);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(356, 30);
            this.panel1.TabIndex = 5;
            // 
            // General_DescriptionINOUT
            // 
            this.General_DescriptionINOUT.AutoSize = true;
            this.General_DescriptionINOUT.Location = new System.Drawing.Point(132, 149);
            this.General_DescriptionINOUT.Name = "General_DescriptionINOUT";
            this.General_DescriptionINOUT.Size = new System.Drawing.Size(100, 13);
            this.General_DescriptionINOUT.TabIndex = 7;
            this.General_DescriptionINOUT.Text = "General Description";
            // 
            // textBoxShowInput
            // 
            this.textBoxShowInput.Location = new System.Drawing.Point(36, 175);
            this.textBoxShowInput.Multiline = true;
            this.textBoxShowInput.Name = "textBoxShowInput";
            this.textBoxShowInput.ReadOnly = true;
            this.textBoxShowInput.Size = new System.Drawing.Size(292, 102);
            this.textBoxShowInput.TabIndex = 6;
            // 
            // ShowInputOutput
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(356, 336);
            this.Controls.Add(this.General_DescriptionINOUT);
            this.Controls.Add(this.textBoxShowInput);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.backgroundInOut);
            this.Controls.Add(this.showInOutLabel);
            this.Controls.Add(this.showInOutBox);
            this.MaximizeBox = false;
            this.Name = "ShowInputOutput";
            this.Text = "ShowInputOutput";
            this.Load += new System.EventHandler(this.ShowInputOutput_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button cancelInOut;
        private System.Windows.Forms.Button ok_showInOut;
        public System.Windows.Forms.Label showInOutLabel;
        public System.Windows.Forms.CheckBox backgroundInOut;
        public System.Windows.Forms.TextBox showInOutBox;
        private System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.TextBox textBoxShowInput;
        public System.Windows.Forms.Label General_DescriptionINOUT;
    }
}