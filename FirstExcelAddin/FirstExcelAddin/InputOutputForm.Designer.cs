namespace FirstExcelAddin
{
    partial class InputOutputForm
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
            this.InputOutputList = new System.Windows.Forms.Label();
            this.InputOutputDescription = new System.Windows.Forms.Label();
            this.inputOutputDescriptionBox = new System.Windows.Forms.TextBox();
            this.clear_button = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.InputListBox = new System.Windows.Forms.TextBox();
            this.remove_button = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // InputOutputList
            // 
            this.InputOutputList.AutoSize = true;
            this.InputOutputList.Location = new System.Drawing.Point(162, 18);
            this.InputOutputList.Name = "InputOutputList";
            this.InputOutputList.Size = new System.Drawing.Size(50, 13);
            this.InputOutputList.TabIndex = 2;
            this.InputOutputList.Text = "Input List";
            // 
            // InputOutputDescription
            // 
            this.InputOutputDescription.AutoSize = true;
            this.InputOutputDescription.Location = new System.Drawing.Point(132, 105);
            this.InputOutputDescription.Name = "InputOutputDescription";
            this.InputOutputDescription.Size = new System.Drawing.Size(100, 13);
            this.InputOutputDescription.TabIndex = 3;
            this.InputOutputDescription.Text = "General Description";
            // 
            // inputOutputDescriptionBox
            // 
            this.inputOutputDescriptionBox.Location = new System.Drawing.Point(36, 134);
            this.inputOutputDescriptionBox.Multiline = true;
            this.inputOutputDescriptionBox.Name = "inputOutputDescriptionBox";
            this.inputOutputDescriptionBox.Size = new System.Drawing.Size(292, 67);
            this.inputOutputDescriptionBox.TabIndex = 4;
            // 
            // clear_button
            // 
            this.clear_button.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.clear_button.Location = new System.Drawing.Point(12, 266);
            this.clear_button.Name = "clear_button";
            this.clear_button.Size = new System.Drawing.Size(75, 23);
            this.clear_button.TabIndex = 31;
            this.clear_button.Text = "CLEAR";
            this.clear_button.Click += new System.EventHandler(this.clear_button_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(277, 266);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 30;
            this.cancelButton.Text = "CANCEL";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // okButton
            // 
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButton.Location = new System.Drawing.Point(196, 266);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 29;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // InputListBox
            // 
            this.InputListBox.AllowDrop = true;
            this.InputListBox.Location = new System.Drawing.Point(36, 44);
            this.InputListBox.Multiline = true;
            this.InputListBox.Name = "InputListBox";
            this.InputListBox.ReadOnly = true;
            this.InputListBox.Size = new System.Drawing.Size(292, 49);
            this.InputListBox.TabIndex = 32;
            // 
            // remove_button
            // 
            this.remove_button.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.remove_button.Enabled = false;
            this.remove_button.Location = new System.Drawing.Point(12, 237);
            this.remove_button.Name = "remove_button";
            this.remove_button.Size = new System.Drawing.Size(75, 23);
            this.remove_button.TabIndex = 33;
            this.remove_button.Text = "REMOVE";
            this.remove_button.UseVisualStyleBackColor = true;
            this.remove_button.Click += new System.EventHandler(this.remove_button_Click);
            // 
            // InputOutputForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(356, 290);
            this.Controls.Add(this.remove_button);
            this.Controls.Add(this.InputListBox);
            this.Controls.Add(this.clear_button);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.inputOutputDescriptionBox);
            this.Controls.Add(this.InputOutputDescription);
            this.Controls.Add(this.InputOutputList);
            this.MaximizeBox = false;
            this.Name = "InputOutputForm";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.InputOutputForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label InputOutputDescription;
        private System.Windows.Forms.Button clear_button;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button okButton;
        public System.Windows.Forms.TextBox inputOutputDescriptionBox;
        public System.Windows.Forms.TextBox InputListBox;
        public System.Windows.Forms.Label InputOutputList;
        public System.Windows.Forms.Button remove_button;


    }
}