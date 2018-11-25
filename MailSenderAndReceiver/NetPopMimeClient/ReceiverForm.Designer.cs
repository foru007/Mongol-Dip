namespace NetPopMimeClient
{
    partial class ReceiverForm
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
            this.subjectPanel = new System.Windows.Forms.Panel();
            this.bodyTextBox = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.close_btn = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // subjectPanel
            // 
            this.subjectPanel.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.subjectPanel.Dock = System.Windows.Forms.DockStyle.Left;
            this.subjectPanel.Location = new System.Drawing.Point(0, 0);
            this.subjectPanel.Name = "subjectPanel";
            this.subjectPanel.Size = new System.Drawing.Size(198, 382);
            this.subjectPanel.TabIndex = 0;
            // 
            // bodyTextBox
            // 
            this.bodyTextBox.Dock = System.Windows.Forms.DockStyle.Top;
            this.bodyTextBox.Location = new System.Drawing.Point(198, 0);
            this.bodyTextBox.Multiline = true;
            this.bodyTextBox.Name = "bodyTextBox";
            this.bodyTextBox.Size = new System.Drawing.Size(512, 336);
            this.bodyTextBox.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.close_btn);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(198, 342);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(512, 40);
            this.panel1.TabIndex = 2;
            // 
            // close_btn
            // 
            this.close_btn.Location = new System.Drawing.Point(357, 3);
            this.close_btn.Name = "close_btn";
            this.close_btn.Size = new System.Drawing.Size(143, 34);
            this.close_btn.TabIndex = 0;
            this.close_btn.Text = "Close";
            this.close_btn.UseVisualStyleBackColor = true;
            this.close_btn.Click += new System.EventHandler(this.close_btn_Click);
            // 
            // ReceiverForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(710, 382);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.bodyTextBox);
            this.Controls.Add(this.subjectPanel);
            this.Name = "ReceiverForm";
            this.Text = "ReceiverForm";
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel subjectPanel;
        private System.Windows.Forms.TextBox bodyTextBox;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button close_btn;
    }
}