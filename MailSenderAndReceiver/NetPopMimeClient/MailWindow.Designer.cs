namespace NetPopMimeClient
{
    partial class MailWindow
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Send_btn = new System.Windows.Forms.Button();
            this.singOut_btn = new System.Windows.Forms.Button();
            this.inbox_btn = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel6 = new System.Windows.Forms.Panel();
            this.body_textBox = new System.Windows.Forms.TextBox();
            this.panel5 = new System.Windows.Forms.Panel();
            this.ok_btn = new System.Windows.Forms.Button();
            this.cancel_btn = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.attach_btn = new System.Windows.Forms.Button();
            this.attachment_textBox = new System.Windows.Forms.TextBox();
            this.subject_textBox = new System.Windows.Forms.TextBox();
            this.to_textBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel6.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(96, 475);
            this.panel1.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.groupBox1.Controls.Add(this.Send_btn);
            this.groupBox1.Controls.Add(this.singOut_btn);
            this.groupBox1.Controls.Add(this.inbox_btn);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(96, 475);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Mail Control";
            // 
            // Send_btn
            // 
            this.Send_btn.Location = new System.Drawing.Point(6, 105);
            this.Send_btn.Name = "Send_btn";
            this.Send_btn.Size = new System.Drawing.Size(75, 23);
            this.Send_btn.TabIndex = 2;
            this.Send_btn.Text = "Send mail";
            this.Send_btn.UseVisualStyleBackColor = true;
            // 
            // singOut_btn
            // 
            this.singOut_btn.Location = new System.Drawing.Point(6, 47);
            this.singOut_btn.Name = "singOut_btn";
            this.singOut_btn.Size = new System.Drawing.Size(75, 23);
            this.singOut_btn.TabIndex = 0;
            this.singOut_btn.Text = "Sign Out";
            this.singOut_btn.UseVisualStyleBackColor = true;
            this.singOut_btn.Click += new System.EventHandler(this.singOut_btn_Click);
            // 
            // inbox_btn
            // 
            this.inbox_btn.Location = new System.Drawing.Point(6, 76);
            this.inbox_btn.Name = "inbox_btn";
            this.inbox_btn.Size = new System.Drawing.Size(75, 23);
            this.inbox_btn.TabIndex = 1;
            this.inbox_btn.Text = "Inbox";
            this.inbox_btn.UseVisualStyleBackColor = true;
            this.inbox_btn.Click += new System.EventHandler(this.inbox_btn_Click);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.panel2.Controls.Add(this.panel4);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(96, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(708, 475);
            this.panel2.TabIndex = 1;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.panel6);
            this.panel4.Controls.Add(this.panel5);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(0, 125);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(708, 350);
            this.panel4.TabIndex = 1;
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.body_textBox);
            this.panel6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel6.Location = new System.Drawing.Point(0, 0);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(708, 312);
            this.panel6.TabIndex = 1;
            // 
            // body_textBox
            // 
            this.body_textBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.body_textBox.Location = new System.Drawing.Point(0, 0);
            this.body_textBox.Multiline = true;
            this.body_textBox.Name = "body_textBox";
            this.body_textBox.Size = new System.Drawing.Size(708, 312);
            this.body_textBox.TabIndex = 7;
            this.body_textBox.Enter += new System.EventHandler(this.body_textBox_Enter);
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.ok_btn);
            this.panel5.Controls.Add(this.cancel_btn);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel5.Location = new System.Drawing.Point(0, 312);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(708, 38);
            this.panel5.TabIndex = 3;
            // 
            // ok_btn
            // 
            this.ok_btn.Location = new System.Drawing.Point(534, 6);
            this.ok_btn.Name = "ok_btn";
            this.ok_btn.Size = new System.Drawing.Size(75, 23);
            this.ok_btn.TabIndex = 8;
            this.ok_btn.Text = "Ok";
            this.ok_btn.UseVisualStyleBackColor = true;
            this.ok_btn.Click += new System.EventHandler(this.ok_btn_Click);
            // 
            // cancel_btn
            // 
            this.cancel_btn.Location = new System.Drawing.Point(626, 6);
            this.cancel_btn.Name = "cancel_btn";
            this.cancel_btn.Size = new System.Drawing.Size(75, 23);
            this.cancel_btn.TabIndex = 9;
            this.cancel_btn.Text = "Cancel";
            this.cancel_btn.UseVisualStyleBackColor = true;
            this.cancel_btn.Click += new System.EventHandler(this.cancel_btn_Click);
            this.cancel_btn.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cancel_btn_KeyDown);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.attach_btn);
            this.panel3.Controls.Add(this.attachment_textBox);
            this.panel3.Controls.Add(this.subject_textBox);
            this.panel3.Controls.Add(this.to_textBox);
            this.panel3.Controls.Add(this.label3);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Controls.Add(this.label1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.panel3.Size = new System.Drawing.Size(708, 125);
            this.panel3.TabIndex = 1;
            this.panel3.Paint += new System.Windows.Forms.PaintEventHandler(this.panel3_Paint);
            // 
            // attach_btn
            // 
            this.attach_btn.Location = new System.Drawing.Point(617, 78);
            this.attach_btn.Name = "attach_btn";
            this.attach_btn.Size = new System.Drawing.Size(75, 23);
            this.attach_btn.TabIndex = 5;
            this.attach_btn.Text = "Attach";
            this.attach_btn.UseVisualStyleBackColor = true;
            // 
            // attachment_textBox
            // 
            this.attachment_textBox.Enabled = false;
            this.attachment_textBox.Location = new System.Drawing.Point(90, 78);
            this.attachment_textBox.Name = "attachment_textBox";
            this.attachment_textBox.Size = new System.Drawing.Size(492, 20);
            this.attachment_textBox.TabIndex = 6;
            // 
            // subject_textBox
            // 
            this.subject_textBox.Location = new System.Drawing.Point(90, 52);
            this.subject_textBox.Name = "subject_textBox";
            this.subject_textBox.Size = new System.Drawing.Size(492, 20);
            this.subject_textBox.TabIndex = 4;
            this.subject_textBox.Enter += new System.EventHandler(this.subject_textBox_Enter);
            // 
            // to_textBox
            // 
            this.to_textBox.Location = new System.Drawing.Point(90, 23);
            this.to_textBox.Name = "to_textBox";
            this.to_textBox.Size = new System.Drawing.Size(492, 20);
            this.to_textBox.TabIndex = 3;
            this.to_textBox.Enter += new System.EventHandler(this.to_textBox_Enter);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(19, 76);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Attachment";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Subject";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 23);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label1.Size = new System.Drawing.Size(20, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "To";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // MailWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(804, 475);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "MailWindow";
            this.Text = "MailWindow";
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button Send_btn;
        private System.Windows.Forms.Button singOut_btn;
        private System.Windows.Forms.Button inbox_btn;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button attach_btn;
        private System.Windows.Forms.TextBox attachment_textBox;
        private System.Windows.Forms.TextBox subject_textBox;
        private System.Windows.Forms.TextBox to_textBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.TextBox body_textBox;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Button ok_btn;
        private System.Windows.Forms.Button cancel_btn;
    }
}