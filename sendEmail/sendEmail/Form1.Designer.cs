namespace sendEmail
{
    partial class Form1
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
            this.label3 = new System.Windows.Forms.Label();
            this.subjectBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.sendTo = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.contentBox = new System.Windows.Forms.TextBox();
            this.Send = new System.Windows.Forms.Button();
            this.ClearField = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.passBox = new System.Windows.Forms.TextBox();
            this.userBox = new System.Windows.Forms.TextBox();
            this.sendFrom = new System.Windows.Forms.TextBox();
            this.AccesMail = new System.Windows.Forms.Button();
            this.Close1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 155);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Subject:";
            // 
            // subjectBox
            // 
            this.subjectBox.Location = new System.Drawing.Point(114, 155);
            this.subjectBox.Name = "subjectBox";
            this.subjectBox.Size = new System.Drawing.Size(318, 20);
            this.subjectBox.TabIndex = 6;
            this.subjectBox.Enter += new System.EventHandler(this.subjectBox_Enter);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(22, 116);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Send To:";
            // 
            // sendTo
            // 
            this.sendTo.Location = new System.Drawing.Point(114, 116);
            this.sendTo.Name = "sendTo";
            this.sendTo.Size = new System.Drawing.Size(318, 20);
            this.sendTo.TabIndex = 5;
            this.sendTo.Enter += new System.EventHandler(this.sendTo_Enter);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(21, 198);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(52, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Contents:";
            // 
            // contentBox
            // 
            this.contentBox.AcceptsReturn = true;
            this.contentBox.Location = new System.Drawing.Point(114, 198);
            this.contentBox.Multiline = true;
            this.contentBox.Name = "contentBox";
            this.contentBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.contentBox.Size = new System.Drawing.Size(318, 140);
            this.contentBox.TabIndex = 7;
            this.contentBox.Enter += new System.EventHandler(this.contentBox_Enter);
            // 
            // Send
            // 
            this.Send.Location = new System.Drawing.Point(181, 351);
            this.Send.Name = "Send";
            this.Send.Size = new System.Drawing.Size(75, 23);
            this.Send.TabIndex = 9;
            this.Send.Text = "Send";
            this.Send.UseVisualStyleBackColor = true;
            this.Send.Click += new System.EventHandler(this.Send_Click);
            this.Send.Enter += new System.EventHandler(this.Send_Enter);
            // 
            // ClearField
            // 
            this.ClearField.Location = new System.Drawing.Point(285, 351);
            this.ClearField.Name = "ClearField";
            this.ClearField.Size = new System.Drawing.Size(75, 23);
            this.ClearField.TabIndex = 10;
            this.ClearField.Text = "Clear Fields";
            this.ClearField.UseVisualStyleBackColor = true;
            this.ClearField.Click += new System.EventHandler(this.ClearField_Click);
            this.ClearField.Enter += new System.EventHandler(this.ClearField_Enter);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(22, 78);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(61, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "Send From:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(84, 12);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(87, 17);
            this.label7.TabIndex = 1;
            this.label7.Text = "User Name :";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(84, 45);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(77, 17);
            this.label8.TabIndex = 14;
            this.label8.Text = "Password :";
            // 
            // passBox
            // 
            this.passBox.Location = new System.Drawing.Point(181, 45);
            this.passBox.Name = "passBox";
            this.passBox.PasswordChar = '$';
            this.passBox.Size = new System.Drawing.Size(162, 20);
            this.passBox.TabIndex = 3;
            this.passBox.Enter += new System.EventHandler(this.passBox_Enter);
            // 
            // userBox
            // 
            this.userBox.Location = new System.Drawing.Point(181, 12);
            this.userBox.Name = "userBox";
            this.userBox.Size = new System.Drawing.Size(162, 20);
            this.userBox.TabIndex = 2;
            this.userBox.TextChanged += new System.EventHandler(this.userBox_TextChanged);
            this.userBox.Enter += new System.EventHandler(this.userBox_Enter);
            // 
            // sendFrom
            // 
            this.sendFrom.Location = new System.Drawing.Point(114, 78);
            this.sendFrom.Name = "sendFrom";
            this.sendFrom.Size = new System.Drawing.Size(318, 20);
            this.sendFrom.TabIndex = 4;
            this.sendFrom.Enter += new System.EventHandler(this.sendFrom_Enter);
            // 
            // AccesMail
            // 
            this.AccesMail.Location = new System.Drawing.Point(12, 351);
            this.AccesMail.Name = "AccesMail";
            this.AccesMail.Size = new System.Drawing.Size(142, 23);
            this.AccesMail.TabIndex = 8;
            this.AccesMail.Text = "Access Mail";
            this.AccesMail.UseVisualStyleBackColor = true;
            this.AccesMail.Click += new System.EventHandler(this.AccesMail_Click);
            this.AccesMail.Enter += new System.EventHandler(this.AccesMail_Enter);
            // 
            // Close1
            // 
            this.Close1.Location = new System.Drawing.Point(384, 351);
            this.Close1.Name = "Close1";
            this.Close1.Size = new System.Drawing.Size(75, 23);
            this.Close1.TabIndex = 11;
            this.Close1.Text = "Close";
            this.Close1.UseVisualStyleBackColor = true;
            this.Close1.Click += new System.EventHandler(this.Close1_Click);
            this.Close1.Enter += new System.EventHandler(this.Close1_Enter);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.ClientSize = new System.Drawing.Size(471, 386);
            this.Controls.Add(this.Close1);
            this.Controls.Add(this.AccesMail);
            this.Controls.Add(this.sendFrom);
            this.Controls.Add(this.userBox);
            this.Controls.Add(this.passBox);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.ClearField);
            this.Controls.Add(this.Send);
            this.Controls.Add(this.contentBox);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.sendTo);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.subjectBox);
            this.Controls.Add(this.label3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Mail sending window Type your Mail ID here";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox subjectBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox sendTo;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox contentBox;
        public System.Windows.Forms.Button Send;
        public System.Windows.Forms.Button ClearField;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        public System.Windows.Forms.TextBox passBox;
        public System.Windows.Forms.TextBox userBox;
        public System.Windows.Forms.TextBox sendFrom;
        public System.Windows.Forms.Button AccesMail;
        private System.Windows.Forms.Button Close1;
    }
}

