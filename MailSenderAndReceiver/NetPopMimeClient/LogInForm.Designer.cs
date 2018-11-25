namespace NetPopMimeClient
{
    partial class LogInForm
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
            this.user_label = new System.Windows.Forms.Label();
            this.pass_label = new System.Windows.Forms.Label();
            this.user_textBox = new System.Windows.Forms.TextBox();
            this.pass_textBox = new System.Windows.Forms.TextBox();
            this.sign_btn = new System.Windows.Forms.Button();
            this.cancel_btn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // user_label
            // 
            this.user_label.AutoSize = true;
            this.user_label.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.user_label.Location = new System.Drawing.Point(31, 30);
            this.user_label.Name = "user_label";
            this.user_label.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.user_label.Size = new System.Drawing.Size(87, 20);
            this.user_label.TabIndex = 0;
            this.user_label.Text = "User name";
            
            // 
            // pass_label
            // 
            this.pass_label.AutoSize = true;
            this.pass_label.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pass_label.Location = new System.Drawing.Point(31, 59);
            this.pass_label.Name = "pass_label";
            this.pass_label.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.pass_label.Size = new System.Drawing.Size(78, 20);
            this.pass_label.TabIndex = 1;
            this.pass_label.Text = "Password";
            this.pass_label.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // user_textBox
            // 
            this.user_textBox.Location = new System.Drawing.Point(124, 32);
            this.user_textBox.Name = "user_textBox";
            this.user_textBox.Size = new System.Drawing.Size(204, 20);
            this.user_textBox.TabIndex = 2;
            this.user_textBox.TextChanged += new System.EventHandler(this.user_textBox_TextChanged);
            this.user_textBox.Enter += new System.EventHandler(this.user_textBox_Enter);
            // 
            // pass_textBox
            // 
            this.pass_textBox.Location = new System.Drawing.Point(124, 61);
            this.pass_textBox.Name = "pass_textBox";
            this.pass_textBox.PasswordChar = '0';
            this.pass_textBox.Size = new System.Drawing.Size(204, 20);
            this.pass_textBox.TabIndex = 3;
            this.pass_textBox.TextChanged += new System.EventHandler(this.pass_textBox_TextChanged);
            this.pass_textBox.Enter += new System.EventHandler(this.pass_textBox_Enter);
            // 
            // sign_btn
            // 
            this.sign_btn.Location = new System.Drawing.Point(176, 112);
            this.sign_btn.Name = "sign_btn";
            this.sign_btn.Size = new System.Drawing.Size(75, 23);
            this.sign_btn.TabIndex = 4;
            this.sign_btn.Text = "Sign In";
            this.sign_btn.UseVisualStyleBackColor = true;
            this.sign_btn.Click += new System.EventHandler(this.sign_btn_Click);
            this.sign_btn.KeyDown += new System.Windows.Forms.KeyEventHandler(this.sign_btn_KeyDown);
            // 
            // cancel_btn
            // 
            this.cancel_btn.Location = new System.Drawing.Point(257, 112);
            this.cancel_btn.Name = "cancel_btn";
            this.cancel_btn.Size = new System.Drawing.Size(75, 23);
            this.cancel_btn.TabIndex = 5;
            this.cancel_btn.Text = "Cancel";
            this.cancel_btn.UseVisualStyleBackColor = true;
            this.cancel_btn.Click += new System.EventHandler(this.cancel_btn_Click);
            this.cancel_btn.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cancel_btn_KeyDown);
            // 
            // LogInForm
            // 
            this.AcceptButton = this.sign_btn;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(375, 164);
            this.ControlBox = false;
            this.Controls.Add(this.cancel_btn);
            this.Controls.Add(this.sign_btn);
            this.Controls.Add(this.pass_textBox);
            this.Controls.Add(this.user_textBox);
            this.Controls.Add(this.pass_label);
            this.Controls.Add(this.user_label);
            this.Name = "LogInForm";
            this.Text = "LogInForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label user_label;
        private System.Windows.Forms.Label pass_label;
        private System.Windows.Forms.TextBox user_textBox;
        private System.Windows.Forms.TextBox pass_textBox;
        private System.Windows.Forms.Button sign_btn;
        private System.Windows.Forms.Button cancel_btn;
    }
}