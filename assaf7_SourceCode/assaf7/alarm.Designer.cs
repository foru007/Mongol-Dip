namespace assaf7
{
    partial class alarm
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
            this.txtnotes = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtnotes
            // 
            this.txtnotes.BackColor = System.Drawing.SystemColors.Info;
            this.txtnotes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtnotes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtnotes.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.txtnotes.Location = new System.Drawing.Point(8, 25);
            this.txtnotes.Multiline = true;
            this.txtnotes.Name = "txtnotes";
            this.txtnotes.ReadOnly = true;
            this.txtnotes.Size = new System.Drawing.Size(169, 113);
            this.txtnotes.TabIndex = 2;
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button1.Location = new System.Drawing.Point(27, 193);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(60, 20);
            this.button1.TabIndex = 3;
            this.button1.Text = "cacel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // alarm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.button1;
            this.ClientSize = new System.Drawing.Size(184, 147);
            this.ControlBox = false;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtnotes);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Location = new System.Drawing.Point(600, 400);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "alarm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "We have a Task Reminder Press any key to listen the reminder (Press Esc to Exit)";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.alarm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.TextBox txtnotes;
        private System.Windows.Forms.Button button1;

    }
}