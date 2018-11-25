namespace assaf7
{
    partial class editevent
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
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dtpkr = new System.Windows.Forms.DateTimePicker();
            this.cmbeventsname = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.Set_Minute = new System.Windows.Forms.Label();
            this.SetMinute = new System.Windows.Forms.TextBox();
            this.SetHour = new System.Windows.Forms.TextBox();
            this.Set_Hour = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtnotes
            // 
            this.txtnotes.Location = new System.Drawing.Point(94, 154);
            this.txtnotes.Multiline = true;
            this.txtnotes.Name = "txtnotes";
            this.txtnotes.Size = new System.Drawing.Size(183, 120);
            this.txtnotes.TabIndex = 10;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 157);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Event Note :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Old Set Event Time ";
            // 
            // dtpkr
            // 
            this.dtpkr.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpkr.Location = new System.Drawing.Point(113, 47);
            this.dtpkr.Name = "dtpkr";
            this.dtpkr.Size = new System.Drawing.Size(103, 20);
            this.dtpkr.TabIndex = 4;
            // 
            // cmbeventsname
            // 
            this.cmbeventsname.FormattingEnabled = true;
            this.cmbeventsname.Location = new System.Drawing.Point(169, 12);
            this.cmbeventsname.Name = "cmbeventsname";
            this.cmbeventsname.Size = new System.Drawing.Size(111, 21);
            this.cmbeventsname.TabIndex = 2;
            this.cmbeventsname.SelectedIndexChanged += new System.EventHandler(this.cmbeventsname_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(5, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(158, 13);
            this.label4.TabIndex = 1;
            this.label4.Text = "Press down key to Chose Event";
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Location = new System.Drawing.Point(169, 293);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(84, 23);
            this.button2.TabIndex = 0;
            this.button2.Text = "&Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(56, 293);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(84, 23);
            this.button1.TabIndex = 11;
            this.button1.Text = "&Update";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Set_Minute
            // 
            this.Set_Minute.AutoSize = true;
            this.Set_Minute.Location = new System.Drawing.Point(20, 122);
            this.Set_Minute.Name = "Set_Minute";
            this.Set_Minute.Size = new System.Drawing.Size(116, 13);
            this.Set_Minute.TabIndex = 7;
            this.Set_Minute.Text = "New Additional  Minute";
            // 
            // SetMinute
            // 
            this.SetMinute.Location = new System.Drawing.Point(139, 122);
            this.SetMinute.Name = "SetMinute";
            this.SetMinute.Size = new System.Drawing.Size(100, 20);
            this.SetMinute.TabIndex = 8;
            // 
            // SetHour
            // 
            this.SetHour.Location = new System.Drawing.Point(114, 87);
            this.SetHour.Name = "SetHour";
            this.SetHour.Size = new System.Drawing.Size(100, 20);
            this.SetHour.TabIndex = 6;
            // 
            // Set_Hour
            // 
            this.Set_Hour.AutoSize = true;
            this.Set_Hour.Location = new System.Drawing.Point(5, 90);
            this.Set_Hour.Name = "Set_Hour";
            this.Set_Hour.Size = new System.Drawing.Size(107, 13);
            this.Set_Hour.TabIndex = 5;
            this.Set_Hour.Text = "New Additional  Hour";
            // 
            // editevent
            // 
            this.AcceptButton = this.button1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.button2;
            this.ClientSize = new System.Drawing.Size(314, 435);
            this.Controls.Add(this.Set_Minute);
            this.Controls.Add(this.SetMinute);
            this.Controls.Add(this.SetHour);
            this.Controls.Add(this.Set_Hour);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmbeventsname);
            this.Controls.Add(this.txtnotes);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dtpkr);
            this.Name = "editevent";
            this.Text = "editevent";
            this.Load += new System.EventHandler(this.editevent_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtnotes;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtpkr;
        private System.Windows.Forms.ComboBox cmbeventsname;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label Set_Minute;
        private System.Windows.Forms.TextBox SetMinute;
        private System.Windows.Forms.TextBox SetHour;
        private System.Windows.Forms.Label Set_Hour;
    }
}