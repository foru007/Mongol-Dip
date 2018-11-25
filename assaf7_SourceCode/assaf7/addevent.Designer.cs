namespace assaf7
{
    partial class addevent
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
            this.Event_Note = new System.Windows.Forms.Label();
            this.EventNotes = new System.Windows.Forms.TextBox();
            this.Event_Name = new System.Windows.Forms.Label();
            this.EventName = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.Set_Minute = new System.Windows.Forms.Label();
            this.SetMinute = new System.Windows.Forms.TextBox();
            this.SetHour = new System.Windows.Forms.TextBox();
            this.Set_Hour = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Event_Note
            // 
            this.Event_Note.AutoSize = true;
            this.Event_Note.Location = new System.Drawing.Point(61, 121);
            this.Event_Note.Name = "Event_Note";
            this.Event_Note.Size = new System.Drawing.Size(61, 13);
            this.Event_Note.TabIndex = 7;
            this.Event_Note.Text = "Event Note";
            // 
            // EventNotes
            // 
            this.EventNotes.Location = new System.Drawing.Point(135, 116);
            this.EventNotes.Multiline = true;
            this.EventNotes.Name = "EventNotes";
            this.EventNotes.Size = new System.Drawing.Size(183, 107);
            this.EventNotes.TabIndex = 8;
            // 
            // Event_Name
            // 
            this.Event_Name.AutoSize = true;
            this.Event_Name.Location = new System.Drawing.Point(64, 9);
            this.Event_Name.Name = "Event_Name";
            this.Event_Name.Size = new System.Drawing.Size(72, 13);
            this.Event_Name.TabIndex = 1;
            this.Event_Name.Text = "Event Name :";
            // 
            // EventName
            // 
            this.EventName.Location = new System.Drawing.Point(153, 7);
            this.EventName.Name = "EventName";
            this.EventName.Size = new System.Drawing.Size(107, 20);
            this.EventName.TabIndex = 2;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(69, 252);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(84, 23);
            this.btnSave.TabIndex = 9;
            this.btnSave.Text = "&Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.button1_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.Location = new System.Drawing.Point(182, 252);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(84, 23);
            this.BtnCancel.TabIndex = 0;
            this.BtnCancel.Text = "&Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // Set_Minute
            // 
            this.Set_Minute.AutoSize = true;
            this.Set_Minute.Location = new System.Drawing.Point(75, 74);
            this.Set_Minute.Name = "Set_Minute";
            this.Set_Minute.Size = new System.Drawing.Size(58, 13);
            this.Set_Minute.TabIndex = 5;
            this.Set_Minute.Text = "Set Minute";
            // 
            // SetMinute
            // 
            this.SetMinute.Location = new System.Drawing.Point(156, 74);
            this.SetMinute.Name = "SetMinute";
            this.SetMinute.Size = new System.Drawing.Size(100, 20);
            this.SetMinute.TabIndex = 6;
            // 
            // SetHour
            // 
            this.SetHour.Location = new System.Drawing.Point(156, 39);
            this.SetHour.Name = "SetHour";
            this.SetHour.Size = new System.Drawing.Size(100, 20);
            this.SetHour.TabIndex = 4;
            // 
            // Set_Hour
            // 
            this.Set_Hour.AutoSize = true;
            this.Set_Hour.Location = new System.Drawing.Point(84, 42);
            this.Set_Hour.Name = "Set_Hour";
            this.Set_Hour.Size = new System.Drawing.Size(49, 13);
            this.Set_Hour.TabIndex = 3;
            this.Set_Hour.Text = "Set Hour";
            // 
            // addevent
            // 
            this.AcceptButton = this.btnSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(342, 284);
            this.Controls.Add(this.Set_Minute);
            this.Controls.Add(this.SetMinute);
            this.Controls.Add(this.SetHour);
            this.Controls.Add(this.Set_Hour);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.EventName);
            this.Controls.Add(this.EventNotes);
            this.Controls.Add(this.Event_Note);
            this.Controls.Add(this.Event_Name);
            this.Name = "addevent";
            this.Text = "addevent";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label Event_Note;
        private System.Windows.Forms.TextBox EventNotes;
        private System.Windows.Forms.Label Event_Name;
        private System.Windows.Forms.TextBox EventName;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Label Set_Minute;
        private System.Windows.Forms.TextBox SetMinute;
        private System.Windows.Forms.TextBox SetHour;
        private System.Windows.Forms.Label Set_Hour;
    }
}