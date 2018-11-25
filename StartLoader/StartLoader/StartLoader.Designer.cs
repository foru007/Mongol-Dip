namespace StartLoader
{
    partial class LoaderForm
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
            this.components = new System.ComponentModel.Container();
            this.loadingPanel = new System.Windows.Forms.Panel();
            this.opacityTimer = new System.Windows.Forms.Timer(this.components);
            this.loadingPanelTimer = new System.Windows.Forms.Timer(this.components);
            this.timeRemaining_label = new System.Windows.Forms.Label();
            this.timeRemainingTimer = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // loadingPanel
            // 
            this.loadingPanel.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.loadingPanel.Location = new System.Drawing.Point(102, 130);
            this.loadingPanel.Name = "loadingPanel";
            this.loadingPanel.Size = new System.Drawing.Size(319, 31);
            this.loadingPanel.TabIndex = 0;
            this.loadingPanel.Paint += new System.Windows.Forms.PaintEventHandler(this.loadingPanel_Paint);
            // 
            // opacityTimer
            // 
            this.opacityTimer.Enabled = true;
            this.opacityTimer.Tick += new System.EventHandler(this.opacityTimer_Tick);
            // 
            // loadingPanelTimer
            // 
            this.loadingPanelTimer.Enabled = true;
            this.loadingPanelTimer.Tick += new System.EventHandler(this.loadingPanelTimer_Tick);
            // 
            // timeRemaining_label
            // 
            this.timeRemaining_label.AutoSize = true;
            this.timeRemaining_label.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.timeRemaining_label.Location = new System.Drawing.Point(106, 175);
            this.timeRemaining_label.Name = "timeRemaining_label";
            this.timeRemaining_label.Size = new System.Drawing.Size(83, 13);
            this.timeRemaining_label.TabIndex = 1;
            this.timeRemaining_label.Text = "Time Remaining";
            // 
            // timeRemainingTimer
            // 
            this.timeRemainingTimer.Enabled = true;
            this.timeRemainingTimer.Interval = 300;
            this.timeRemainingTimer.Tick += new System.EventHandler(this.timeRemainingTimer_Tick);
            // 
            // LoaderForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.BackgroundImage = global::StartLoader.Properties.Resources.screen;
            this.ClientSize = new System.Drawing.Size(440, 250);
            this.Controls.Add(this.timeRemaining_label);
            this.Controls.Add(this.loadingPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "LoaderForm";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel loadingPanel;
        private System.Windows.Forms.Timer opacityTimer;
        private System.Windows.Forms.Timer loadingPanelTimer;
        private System.Windows.Forms.Label timeRemaining_label;
        private System.Windows.Forms.Timer timeRemainingTimer;
    }
}

