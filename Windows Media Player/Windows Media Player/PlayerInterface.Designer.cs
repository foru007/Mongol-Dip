namespace PaglaPlayer
{
    partial class PaglaPlayerPane
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PaglaPlayerPane));
            this.playerInterface = new AxWMPLib.AxWindowsMediaPlayer();
            ((System.ComponentModel.ISupportInitialize)(this.playerInterface)).BeginInit();
            this.SuspendLayout();
            // 
            // playerInterface
            // 
            this.playerInterface.Dock = System.Windows.Forms.DockStyle.Fill;
            this.playerInterface.Enabled = true;
            this.playerInterface.Location = new System.Drawing.Point(0, 0);
            this.playerInterface.Name = "playerInterface";
            this.playerInterface.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("playerInterface.OcxState")));
            this.playerInterface.Size = new System.Drawing.Size(490, 460);
            this.playerInterface.TabIndex = 0;
            this.playerInterface.KeyDownEvent += new AxWMPLib._WMPOCXEvents_KeyDownEventHandler(this.axWindowsMediaPlayer1_KeyDownEvent);
            // 
            // PaglaPlayerPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(490, 460);
            this.Controls.Add(this.playerInterface);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "PaglaPlayerPane";
            this.Text = "Windows Media Player";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.playerInterface)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AxWMPLib.AxWindowsMediaPlayer playerInterface;
    }
}

