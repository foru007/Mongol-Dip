namespace Search
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
            this.urlBox = new System.Windows.Forms.TextBox();
            this.search = new System.Windows.Forms.TextBox();
            this.searchbutton1 = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.SuspendLayout();
            // 
            // urlBox
            // 
            this.urlBox.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.urlBox.Location = new System.Drawing.Point(1, 1);
            this.urlBox.Name = "urlBox";
            this.urlBox.ReadOnly = true;
            this.urlBox.Size = new System.Drawing.Size(794, 20);
            this.urlBox.TabIndex = 0;
            // 
            // search
            // 
            this.search.Location = new System.Drawing.Point(1, 27);
            this.search.Name = "search";
            this.search.Size = new System.Drawing.Size(236, 20);
            this.search.TabIndex = 1;
            this.search.Enter += new System.EventHandler(this.search_Enter);
            // 
            // searchbutton1
            // 
            this.searchbutton1.Location = new System.Drawing.Point(291, 26);
            this.searchbutton1.Name = "searchbutton1";
            this.searchbutton1.Size = new System.Drawing.Size(440, 23);
            this.searchbutton1.TabIndex = 2;
            this.searchbutton1.Text = "Search Button Press Enter For Start Searching & Press Any other Key to Stop Searc" +
                "hing";
            this.searchbutton1.UseVisualStyleBackColor = true;
            this.searchbutton1.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.searchbutton1_PreviewKeyDown);
            this.searchbutton1.Click += new System.EventHandler(this.searchbutton1_Click);
            // 
            // listView1
            // 
            this.listView1.BackColor = System.Drawing.Color.White;
            this.listView1.Location = new System.Drawing.Point(-6, 56);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(801, 302);
            this.listView1.TabIndex = 4;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
            this.listView1.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.listView1_ItemSelectionChanged);
            this.listView1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.listView1_KeyDown);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(795, 351);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.searchbutton1);
            this.Controls.Add(this.search);
            this.Controls.Add(this.urlBox);
            this.Name = "Form1";
            this.Text = "File Search Window";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox urlBox;
        private System.Windows.Forms.TextBox search;
        private System.Windows.Forms.Button searchbutton1;
        private System.Windows.Forms.ListView listView1;
    }
}

