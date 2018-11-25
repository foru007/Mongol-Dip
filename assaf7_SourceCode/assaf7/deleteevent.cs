using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace assaf7
{
    public partial class deleteevent : Form
    {
        int i;
        public deleteevent()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void deleteevent_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < Task.CurrentQueue.Count; i++)
            {
                if (Task.CurrentQueue[i].Time > DateTime.Now)
                {
                    ChooserEventPressDownKey.Items.Add(Task.CurrentQueue[i].Name);
                }
            }
        }

        private void cmbeventsname_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (i = 0; i < Task.CurrentQueue.Count; i++)
            {
                if (Task.CurrentQueue[i].Name == ChooserEventPressDownKey.Text)
                {
                    dateTimePicker1.Value = Task.CurrentQueue[i].Time;
                    textBox1.Text = Task.CurrentQueue[i].Notes;
                    break;
                }
            }
            button1.Enabled = true;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Task.CurrentQueue.RemoveAt(i);
            button1.Enabled = false;
            this.Close();
       
		}
    }
}
