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
    public partial class editevent : Form
    {
        int i;
        public editevent()
        {
            InitializeComponent();
        }

        private void editevent_Load(object sender, EventArgs e)
        {
            button1.Enabled = false;
            for (int i = 0; i < Task.CurrentQueue.Count; i++)
            {
                if (Task.CurrentQueue[i].Time > DateTime.Now)
                {
                    cmbeventsname.Items.Add(Task.CurrentQueue[i].Name);
                }
            }
        }

        private void cmbeventsname_SelectedIndexChanged(object sender, EventArgs e)
        {
            for ( i = 0; i < Task.CurrentQueue.Count; i++)
            {
                if (Task.CurrentQueue[i].Name == cmbeventsname.Text)
                {
                    dtpkr.Value = Task.CurrentQueue[i].Time;
                    txtnotes.Text = Task.CurrentQueue[i].Notes;
                    break;
                }
            }
            button1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                int hour, minute, total_add_minute;

                if (SetHour.Text == "")
                    hour = 0;
                else hour = Convert.ToInt32(SetHour.Text);
                if (SetMinute.Text == "")
                    minute = 0;
                else minute = Convert.ToInt32(SetMinute.Text);

                total_add_minute = hour * 60 + minute;

                DateTime date = dtpkr.Value.AddMinutes(total_add_minute);

                Task.CurrentQueue[i].Time = date;

                //Task.CurrentQueue[i].Name = cmbeventsname.Text;
                // Task.CurrentQueue[i].Time=dtpkr.Value;
                Task.CurrentQueue[i].Notes = txtnotes.Text;
                button1.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("", "Wrong. New Additional Hour and New Additional Minute Format Error Please try again");
                return;
            }
             this.Close();
        }
    }
}
