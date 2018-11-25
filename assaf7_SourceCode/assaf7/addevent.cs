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
    public partial class addevent : Form
    {
        public addevent()
        {
            InitializeComponent();
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
                DateTime date = DateTime.Now.AddMinutes(total_add_minute);

                Task temp = new Task(EventName.Text, date, EventNotes.Text);
                Task.CurrentQueue.Add(temp);
            }
            catch(Exception ex)
            {
                MessageBox.Show("", "Wrong. Hour and Minute Format Error Please try again");
                return;
            }
            this.Close();

        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        
    }
}
