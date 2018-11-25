using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Taskbar;
namespace assaf7
{
    public partial class Events : Form
    {
        List<int> values = new List<int>();
        private TaskbarManager win_taskbar = TaskbarManager.Instance;
		private int steps=0;
        private ThumbnailToolbarButton Nexttask;
        private ThumbnailToolbarButton Previoustask;
		public string filename;
        public Events()
        {
            InitializeComponent();
        }

        private void Events_Load(object sender, EventArgs e)
        {
			steps = 0;
            Task.CurrentQueue=new List<Task>();
            Open();
        }

        public void reload()
        {
            dataGridView1.Refresh();
        }

		//to open file and read all events from it and put it in list of events
        private static void Open()
        {
            Stream file = new FileStream("data.bin", FileMode.Open, FileAccess.Read);
            IFormatter formatter = new BinaryFormatter();
            Task.CurrentQueue = (List<Task>)formatter.Deserialize(file);
            file.Close();
        }
		//load events in the data grid
        public void LoadDataGrid()
        {
            dataGridView1.Rows.Clear();
             for (int i = 0; i < Task.CurrentQueue.Count; i++)
            {
                if (Task.CurrentQueue[i].Time > DateTime.Now)
                {
                    dataGridView1.Rows.Add(Task.CurrentQueue[i].Name, Task.CurrentQueue[i].Time, Task.CurrentQueue[i].Notes);
                }
            }
        }
		//Save events in file when the program close
        private static void Save()
        {
            // to serialize an object of any serializable class
            IFormatter formatter = new BinaryFormatter();
            Stream file = new FileStream("data.bin", FileMode.Create, FileAccess.Write);
            formatter.Serialize(file, Task.CurrentQueue);
            file.Close();
        }

        //the refresh button for load the time left in minutes to anew list and sort is Asc.
		//and delete the event that time has been gone.
        public void button1_Click(object sender, EventArgs e)
        {
            LoadDataGrid();
            values.Clear();
            for (int i = 0; i < Task.CurrentQueue.Count; i++)
            {
                TimeSpan temp = Task.CurrentQueue[i].Time-DateTime.Now;
				if (temp.TotalMinutes > 0)
				{
					values.Add((int)temp.TotalMinutes);
				}
				else
				{
					Task.CurrentQueue.RemoveAt(i);
					i--;
				}
            }
            values.Sort();
            cmbtimes.Items.Clear();
            for (int i = 0; i < values.Count; i++)
            {
                cmbtimes.Items.Add(values[i]);
            }
			if (this.cmbtimes.Items.Count > steps)
				this.cmbtimes.SelectedIndex = steps;
			else if (this.cmbtimes.Items.Count != 0)
				this.cmbtimes.SelectedIndex = steps = 0;
			else
			{
				this.cmbtimes.Text = "";
				progressBar1.Value = 0;
				win_taskbar.SetProgressValue(0, 1);
			}
        }
        //timer to check the time for event and make the alarm
        private void timer1_Tick(object sender, EventArgs e)
        {
            for (int i = 0; i < Task.CurrentQueue.Count; i++)
            {
                if (Task.CurrentQueue[i].Time >= DateTime.Now&&Task.CurrentQueue[i].Time<DateTime.Now.AddMilliseconds(timer1.Interval))
                {
                    alarm frmalarm = new alarm(this);
                    frmalarm.Visible = true;
					frmalarm.Activate();
					frmalarm.txtnotes.Text=Task.CurrentQueue[i].Notes;
                    this.Opacity=0;
                }
            }

			button1_Click(null, null);
        }
		//manage events (add,edit,delete) forms
        #region btn Handles
        private void btndelete_Click(object sender, EventArgs e)
        {
            deleteevent frmaddevent = new deleteevent();
            frmaddevent.ShowDialog();
            LoadDataGrid();
        }
        private void btnEdit_Click(object sender, EventArgs e)
        {
            editevent frmaddevent = new editevent();
            frmaddevent.ShowDialog();
            LoadDataGrid();

        }
        private void btnadd_Click(object sender, EventArgs e)
        {
            addevent frmaddevent = new addevent();
            frmaddevent.ShowDialog();
			button1_Click(null, null);
        }
        private void Events_FormClosing(object sender, FormClosingEventArgs e)
        {
            Save();
        }
        private void Events_FormClosed(object sender, FormClosedEventArgs e)
        {
            Save();
        }
        #endregion


		//combobox that holds the the time left in minutes for all events
		//and show events windows taskbar and progressbar for this events
        private void cmbtimes_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index=cmbtimes.SelectedIndex;
            TimeSpan t = Task.CurrentQueue[index].Time - Task.CurrentQueue[index].now;
            progressBar1.Maximum = (int)Math.Ceiling(t.TotalMinutes);
            progressBar1.Value = (int)Math.Ceiling((DateTime.Now-Task.CurrentQueue[index].now).TotalMinutes);
            win_taskbar.SetProgressValue(progressBar1.Value, (int)t.TotalMinutes);
            #region Progressbar

            int dif = (int)t.TotalMinutes;
            if (progressBar1.Value > (dif * 3 / 4))
            {
                win_taskbar.SetProgressState(TaskbarProgressBarState.Paused);
            }
            if (progressBar1.Value > (dif * 7 / 8))
            {
                win_taskbar.SetProgressState(TaskbarProgressBarState.Error);
            }
            #endregion
        }
		//make Two button in windows taskbar
		//one for next event
		//other for previous event
		private void Events_Shown(object sender, EventArgs e)
		{
			Nexttask = new ThumbnailToolbarButton(Properties.Resources.nextArrow, "Next Event");
			Nexttask.Click += new EventHandler<ThumbnailButtonClickedEventArgs>(btnNext_Click);
			Previoustask = new ThumbnailToolbarButton(Properties.Resources.prevArrow, "Previous Event");
			Previoustask.Click += new EventHandler<ThumbnailButtonClickedEventArgs>(btnPrevious_Click);
			TaskbarManager.Instance.ThumbnailToolbars.AddButtons(this.Handle, Previoustask,Nexttask);
			button1_Click(null, null);
		}
		private void btnNext_Click(object sender, EventArgs e)
		{
			steps++;
			button1_Click(null, null);
		}
		private void btnPrevious_Click(object sender, EventArgs e)
		{
			if(steps!=0)
			steps--;
			button1_Click(null, null);
		}
		//choose file for alarm you want to hear *.wav
		private void chooseFileForAlarmToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				filename = openFileDialog1.FileName;
			}
		}
		private void exitToolStripMenuItem_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
		{
			MessageBox.Show("Developer : Ahmed Assaf \n Mail : Des-life@hotmail.com", "About me ....:)");
		}
    }
}
