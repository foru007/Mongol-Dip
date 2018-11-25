using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Media;
using System.Runtime.InteropServices;
namespace assaf7
{
    public partial class alarm : Form
    {
		Events form;
        public alarm(Events form)
        {
            InitializeComponent();
			this.form = form;
		}
		//this Section for playing audio file for the alarm
		#region sound
		[System.Runtime.InteropServices.DllImport("winmm.DLL", EntryPoint = "PlaySound", SetLastError = true, CharSet = CharSet.Unicode, ThrowOnUnmappableChar = true)]
        private static extern bool PlaySound(string szSound, System.IntPtr hMod, PlaySoundFlags flags);

        [System.Flags] 
        public enum PlaySoundFlags : int
        {
            SND_SYNC = 0x0000,
            SND_ASYNC = 0x0001, 
            SND_NODEFAULT = 0x0002, 
            SND_LOOP = 0x0008, 
            SND_NOSTOP = 0x0010,
            SND_NOWAIT = 0x00000000, 
            SND_FILENAME = 0x00020000, 
            SND_RESOURCE = 0x00040004 
        }
		#endregion
		//Close button for Press Esc
		private void button1_Click(object sender, EventArgs e)
        {
            Events frmevents = new Events();
			form.Opacity = 100;
			form.button1_Click(this, null);
            this.Close();
        }
		//recive audio file from opendialogfile 
		private void alarm_Load(object sender, EventArgs e)
		{
			//if you don't choose :)
			if (form.filename == null)
			{
                System.Threading.Thread.Sleep(4000);
                form.filename = "audio.wav";
			}
			PlaySound(form.filename, new System.IntPtr(), PlaySoundFlags.SND_SYNC);
		}
		//Event to play audio file.
		void Sp_LoadCompleted(object sender, AsyncCompletedEventArgs e)
		{
			((System.Media.SoundPlayer)sender).Play();
        }
    }
}
