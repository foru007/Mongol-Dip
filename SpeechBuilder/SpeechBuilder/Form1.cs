using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Gma.UserActivityMonitor;

namespace SpeechBuilder
{
    public partial class Form1 : Form
    {
        SpeechControl speechControl;

        public Form1()
        {
            InitializeComponent();
            speechControl = new SpeechControl();
            //speechControl.selectSpeaker();
            speechControl.speak("Hello this is my pen Hello this is my pen Hello this is my pen Hello this is my pen");
            speechControl.speak("Hello this is my pen Hello this is my pen Hello this is my pen Hello this is my pen");
            speechControl.speak("Hello this is my pen Hello this is my pen Hello this is my pen Hello this is my pen");
        }

        private void start_btn_Click(object sender, EventArgs e)
        {
            HookManager.KeyDown += HookManager_KeyDown;
        }

        private void stop_btn_Click(object sender, EventArgs e)
        {
            HookManager.KeyDown -= HookManager_KeyDown;
        }

        private void HookManager_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString().Equals("S"))
            {
                speechControl.volume++;
            }
            if (e.KeyCode.ToString().Equals("T"))
            {
                speechControl.volume--;
            }
        }

        private void HookManager_KeyUp(object sender, KeyEventArgs e)
        {
            
        }

    }
}
