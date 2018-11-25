using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using SpeechBuilder;


namespace sendEmail
{
    public partial class Form2 : Form
    {
        private SpeechControl speaker;
        private AccessMail accessmail;
        int count=0;
        int mcount = 0;
        String u="";
        String p = "";
        public Form2(SpeechControl speaker, String us, String ps)
        {
            
            u = us;
            p = ps;
            InitializeComponent();
            accessmail = new AccessMail(this);
            mcount=accessmail.Access(u, p,count);
            //count = mcount;
            //accessmail.sh();
            this.speaker = speaker;

        }
             

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 f = new Form1(speaker);
            f.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            speaker.speak("From");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            count += 1;
            int a=accessmail.Access(u, p, count);
            if (a == 0) speaker.speak("Next mail");
            else if (a == 1) speaker.speak("this is last mail");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            count -= 1;
            //if (count ==0){count=mcount;}
            int a=accessmail.Access(u, p, count);
            if (a == 0) speaker.speak("Previous mail");
            else if (a == 2) speaker.speak("this is First mail");
        }

        private void button3_Enter(object sender, EventArgs e)
        {
            speaker.speak("This is Next Mail Inbox Button");
        }

        private void button4_Enter(object sender, EventArgs e)
        {
            speaker.speak("This is Previous Mail Inbox Button");
        }
    }
}
