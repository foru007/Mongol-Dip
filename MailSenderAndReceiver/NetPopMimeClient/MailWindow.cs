using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using SpeechBuilder;

namespace NetPopMimeClient
{
    public partial class MailWindow : Form
    {
        private MailUser mailUser;
        private LogInForm logIn;
        private SpeechControl speaker;

        public MailWindow( MailUser mailUser, LogInForm logIn , SpeechControl speaker)
        {
            InitializeComponent();
            this.mailUser = mailUser;
            this.logIn = logIn;

            this.speaker = speaker;
        }

        public MailWindow( MailUser mailUser , SpeechControl speaker)
        {
            InitializeComponent();
            this.mailUser = mailUser;

            this.speaker = speaker;

            logIn = new LogInForm(mailUser, speaker );
        }

        private void ok_btn_Click(object sender, EventArgs e)
        {
            LogRW log = new LogRW();
            log.logReader();

            MailSender mailSender = new MailSender
                (
                    to_textBox.Text.ToString(), 
                    log.logUser(),
                    log.logPassWord(), 
                    subject_textBox.Text.ToString(),
                    body_textBox.Text.ToString()
                );

            mailSender.send();
            
            to_textBox.Text = "";
            subject_textBox.Text = "";
            body_textBox.Text = "";

        }

        private void singOut_btn_Click(object sender, EventArgs e)
        {
            LogRW log = new LogRW();
            log.logDelete();
            logIn.Show();
            this.Hide();
        }

        private void cancel_btn_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }


        private void cancel_btn_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString().Equals(Keys.Enter))
            {
                this.Dispose();
            }
        }

        private void inbox_btn_Click(object sender, EventArgs e)
        {
            ReceiverForm receive = new ReceiverForm( speaker );
            receive.Show();
            
            
        }

        private void to_textBox_Enter(object sender, EventArgs e)
        {
            speaker.speak("Type receiver address");
        }

        private void subject_textBox_Enter(object sender, EventArgs e)
        {
            speaker.speak("Type mail subject");
        }

        private void body_textBox_Enter(object sender, EventArgs e)
        {
            speaker.speak("Compose mail here");
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

    }
}
