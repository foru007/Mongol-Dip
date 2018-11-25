using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SpeechBuilder;

namespace NetPopMimeClient
{
    public partial class LogInForm : Form
    {
        private MailUser mailUser;
        private SpeechControl speaker;

        public LogInForm( MailUser mailUser , SpeechControl speaker)
        {
            InitializeComponent();
            this.mailUser = mailUser;
            this.speaker = speaker;
            speaker.speak("Type your mail address here");
            
        }
        

        private void cancel_btn_Click(object sender, EventArgs e)
        {

            this.Dispose();
        }

        private void sign_btn_Click(object sender, EventArgs e)
        {
            String userName = user_textBox.Text.ToString();
            String passWord = pass_textBox.Text.ToString();
            InternetConnection inetConnection = new InternetConnection();

            if (!inetConnection.isAvailable())
            {
                MessageBox.Show("Internet connection is not available at this moment!!!!", "Error Message", MessageBoxButtons.OK);
                this.Dispose();
            }
            if (mailUser.isValidUser(userName, passWord))
            {
                LogRW log = new LogRW();
                log.logWriter(userName, passWord);
                
                this.Hide();
                user_textBox.Text = "";
                pass_textBox.Text = "";

                MailWindow mailWindow = new MailWindow(mailUser, this, speaker );
                mailWindow.Show();

            }
            else
            {
                //MessageBox.Show(user_textBox.Text.ToString() + " " + pass_textBox.Text.ToString());
                MessageBox.Show("User name or password is invalid!!!!", "Error Message", MessageBoxButtons.OK);
            }
        }

       

        private void sign_btn_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString().Equals(Keys.Enter))
            {
                String userName = user_textBox.Text.ToString();
                String passWord = pass_textBox.ToString();
                InternetConnection inetConnection = new InternetConnection();

                if (!inetConnection.isAvailable())
                {
                    MessageBox.Show("Internet connection is not available at this moment!!!!", "Error Message", MessageBoxButtons.OK);
                    this.Dispose();
                }
                if (mailUser.isValidUser(userName, passWord))
                {
                    LogRW log = new LogRW();
                    log.logWriter(userName, passWord);
                    
                    this.Dispose();
                }
                else
                {
                    //MessageBox.Show(user_textBox.ToString() + " " + pass_textBox.ToString());
                    MessageBox.Show("User name or password is invalid!!!!", "Error Message", MessageBoxButtons.OK);
                }
            }
            
        }

        private void cancel_btn_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString().Equals(Keys.Enter))
            {
                this.Dispose();
            }
        }

        private void user_textBox_Enter(object sender, EventArgs e)
        {
            speaker.speak("Type your mail address here");
        }

        private void pass_textBox_Enter(object sender, EventArgs e)
        {
            speaker.speak("Type your password here");
        }

        private void pass_textBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void user_textBox_TextChanged(object sender, EventArgs e)
        {

        }

        
    }
}
