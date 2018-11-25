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
    public partial class Form1 : Form
    {
        //private AccessMail accessmail;
        private SpeechControl speaker;
        public Form1(SpeechControl speaker)
        {
            
            InitializeComponent();
            //accessmail = new AccessMail(this);
            this.speaker = speaker;
            
            
        }
                            

        private void AccesMail_Click(object sender, EventArgs e)
        {
            //accessmail.Access();
            Form2 ff = new Form2(speaker,userBox.Text,passBox.Text);
            ff.Show();
            this.Hide();
            //panel1.Show();

            //accessmail.Access();

        }

        private void Send_Click(object sender, EventArgs e)
        {
            try
            {   int s=0;

                SmtpClient client = new SmtpClient("smtp.gmail.com");
                client.EnableSsl = true;
                MailMessage message = new MailMessage(sendFrom.Text, sendTo.Text);

                message.Body = contentBox.Text;

                message.Subject = subjectBox.Text;

                client.Credentials = new System.Net.NetworkCredential(userBox.Text, passBox.Text);
                                
                if ("587" != null)
                    client.Port = System.Convert.ToInt32("587");
                                
                client.Send(message);
                s++;
                if (s > 0) {speaker.speak("Message send"); }

            }
                           
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Cannot send message: To return press Enter");
                
            }
        }

        private void ClearField_Click(object sender, EventArgs e)
        {           
            sendTo.Clear();
            sendFrom.Clear();
            userBox.Clear();
            passBox.Clear();
            subjectBox.Clear();
            contentBox.Clear();
        }
        private void Close1_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

       

        private void userBox_Enter(object sender, EventArgs e)
        {
            speaker.speak("Type your Mail ID here");
        }

        private void passBox_Enter(object sender, EventArgs e)
        {
            speaker.speak("Type your password here");
        }

        private void sendFrom_Enter(object sender, EventArgs e)
        {
            speaker.speak("Type your Email Address here");
        }

        private void sendTo_Enter(object sender, EventArgs e)
        {
            speaker.speak("Type Receiver Email Address here");
        }

        private void subjectBox_Enter(object sender, EventArgs e)
        {
            speaker.speak("Type your Email Subject here");
        }

        private void contentBox_Enter(object sender, EventArgs e)
        {
            speaker.speak("Type your Email Message here");
        }

        private void AccesMail_Enter(object sender, EventArgs e)
        {
            speaker.speak("This is Access Mail Button");
          
        }

        private void Send_Enter(object sender, EventArgs e)
        {
            speaker.speak("This is Mail Sending Button");
        }

        private void ClearField_Enter(object sender, EventArgs e)
        {
            speaker.speak("This is Mail Window Component Clear Button");
        }

        private void Close1_Enter(object sender, EventArgs e)
        {
            speaker.speak("This is Mail Window Closing Button");
        }

        
        private void userBox_TextChanged(object sender, EventArgs e)
        {
            sendFrom.Text = userBox.Text;
        }

        private void Back_Click(object sender, EventArgs e)
        {
            //panel1.Hide();
        }

      
                       

    }

        
}
