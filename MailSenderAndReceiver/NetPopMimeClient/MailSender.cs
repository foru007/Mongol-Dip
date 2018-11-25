using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;
using System.Net;
using System.Windows.Forms;

namespace NetPopMimeClient
{
    /// <summary>
    /// yahoo এবং gmail এ মেইল send করা
    /// </summary>
    class MailSender
    {
        private String to, from, passWord, subject, body;
        private int serverName = -1;

        public MailSender(String to, String from, String passWord, 
            String subject,String body)
        {
            this.to = to;
            this.from = from;
            this.passWord = passWord;
            this.subject = subject;
            this.body = body;
        }



        public void send()
        {
            MailMessage mail = new MailMessage();
            
            mail.To.Add(to);
            mail.From = new MailAddress(from);
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = true;


            SmtpClient smtp = new SmtpClient();


            checkMailServer( from );

            if (serverName == (int)MailServer.Yahoo)
            {
                smtp.Host = "smtp.mail.yahoo.com"; //Or Your SMTP Server Address
                //Console.WriteLine("yahoo");
            }
            else if (serverName == (int)MailServer.Gmail)
            {
                smtp.Host = "smtp.www.gmail.com"; //Or Your SMTP Server Address
                //gmail এর জন্য ssl enable করতে হয়
                smtp.EnableSsl = true;
                //Console.WriteLine("gmail");
            }
            else
            {
                //Console.WriteLine("others");
            }
            smtp.Credentials = new System.Net.NetworkCredential(from, passWord);

            try
            {
                smtp.Send(mail);
                MessageBox.Show("Message sent Successfully", "Success", MessageBoxButtons.OK);
            }
            catch( Exception )
            {
                //Console.WriteLine("Sending failed");

                MessageBox.Show("Cannot send mail at this moment!!!! Try again......", "Error Message", MessageBoxButtons.OK);
            }

        }

        private void checkMailServer( String name )
        {
            //মেইল server এর নাম select করা
            int index = name.IndexOf('@');
            name = name.Remove(0, index + 1);
            
            //Console.WriteLine(name);

            if (name.Equals("www.yahoo.com"))
            {
                serverName = (int)MailServer.Yahoo;
            }
            else if (name.Equals("www.gmail.com"))
            {
                serverName = (int)MailServer.Gmail;
            }
            else
            {
                serverName = (int)MailServer.Others;
            }
        }

        enum MailServer
        {
            Yahoo = 1,
            Gmail = 2,
            Others = 3
        };

    }
}
