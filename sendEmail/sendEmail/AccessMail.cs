using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Net.Mail;
using System.Net.NetworkInformation;
using System.Net.Security;
using System.Windows.Forms;
using System.Web;
using Aspose.Network.Imap;
using Aspose.Network.Mail;
using SpeechBuilder;

namespace sendEmail
{
    class AccessMail
    {
        private SpeechControl speaker;
        private DataTable dt;
        private Form2 m;
        private int messcount=0;
        private int di=0;
        //private Form2 f2;
        public AccessMail(Form2 my)
        {
            m = my;
            //f2 = new Form2();

           
        }
        public AccessMail()
        {
         
        }
                      
        //}
        //public AccessMail(Form2 mi)
        //{
        //    f2= mi;

        //}
        public void dd()
        {
            dt = new DataTable();
            dt.DefaultView.AllowEdit.Equals(false);
            dt.Columns.Add("From");
            dt.Columns.Add("Subject");
            dt.Columns.Add("Date");
        }

        public int Access(String us, String ps,int ss)
        {
            
            try
            {
                ImapClient ccc = new ImapClient();
                ccc.Host = "imap.gmail.com";
                ccc.Port = 993;
                ccc.Username = us;
                ccc.Password = ps;
                ccc.EnableSsl = true;
                ccc.SecurityMode = ImapSslSecurityMode.Implicit;
                ccc.Connect(true);
                Console.WriteLine("Connected to IMAP server.");
                ccc.SelectFolder("inbox");

                ImapMessageInfoCollection mgs = ccc.ListMessages();
                messcount = mgs.Count();
                int s = messcount - ss;
                if (s <= 0) { s = 1; di = 1; }
                if (s > messcount) { s = messcount; di = 2; }
                MailMessage mm = ccc.FetchMessage(s);
                m.textBox1.Text = "From : " + mm.From.ToString();
                m.textBox2.Text = "Subject : " + mm.Subject.ToString();
                m.textBox3.Text = "Date : " + mm.Date.ToString();
                m.richTextBox1.Text = "Message : " + mm.Body.ToString();
                          

                //m.listBox1.DataSource = dt;
                //m.listBox1.DataSource = dt;


                //listBox1.DataBind();


                //listBox1.Items.Add();
                //listBox1.Items.Add(mgs.Count.ToString());


                ccc.Disconnect();
                //speaker.speak("Access mail Done");
                



            }
            catch (Exception ex) 
            { 
                               
            }
            return di;
        }
        //public void sh()
        //{
        //    f2.listBox1.DataSource = dt;
        //}

    }
}
