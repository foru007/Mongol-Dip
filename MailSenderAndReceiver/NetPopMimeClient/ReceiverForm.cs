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
    public partial class ReceiverForm : Form
    {
        MailUser user = null;
        LinkLabel[ ] linkLabel;
        private SpeechControl speaker;

        public ReceiverForm( SpeechControl speaker )
        {
            InitializeComponent();

            user = new MailUser( speaker);
            linkLabel = new LinkLabel[1000];

            this.speaker = speaker;

            settingReceiverForm();

        }

        public void settingReceiverForm()
        {
            LogRW reader = new LogRW();
            reader.logReader();
            
            try
            {
                if (user.isValidUser(reader.logUser(), reader.logPassWord()))
                {

                    user.readAllMessage();
                    int noOfMessage = user.noOfUnreadMessage();

                    for (int i = 0; i < noOfMessage; i++)
                    {
                        linkLabel[i] = new LinkLabel();
                        linkLabel[i].AutoSize = true;
                        linkLabel[i].Cursor = System.Windows.Forms.Cursors.Hand;
                        linkLabel[i].DisabledLinkColor = System.Drawing.Color.PaleGreen;
                        linkLabel[i].Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                        linkLabel[i].LinkColor = System.Drawing.Color.Black;
                        linkLabel[i].Location = new System.Drawing.Point(4, i * 20);
                        linkLabel[i].Text = user.messageSubject(i);
                        linkLabel[i].PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.linkLabel_PreviewKeyDown);
                        linkLabel[i].LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);

                        subjectPanel.Controls.Add(linkLabel[i]);
                        Console.WriteLine(user.messageSubject(i));
                    }
                    if (noOfMessage >= 1)
                    {
                        linkLabel[0].Focus();
                    }
                    else
                    {
                        speaker.speak("No unread message in your inbox. Try later");
                    }
                }
            }
            catch( Exception )
            {
            }
            //Console.WriteLine( noOfMessage );
        }

        private void linkLabel_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if( e.KeyCode.ToString().Equals("Return"))
            {
                LinkLabel label =  sender as LinkLabel;
                int pos = user.getPosOfMessage( label.Text );
                bodyTextBox.Text = user.messagebody( pos );
            }
            
        }

        private void linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel label = sender as LinkLabel;
            int pos = user.getPosOfMessage(label.Text);
            bodyTextBox.Text = user.messagebody(pos);
        }

        private void linkLabel_Enter(object sender, EventArgs e)
        {
        }

        private void close_btn_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

    }
}
