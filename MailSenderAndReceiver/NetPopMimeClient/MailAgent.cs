using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Net.Mail;
using Net.Mime;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using SpeechBuilder;

namespace NetPopMimeClient
{
    public class MailAgent
    {
        private MailUser mailUser;
        private MailWindow mailWindow;
        private SpeechControl speaker;

        public MailAgent( SpeechControl speaker )
        {
            this.speaker = speaker;
        }

        public void mailSender()
        {
            //log.logReader();
            LogRW log = new LogRW();
            //log.logDelete();

            if (log.logExist())
            {
                mailUser = new MailUser( speaker );

                log.logReader();
                //mailUser.isValidUser(log.logUser(), log.logPassWord());

                mailWindow = new MailWindow(mailUser, speaker );
                //Application.Run(mailWindow);
                mailWindow.Show();
                //Console.WriteLine("Exist ");
            }
            else
            {
                MailUser mailUser = new MailUser( speaker );
                LogInForm login = new LogInForm(mailUser, speaker );
                //Application.Run(login);
                login.Show();

                //Console.WriteLine("Not exist");
            }
        }
    }
}
