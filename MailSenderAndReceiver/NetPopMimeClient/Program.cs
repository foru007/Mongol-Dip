using System;
using System.Collections.Generic;
using System.Text;
using Net.Mail;
using Net.Mime;
using System.Net.NetworkInformation;
using System.Windows.Forms;

namespace NetPopMimeClient
{
    internal class Program
    {

        //private static void Main(string[] args)
        //{
        //    //InternetConnection iconn = new InternetConnection();

        //    //if (iconn.isAvailable())
        //    //{
        //    //    Console.WriteLine("yes");
        //    MailUser mailUser = new MailUser();

        //    //    Console.WriteLine(inbox.messageFrom(1));
        //    //    Console.WriteLine(inbox.messageSubject(1));
        //    //    Console.WriteLine(inbox.messageDeliveryDate(1));
        //    //    Console.WriteLine(inbox.messagebody(1));

        //    //    //MailSender m = new MailSender("alamgir_sustcse@yahoo.com", "alamgir_sustcse@yahoo.com", "shahana", "Subject2", "This is body text");
        //    //    //m.send();
        //    //    Console.WriteLine("sending finished");
        //    //}
        //    //else
        //    //{
        //    //    Console.WriteLine("No");
        //    //}

            
        //    LogRW log = new LogRW();
        //    LogInForm login = new LogInForm(mailUser);

        //    //log.logWriter();
        //    //log.logReader();

        //    //if (log.logExist())
        //    //{
        //    //    Console.WriteLine("Exist ");
        //    //    log.logDelete();
        //    //}
        //    //else
        //    //{
        //    //    Console.WriteLine("Not exist");
        //    //}

        //    Console.ReadLine();
        //}

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            

            
            ////log.logReader();
            //LogRW log = new LogRW();
            ////log.logDelete();

            //if (log.logExist())
            //{
            //    MailUser mailUser = new MailUser();

            //    log.logReader();
            //    mailUser.isValidUser(log.logUser(), log.logPassWord());

            //    MailWindow mailWindow = new MailWindow( mailUser );
            //    Application.Run(mailWindow);
                

            //    Console.WriteLine("Exist ");
            //}
            //else
            //{
            //    MailUser mailUser = new MailUser();
            //    LogInForm login = new LogInForm(mailUser);
            //    Application.Run(login);

            //    Console.WriteLine("Not exist");
            //}

            //ReceiverForm receive = new ReceiverForm();
            //receive.Show();
            //Application.Run( new ReceiverForm() );

            //MailAgent mail = new MailAgent();
            //mail.mailSender();

            Console.ReadLine();

            
        }
    }
}
