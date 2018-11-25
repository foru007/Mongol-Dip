using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using Net.Mail;
using Net.Mime;
using SpeechBuilder;

namespace NetPopMimeClient
{
    /// <summary>
    /// এই ক্লাস থেকে জি মেইল থেকে ইনবক্স রিড করে,
    /// কয়টি আনরিড মেসেজ আছে,
    /// মেসেজ এর সাবজেক্ট কি,
    /// কে পাঠিয়েছে,
    /// এবং মেসেজে কি লেখা আছে,
    /// তা জানা যাবে
    /// </summary>
    public class MailUser
    {

        private Pop3Client client;

        private ArrayList from;
        private ArrayList subject;
        private ArrayList attachMentCount;
        private ArrayList body;
        private ArrayList dateOfArrival;
        private int nMessage;

        private SpeechControl speaker;


        private String userName = "";
        private String passWord = "";

        //জিমেইলের জন্য ডিফল্ট সারভার নেইম এবং পোট্র নাম্বার
        /******************************************/
        private String popServer = "pop.gmail.com";
        private int popPort = 995;
        /******************************************/

        public MailUser( SpeechControl speaker )
        {
            from = new ArrayList();
            subject = new ArrayList();
            attachMentCount = new ArrayList();
            body = new ArrayList();
            dateOfArrival = new ArrayList();

            this.speaker = speaker;

            nMessage = 0;

        }


        public bool isValidUser(String userName, String passWord)
        {
            this.userName = userName;
            this.passWord = passWord;
            try
            {
                client = new Pop3Client(popServer, popPort, true, userName, passWord);
                client.Authenticate();
            }
            catch( Exception )
            {
                //Console.WriteLine( "User is not valid ");
                return false;
            }
            return true;
        }

        public void readAllMessage()
        {
            
            //client.Trace += new Action<string>(Console.WriteLine);

            //connects to Pop3 Server, Executes POP3 USER and PASS
            try
            {
                //client.Authenticate();

                //
                client.Stat();

                //সব মেসেজ রিড করে তা store করা।
                //from
                //subject
                //no of attachment
                //date of arrival
                //body

                //**********************************************************//
                foreach (Pop3ListItem item in client.List())
                {
                    MailMessageEx message = client.RetrMailMessageEx(item);

                    from.Add(message.From);
                    subject.Add(message.Subject);
                    attachMentCount.Add(message.Attachments.Count);
                    body.Add(message.Body);
                    dateOfArrival.Add(message.DeliveryDate);

                    nMessage++;

                    client.Dele(item);

                }
                //**********************************************************//

                client.Noop();
                client.Rset();
                client.Quit();
            }
            catch (Exception)
            {
                Console.WriteLine("Cannot read at this moment");
            }
            
            //Console.WriteLine("Finished");
        }
        /// <summary>
        /// আনরিড মেসেজের সংখ্যা জানা
        /// </summary>
        /// <returns></returns>
        public int noOfUnreadMessage()
        {
            return nMessage;
        }


        /// <summary>
        /// কে মেসেজ টি পাঠিয়েছে?
        /// </summary>
        /// <param name="pos"></param>
        /// <returns></returns>
        public String messageFrom( int pos )
        {
            if (pos >= 0 && pos < nMessage)
            {
                return from[pos].ToString();
            }
            return null;
        }

        /// <summary>
        /// মেসেজের বিসয় জানা
        /// </summary>
        /// <param name="pos"></param>
        /// <returns></returns>
        public String messageSubject(int pos)
        {
            if (pos >= 0 && pos < nMessage)
            {
                return subject[pos].ToString();
            }
            return null;
        }

        /// <summary>
        /// মেসেজ টি কি?
        /// </summary>
        /// <param name="pos"></param>
        /// <returns></returns>
        public String messagebody(int pos)
        {
            if (pos >= 0 && pos < nMessage)
            {
                return body[pos].ToString();
            }
            return null;
        }

        /// <summary>
        /// মেসেজ পাওয়ার সময়
        /// </summary>
        /// <param name="pos"></param>
        /// <returns></returns>
        public String messageDeliveryDate(int pos)
        {
            if (pos >= 1 && pos <= nMessage)
            {
                return dateOfArrival[pos - 1].ToString();
            }
            return null;
        }

        /// <summary>
        /// মেসেজের সাথে কয়টি attachment আছে
        /// </summary>
        /// <param name="pos"></param>
        /// <returns></returns>
        public String messageAttachment(int pos)
        {
            if (pos >= 0 && pos < nMessage)
            {
                return attachMentCount[pos].ToString();
            }
            return null;
        }

        public int getPosOfMessage( String message )
        {
            int i = 0;

            foreach(String msg in subject )
            {
                if( msg.Equals( message ) )
                {
                    return i;
                }
                i++;
            }
            return i;
        }

    }
}
