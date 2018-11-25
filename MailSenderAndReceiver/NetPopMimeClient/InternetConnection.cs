using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.NetworkInformation;

namespace NetPopMimeClient
{
    public class InternetConnection
    {
        public bool isAvailable()
        {
            String[] list = {"www.mail.yahoo.com", "www.gmail.com"};
            
            Ping ping = new Ping();
            PingReply pingReply;

            int pingReturn = 0;
            
            try
            {
                foreach (String site in list)
                {
                    pingReply = ping.Send(site);
                    if (pingReply.Status == IPStatus.Success)
                    {
                        pingReturn++;
                    }
                }
            }
            catch( Exception )
            {
                return false;
            }

            if (pingReturn == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
            
        }
    }
}
