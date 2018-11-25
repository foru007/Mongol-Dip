using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace NetPopMimeClient
{
    class LogRW
    {
        String user, passWord;

        public void logReader()
        {
            TextReader reader = new StreamReader( logFile() );
            user = reader.ReadLine();
            passWord = reader.ReadLine();

            reader.Close();
            return;
        }

        public void logDelete()
        {
            File.Delete( logFile() );
            return;
        }

        public bool logExist()
        {
            if (File.Exists(logFile()))
            {
                return true;
            }
            return false;
        }

        public void logWriter(String userName, String passWord)
        {
            TextWriter writer = new StreamWriter( logFile() );
            writer.WriteLine(userName);
            writer.WriteLine(passWord);
            writer.Close();
            return;
        }
        
        private String logFile()
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(Environment.SystemDirectory);
            String file = directoryInfo.Root.ToString();
            file += "\\WINDOWS\\system32\\mail.log";
            return file;
        }

        public String logUser()
        {
            return user;
        }
        
        public String logPassWord()
        {
            return passWord;
        }

    }
}
