using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using SpeechBuilder;
using System.Threading;
using StartLoader;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace Thesis
{
    static class Program
    {        
        private static SpeechControl speaker;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //try
            //{
            //    speaker = new SpeechControl();
            //    ReadyToStart ready = new ReadyToStart();
            //    Thread thread = new Thread( new ThreadStart( ready.start ));
            //    thread.Start();
            //    Thread.Sleep( 10000 );
            //    //ready.start();
            //    Application.Run(new Form1(speaker));
            //}
            //catch( Exception )
            //{
            //}            
            speaker = new SpeechControl();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            // 2/11/2011:: Make Subachan runable from c#  code
            //ProcessControl pr = new ProcessControl();
            //pr.ProcessKill();
            RunSubachan rs = new RunSubachan();
            rs.Rsubachan();
            FolderOptionControl folderOption = new FolderOptionControl();
            folderOption.ShowFileExtension();
            //System.Threading.Thread.Sleep(3000);
            //End::
            //Application.Run(new LoaderForm());
            Application.Run(new Form1( speaker ));            

        }
    }
}
