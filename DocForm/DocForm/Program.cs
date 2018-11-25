using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DocForm
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //DocForm form = new DocForm("C:\\Documents and Settings\\RASHED\\Desktop\\Test.doc");
            Application.Run();
        }
    }
}
