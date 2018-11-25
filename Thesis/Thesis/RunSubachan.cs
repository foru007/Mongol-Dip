using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace Thesis
{
    public class RunSubachan
    {
        //2/11/2011 :: Make Subachan runable from c#  code
        public void Rsubachan()
        {
            System.Diagnostics.Process proc = null;
            try
            {
                string targetDir = string.Format(@"C:\Resources\Subachan");//this is directory
                proc = new System.Diagnostics.Process();
                proc.StartInfo.WorkingDirectory = targetDir;
                proc.StartInfo.FileName = "Subachan.jar";
                proc.StartInfo.Arguments = string.Format("10");//this is argument
                proc.StartInfo.CreateNoWindow = false;
                proc.Start();                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            }
            //End::
        }
    }
}
