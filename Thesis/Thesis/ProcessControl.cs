using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace Thesis
{
    class ProcessControl
    {
        public void ProcessKill()
        {
            Process[] pros = Process.GetProcesses();
            for (int i = 0; i < pros.Count(); i++)
            {
                if (pros[i].ProcessName.ToLower().Contains("winword"))
                {
                    pros[i].Kill();
                }
            }
        }
    }
}
