using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Diagnostics;

namespace ThesisMain
{
    class Communication
    {
        Win32API Win32API = new Win32API();

        public String getWindowName()
        {
            const int nChars = 1000;
            int handle = 0;
            StringBuilder Buff = new StringBuilder(nChars);

            handle = (int)Win32API.GetForegroundWindow();
            if (Win32API.GetWindowText((IntPtr)handle, Buff, nChars) > 0)
            {
                return Buff.ToString();
            }
            return null;
        }

        public StringBuilder getSelecedText()
        {
            const int WM_GETTEXT = 13;                                      
            StringBuilder builder = new StringBuilder(500);

            int foregroundWindowHandle = Win32API.GetForegroundWindow();

            uint remoteThreadId = Win32API.GetWindowThreadProcessId(foregroundWindowHandle, 0);
            uint currentThreadId = Win32API.GetCurrentThreadId();
            Win32API.AttachThreadInput(remoteThreadId, currentThreadId, true);
            int focused = Win32API.GetFocus();

            Win32API.SendMessage(focused, WM_GETTEXT, builder.Capacity, builder);
            return builder;

        }

        /// <summary>
        /// Retrieves name of active Process.
        /// </summary>
        /// <returns>Active Process Name</returns>
        public string GetActiveProcess()
        {
            const int nChars = 256;
            int handle = 0;
            StringBuilder Buff = new StringBuilder(nChars);
            handle = (int)Win32API.GetForegroundWindow();

            // If Active window has some title info
            if (Win32API.GetWindowText(handle, Buff, nChars) > 0)
            {
                uint lpdwProcessId;
                uint dwCaretID = Win32API.GetWindowThreadProcessId(handle, out lpdwProcessId);
                uint dwCurrentID = (uint)Thread.CurrentThread.ManagedThreadId;
                return Process.GetProcessById((int)lpdwProcessId).ProcessName;

            }
            // Otherwise either error or non client region
            return String.Empty;
        }

        public bool isClientRegion()
        {
            if (GetActiveProcess() == "" || GetActiveProcess().ToLower() == "explorer")
            {
                return false;
            }
            return true;
        }
    }
}
