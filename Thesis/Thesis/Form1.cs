using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ManualFolderBrowser;
using ThesisMain;
using SpeechBuilder;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Accessibility;
using System.Management;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Threading;
using System.Reflection;
using Microsoft.CSharp;
using System.CodeDom.Compiler;

namespace Thesis
{
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00020400-0000-0000-C000-000000000046")]
    public interface IDispatch
    {
    }
    public partial class Form1 : Form
    {
        static uint x = 0;
        static uint NewX = 0;
        Boolean ff = false;
        Boolean obSelection = false;
        int hwnd = 0;
        int hWndChild = 0;      
        private int occupiedBuffer = 0;
        private int excelHookCount = 0;
        private int wordHookCount = 0;
        private int pPointHookCount = 0;
        private String fileEx = null;
        private int MAX_TITLE = 256;
        Thread thread = null;
        int c = 0;
        private static Word.Application word = null;
        private Excel.Application excel = null;
        private PowerPoint.Application ppt = null;
        public ManagementEventWatcher mgmtWtch;  

        private enum SystemEvents : uint
        {
            EVENT_MIN = 0x00000001,
            EVENT_MAX = 0x7FFFFFFF,
            EVENT_CONSOLE_CARET = 0x4001,
            EVENT_OBJECT_DESCRIPTIONCHANGE = 0x800D,
            EVENT_OBJECT_ACCELERATORCHANGE = 0x8012,
            EVENT_OBJECT_STATECHANGE = 0x800A,
            EVENT_SYSTEM_SOUND = 0x0001,
            EVENT_SYSTEM_DESTROY = 0x8001,
            EVENT_SYSTEM_DRAGDROPSTART = 0x000E,
            EVENT_SYSTEM_MINIMIZESTART = 0x0016,
            EVENT_SYSTEM_MINIMIZEEND = 0x0017,
            EVENT_SYSTEM_FOREGROUND = 0x0003,
            EVENT_SYSTEM_MENUSTART = 0x0004,
            EVENT_OBJECT_FOCUS = 0x8005,
            EVENT_OBJECT_SELECTION = 0x8006,
            EVENT_OBJECT_NAMECHANGE = 0x800c,
            EVENT_OBJECT_SELECTIONADD = 0x8007,
            EVENT_OBJECT_CREATE = 0x8000,
            EVENT_OBJECT_VALUECHANGE = 0x800E,
            EVENT_OBJECT_TEXTSELECTIONCHANGED = 0x8014,
            EVENT_OBJECT_PARENTCHANGE = 0x800F,
            EVENT_OBJECT_END = 0x80FF
        }
        public const uint WINEVENT_OUTOFCONTEXT = 0x0000;
        public const uint WINEVENT_SKIPOWNTHREAD = 0x0001;
        public const uint WINEVENT_SKIPOWNPROCESS = 0x0002;
        public const uint WINEVENT_INCONTEXT = 0x0004;

        Guid IID_IDispatch;
        private int oldMessageFilter;

        #region APIs
        public delegate bool EnumChildCallback(int hwnd, ref int lParam);

        [DllImport("ole32.dll")]
        static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string
           lpszProgID, out Guid pclsid);

        [DllImport("ole32.dll", ExactSpelling = true)]
        public static extern int CoRegisterMessageFilter(int newFilter, ref int oldMsgFilter);

        [DllImport("User32")]
        public static extern bool EnumChildWindows(
            int hWndParent, EnumChildCallback lpEnumFunc, ref int lParam);

        [DllImport("User32")]
        public static extern int FindWindowEx(
            int hwndParent, int hwndChildAfter, string lpszClass,
            int missing);

        // AccessibleObjectFromWindow gets the IDispatch pointer of an object
        // that supports IAccessible, which allows us to get to the native OM.       
        //[DllImport("Oleacc.dll")]
        //private static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, ref PowerPoint.DocumentWindow ptr);

        [DllImport("Oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            int hwnd, uint dwObjectID,
            byte[] riid,
             ref IntPtr exW);

        [DllImport("User32")]
        public static extern int GetClassName(
            int hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern void SHChangeNotify(UInt32 wEventId, UInt32 uFlags, IntPtr dwItem1, IntPtr dwItem2);

        static Accessibility.IAccessible iAccessible;//interface: Accessibility namespace
        static object ChildId;

        [DllImport("user32.dll")]
        private static extern uint RealGetWindowClass(IntPtr hWnd, StringBuilder pszType, uint cchType);

        [DllImport("oleacc.dll")]
        public static extern uint WindowFromAccessibleObject(IAccessible pacc, ref IntPtr phwnd);

        [DllImport("oleacc.dll")]
        private static extern IntPtr AccessibleObjectFromEvent(IntPtr hwnd, uint dwObjectID, uint dwChildID,
            out IAccessible ppacc, [MarshalAs(UnmanagedType.Struct)] out object pvarChild);
        [DllImport("oleacc.dll")]
        public static extern uint AccessibleChildren(IAccessible paccContainer, int iChildStart, int cChildren, [Out] object[] rgvarChildren, out int pcObtained);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out Point pt);

        [DllImport("user32.dll", EntryPoint = "WindowFromPoint", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr WindowFromPoint(Point pt);

        [DllImport("user32.dll", EntryPoint = "SendMessageW")]
        public static extern int SendMessageW([InAttribute] System.IntPtr hWnd, int Msg, int wParam, IntPtr lParam);
        public const int WM_GETTEXT = 13;

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        internal static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        internal static extern IntPtr GetFocus();


        [DllImport("user32.dll")]
        static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr
        hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess,
        uint idThread, uint dwFlags);
        delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, uint idObject, uint idChild, uint dwEventThread, uint dwmsEventTime);
        #endregion

        private WinEventDelegate dEvent;
        private IntPtr pHook;
        private Browser browser;
        public static String URL;
        private Boolean flag;
        private SpeechControl speaker;
        Win32API Win32API = new Win32API();
        public Form1(SpeechControl speaker)
        {
            /*
            ProcessModule objCurrentModule = Process.GetCurrentProcess().MainModule;
            objKeyboardProcess = new LowLevelKeyboardProc(captureKey);
            ptrHook = SetWindowsHookEx(13, objKeyboardProcess, GetModuleHandle(objCurrentModule.ModuleName), 0);
             */

            InitializeComponent();
            this.speaker = speaker;
            browser = new Browser(listView, speaker, urlBox);
            new ThesisMain.Form1(speaker);
            flag = true;            
            //this.Opacity = 0.0;
            dEvent = this.WinEvent;
            pHook = SetWinEventHook(
                    (uint)SystemEvents.EVENT_MIN,
                    (uint)SystemEvents.EVENT_MAX,
                    IntPtr.Zero,
                    dEvent,
                    (uint)0,
                    (uint)0,
                    WINEVENT_OUTOFCONTEXT                    
                    );
        }
        public static IntPtr GetControlHandlerFromEvent(IntPtr hWnd, uint idObject, uint idChild)
        {
            //IntPtr hwnd = GetFocusedWindow();
            IntPtr handler = IntPtr.Zero;
            //IAccessible accWindow = null;
            //object objChild;
            handler = AccessibleObjectFromEvent(hWnd, idObject, idChild, out iAccessible, out ChildId);
            WindowFromAccessibleObject(iAccessible, ref handler);
            return handler;
        }
        public string Gettext()
        {
            try
            {
                if (iAccessible != null && ChildId != null)
                {
                    return iAccessible.get_accName(ChildId);
                }
                else return " ";
            }
            catch (Exception ex)
            {
                return null;
            }

        }
        private string GetWindowClass(IntPtr Hwnd)
        {
            // This function gets the name of a window class from a window handle
            StringBuilder Title = new StringBuilder(256);
            RealGetWindowClass(Hwnd, Title, 256);
            return Title.ToString().Trim();
        }
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
        private void Speak()
        {
            String getText = Gettext();
            speaker.speak(getText);
        }
        private void WinEvent(IntPtr hWinEventHook, uint eventType, IntPtr hWnd, uint idObject, uint idChild, uint dwEventThread, uint dwmsEventTime)
        {
            //MessageBox.Show("fdkjfkdj");
            //i++;
            //Console.WriteLine("Object"+i);               
            if (eventType == (uint)SystemEvents.EVENT_OBJECT_SELECTION)
            {
                uint processIID = 0;
                uint ID = GetWindowThreadProcessId(GetForegroundWindow(), out processIID);
                if (Process.GetProcessById((int)processIID).ProcessName.ToLower() == "explorer") { obSelection=true; return; }
                NewX = idObject;
                if (x == NewX && idChild == 0) return;
                x = NewX;
                GetControlHandlerFromEvent(hWnd, idObject, idChild);
                if (thread != null) { thread.Abort(); thread = null; }
                thread = new Thread(new ThreadStart(Speak));
                thread.Start();
            }
            else if (eventType == (uint)SystemEvents.EVENT_OBJECT_FOCUS)
            {
                NewX = idObject;
                if (x == NewX && idChild == 0) return;
                x = NewX;
                //Console.WriteLine("On focus change event");
                GetControlHandlerFromEvent(hWnd, idObject, idChild);
                if (thread != null) { thread.Abort(); thread = null; }
                thread = new Thread(new ThreadStart(Speak));
                thread.Start();
                obSelection = true;
            }            
            else if (eventType == (uint)SystemEvents.EVENT_SYSTEM_FOREGROUND)
            {
                GetControlHandlerFromEvent(hWnd, idObject, idChild);
                if (thread != null) { thread.Abort(); thread = null; }
                thread = new Thread(new ThreadStart(Speak));
                thread.Start();
                //Console.WriteLine("On Foreground change event");
                if (GetActiveProcess().ToString().Equals("EXCEL"))
                {
                    //Console.WriteLine("hookCount " + excelHookCount);
                    //System.Threading.Thread.Sleep(3000);                  
                    int i = 0;
                    //Console.WriteLine(excel.Name);
                    if (excelHookCount == 0 && !ff)
                    {
                        ff = true;
                        excel = GetAccessibleObjectFromMarshal();
                        ExcellForm.Form1 ExF1 = new ExcellForm.Form1(excel, speaker);
                        ff = false;
                    }
                    excelHookCount++;
                    //Console.WriteLine("ffffffffffffffffffffff");
                }
                else if (GetActiveProcess().ToString().Equals("WINWORD"))
                {
                    if (wordHookCount == 0 && !ff)
                    {
                        while (true)
                        {
                            try
                            {
                                ff = true;
                                word = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                                if (word != null) { ff = false; break; }
                                Console.WriteLine("not found word object");
                                //ff = false;
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show(ex.ToString()); 
                            }
                        }
                        DocForm.DocForm docForm = new DocForm.DocForm(word, speaker);
                    }
                    wordHookCount++;

                }
                else if (GetActiveProcess().ToString().Equals("POWERPNT"))
                {
                    //System.Threading.Thread.Sleep(3000);
                    Console.WriteLine("hookCount ffffffffff  " + pPointHookCount);
                    if (pPointHookCount == 0 && !ff)
                    {
                        ff = true;
                        ppt = (PowerPoint.Application)Microsoft.VisualBasic.Interaction.GetObject("", "PowerPoint.Application");                        
                        Console.WriteLine(ppt.Name);
                        PPTForm.Form1 PpF1 = new PPTForm.Form1(ppt, speaker);
                        ff = false;
                    }
                    pPointHookCount++;
                    Console.WriteLine("hookCount " + pPointHookCount);
                }
            }
            else if (eventType == (uint)SystemEvents.EVENT_OBJECT_NAMECHANGE)
            {
                //x = hWnd;
                NewX = idObject;
                if (x == NewX && idChild == 0) return;
                x = NewX;
                uint processID = 0;
                uint processID1 = 0;
                uint id = GetWindowThreadProcessId(GetForegroundWindow(), out processID);
                //Console.WriteLine("process "+processID.ToString());
                uint id1 = GetWindowThreadProcessId(hWnd, out processID1);
                if (Process.GetProcessById((int)processID).ProcessName.ToLower() != "explorer") return;
                //if (obSelection) {return; }
                if (processID == processID1)
                {
                    GetControlHandlerFromEvent(hWnd, idObject, idChild);
                    Console.WriteLine("hwnd " + hWnd + " Name Change: " + Gettext());
                    if (thread != null) { thread.Abort(); thread = null; }
                    speaker.stop();
                    thread = new Thread(new ThreadStart(Speak));
                    thread.Start();
                }
            }
        }
        private void WaitForProcess()
        {
            try
            {
                WqlEventQuery query1 = new WqlEventQuery("__InstanceCreationEvent", new TimeSpan(0, 0, 3), "TargetInstance isa \"Win32_Process\"");
                WqlEventQuery query2 = new WqlEventQuery("__InstancedeletionEvent", new TimeSpan(0, 0, 3), "TargetInstance isa \"Win32_Process\"");           
                ManagementEventWatcher startWatch1 = new ManagementEventWatcher(query1);
                ManagementEventWatcher startWatch2 = new ManagementEventWatcher(query2);
                startWatch1.EventArrived
                                    += new EventArrivedEventHandler(startWatch_EventArrived);
                startWatch2.EventArrived
                                    += new EventArrivedEventHandler(endWatch_EventArrived);           
                startWatch1.Start();
                startWatch2.Start();

            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            
        }
        private bool EnumChildProc(int hwnd, ref int lParam)
        {
            StringBuilder windowClass = new StringBuilder(128);
            GetClassName(hwnd, windowClass, 128);
            if (windowClass.ToString() == "paneClassDC")
            {
                lParam = hwnd;
            }
            else if (windowClass.ToString() == "_WwG")
            {
                lParam = hwnd;
            }
            else if (windowClass.ToString() == "EXCEL7")
            {
                lParam = hwnd;
            }
            return true;
        }
        private Excel.Application GetAccessibleObjectFromMarshal()
        {
            try
            {
                //Console.WriteLine("excel.Name=");
                while (true)
                {
                    hwnd = FindWindowEx(0, 0, "XLMain", 0);
                    if (hwnd != 0) break;
                }
                //if (hwnd == 0) GetAccessibleObject();
                if (hwnd != 0)
                {
                    // Walk the children of this window to see if any are
                    // IAccessible.

                    EnumChildCallback cb =
                        new EnumChildCallback(EnumChildProc);
                    while (true)
                    {
                        EnumChildWindows(hwnd, cb, ref hWndChild);
                        if (hWndChild != 0) break;
                        //Console.WriteLine("Enum Child");
                    }
                    //Console.WriteLine(hWndChild.ToString());
                    if (hWndChild != 0)
                    {
                        // OBJID_NATIVEOM gets us a pointer to the native 
                        // object model.
                        uint OBJID_NATIVEOM = 0xFFFFFFF0;
                        //uint OBJID_NATIVEOM = 0x00000000;
                        //Guid IID_IDispatch = new Guid("{618736e0-3c3d-11cf-810c-00aa00389b71}");
                        Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
                        //PowerPoint.DocumentWindow ptr = null;
                        IntPtr ptr = IntPtr.Zero;
                        int i = 0;
                        int hr = -1;
                        while (true)
                        {
                            i++;
                            //if (i >= 200) break;
                            hr = AccessibleObjectFromWindow(
                                hWndChild, OBJID_NATIVEOM,
                                IID_IDispatch.ToByteArray(), ref ptr);
                            //Console.WriteLine(hr);
                            if (hr >= 0) break;
                            else if (i >= 300) GetAccessibleObjectFromMarshal();
                        }
                        if (hr >= 0)
                        {
                            int j = 0;
                            while (true)
                            {
                                j++;
                                //Console.WriteLine("j=" + j);
                                Excel.Window ew = (Excel.Window)Marshal.GetObjectForIUnknown(ptr);
                                excel = ew.Application;
                                //Console.WriteLine(excel.Name);
                                if (excel != null) break;
                            }
                            //MessageBox.Show(ppt.Name);
                            //Console.WriteLine(ppt.Name);
                            return excel;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
               // MessageBox.Show("on GetAccessibleObjectFromMarshal= " + ex.ToString());
            }
            return null;
        }
        private void GetAccessibleObject()
        {
            //MessageBox.Show("fffffffff");
            try
            {

                //int hwnd = 0;
                // Walk the children of the desktop to find PowerPoint’s main
                // window.
                //System.Threading.Thread.Sleep(2000);
                //hwnd = FindWindowEx(0, 0, "PP12FrameClass", 0);
                Console.WriteLine("Enter");
                if (fileEx == "WINWORD.EXE")
                {
                    //word = new Microsoft.Office.Interop.Word.Application();
                    while (true)
                    {
                        hwnd = FindWindowEx(0, 0, "OpusApp", 0);
                        if (hwnd != 0) break;
                        Console.WriteLine(hwnd);
                    }
                    CLSIDFromProgID("Word.Application", out IID_IDispatch);
                    //MessageBox.Show(hwnd.ToString());
                }
                else if (fileEx == "POWERPNT.EXE")
                {
                    while (true)
                    {

                        hwnd = FindWindowEx(0, 0, "PP12FrameClass", 0);
                        if (hwnd != 0) break;
                        //Console.WriteLine(i);
                    }
                    CLSIDFromProgID("PowerPoint.Application", out IID_IDispatch);
                    //MessageBox.Show(hwnd.ToString());
                }
                else if (fileEx == "EXCEL.EXE")
                {
                    while (true)
                    {
                        hwnd = FindWindowEx(0, 0, "XLMain", 0);
                        if (hwnd != 0) break;
                    }
                    CLSIDFromProgID("Excel.Application", out IID_IDispatch);
                }
                //if (hwnd == 0) GetAccessibleObject();
                if (hwnd != 0)
                {
                    // Walk the children of this window to see if any are
                    // IAccessible.

                    EnumChildCallback cb =
                        new EnumChildCallback(EnumChildProc);
                    while (true)
                    {
                        EnumChildWindows(hwnd, cb, ref hWndChild);
                        if (hWndChild != 0) break;
                        Console.WriteLine("Enum Child");
                    }
                    //Console.WriteLine(hWndChild.ToString());
                    if (hWndChild != 0)
                    {
                        // OBJID_NATIVEOM gets us a pointer to the native 
                        // object model.                           
                        uint OBJID_NATIVEOM = 0xFFFFFFF0;
                        //uint OBJID_NATIVEOM = 0x00000000;
                        //
                        //MessageBox.Show(IID_IDispatch.ToString());
                        //IID_IDispatch = new Guid("{00000000-0000-0000-0000-000000000000}");
                        IID_IDispatch = new Guid("{00020962-0000-0000-C000-000000000046}");
                        //Guid IID_IDispatch = new Guid("{00000016-0000-0000-C000-000000000046}");
                        //PowerPoint.DocumentWindow ptr = null;
                        IntPtr ptr = IntPtr.Zero;
                        int hr = -1;
                        int i = 0;
                        while (true)
                        {
                            i++;
                            //if (i >= 200) break;
                            hr = AccessibleObjectFromWindow(
                                hWndChild, OBJID_NATIVEOM,
                                null, ref ptr);
                            Console.WriteLine(hr);
                            if (hr >= 0) break;
                            //else if (i >= 1300) GetAccessibleObject();
                        }
                        //Console.WriteLine(hr);
                        Console.WriteLine(ptr);
                        if (hr >= 0)
                        {
                            Console.WriteLine("fffffffffffffffffffffffff");
                            if (fileEx == "POWERPNT.EXE")
                            {
                                while (true)
                                {
                                    Console.WriteLine("Get Marshal");
                                    PowerPoint.DocumentWindow pdw = (PowerPoint.DocumentWindow)Marshal.GetObjectForIUnknown(ptr);
                                    ppt = pdw.Application;
                                    if (ppt != null) break;
                                }
                                Console.WriteLine(ppt.Name);
                                //MessageBox.Show(ppt.Name);
                            }
                            else if (fileEx == "WINWORD.EXE")
                            {
                                while (true)
                                {
                                    Word.Window ww = (Word.Window)Marshal.GetObjectForIUnknown(ptr);
                                    word = ww.Application;
                                    if (word != null) break;
                                }
                                //MessageBox.Show(word.Name);
                            }
                            else if (fileEx == "EXCEL.EXE")
                            {
                                //excel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                                int j = 0;
                                while (true)
                                {
                                    j++;
                                    Console.WriteLine("j=" + j);
                                    Excel.Window ew = (Excel.Window)Marshal.GetObjectForIUnknown(ptr);
                                    excel = ew.Application;
                                    //Console.WriteLine(excel.Name);
                                    if (excel != null) break;

                                }
                                //MessageBox.Show(excel.Name);
                            }
                            //MessageBox.Show(ppt.Name);
                            //Console.WriteLine(ppt.Name);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("on GetAccessibleObject= " + ex.ToString());
            }

        }       
        private void startWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {            
            ManagementBaseObject bobj = ((ManagementBaseObject)e.NewEvent["TargetInstance"]);
            //MessageBox.Show(bobj["Name"].ToString());//return process name ex:notepad.exe          
            String processName = bobj["Name"].ToString();
            Console.WriteLine(processName);           
            try
            {
                //System.Threading.Thread.Sleep(6000);
                if (processName == "WINWORD.EXE")
                {
                    fileEx = processName;
                    //wordHookCount = 0;                    
                }
                else if (processName == "EXCEL.EXE")
                {
                    fileEx = processName;
                    //excelHookCount = 0;
                    //hwnd = int.Parse(bobj["Handle"].ToString());
                    //int i = 0;

                    //thread = new Thread(new ThreadStart(GetAccessibleObject));
                    //thread.Start();
                    //GetAccessibleObject();
                    //GetAccessibleObjectFromMarshal();
                    //i = 0;
                    //while (true)
                    //{
                    //    if (i >= 250) break;
                    //    i++;
                    //    Console.WriteLine(i);
                    //}
                    //MessageBox.Show(hWndChild.ToString());
                    //Console.WriteLine(excel.Name);
                    //ExcellForm.Form1 ExF1 = new ExcellForm.Form1(excel,speaker);
                    //ExcellForm.Form1 ExF1 = new ExcellForm.Form1();                    
                    //try
                    //{
                    //    Form2 f2 = new Form2();
                    //    f2.Show();
                    //    f2.Dispose();
                    //    System.Threading.Thread.Sleep(5000);
                    //    excel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                    //}
                    //catch (Exception ex)
                    //{
                    //    //excel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    //    MessageBox.Show("excel problem");
                    //}
                    //GetAccessibleObject(fileEx);
                    //ExcellForm.Form1 ExF1 = new ExcellForm.Form1(excel, speaker);

                }
                else if (processName == "POWERPNT.EXE")
                {
                    fileEx = processName;
                    //pPointHookCount = 0;
                    //Console.WriteLine(processName);
                    //System.Threading.Thread.Sleep(20000);
                    //ppt = (PowerPoint.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
                    //thread = new Thread(new ThreadStart(GetAccessibleObject));
                    //thread.Start();
                    //GetAccessibleObject();
                    //GetAccessibleObjectFromMarshal();
                    //System.Threading.Thread.Sleep(2000);
                    //int i = 0;
                    //while (true)
                    //{
                    //    if (i >= 300) break;
                    //    i++;
                    //    Console.WriteLine(i);
                    //}
                    //Console.WriteLine(ppt.Name);
                    //MessageBox.Show(ppt.Name);
                    //PPTForm.Form1 PpF1 = new PPTForm.Form1(ppt, speaker);
                }
                ////Console.WriteLine(c);

            }
            catch (Exception ex)
            {
                MessageBox.Show("on Event arrived" + ex.ToString());
                //excel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

            }
        }
        private void endWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject mBaseObj = ((ManagementBaseObject)e.NewEvent["TargetInstance"]);
            //MessageBox.Show(bobj["Name"].ToString());//return process name ex:notepad.exe          
            String endProcessName = mBaseObj["Name"].ToString();
            Console.WriteLine("endProcessName"+endProcessName);
            if (endProcessName == "WINWORD.EXE")
            {
                wordHookCount = 0;
            }
            else if (endProcessName == "EXCEL.EXE")
            {
                excelHookCount = 0;
            }
            else if (endProcessName == "POWERPNT.EXE")
            {
                pPointHookCount = 0;
            }                 
        }
        /*
        private IntPtr captureKey(int nCode, IntPtr wp, IntPtr lp)
        {
            if (nCode >= 0)
            {
                KBDLLHOOKSTRUCT objKeyInfo = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lp, typeof(KBDLLHOOKSTRUCT));

                if (objKeyInfo.key == Keys.RWin || objKeyInfo.key == Keys.LWin || objKeyInfo.key == Keys.F1 || objKeyInfo.key == Keys.F6) // Disabling Windows keys
                {
                    return (IntPtr)1;
                }
            }
            return CallNextHookEx(ptrHook, nCode, wp, lp);
        }
        // End 5/12/2011 :: Disable Start menu
         */
        //public Form1(SpeechControl speaker)
        //{
        //    InitializeComponent();

        //    browser = new Browser(listView, speaker, urlBox);
        //    new ThesisMain.Form1(speaker);

        //    this.speaker = new SpeechControl();

        //    flag = true;
        //    this.Opacity = 0.0;

        //}

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            browser.setDrives();
        }

        private void linkLabel1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode.ToString().Equals("Right"))
            {
                listView.Select();
                listView.Focus();
                if (urlBox.Text == "My Computer")
                    browser.setDrives();
                else
                    browser.settingBrowser();
            }
        }


        private void linkLabel1_Enter(object sender, EventArgs e)
        {
            if (!flag)
            {
                listView.Select();
                flag = true;
            }

        }
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            DirectoryInfo directoryInfo = new DirectoryInfo(Environment.SystemDirectory);
            DirectoryInfo dirInfo = new DirectoryInfo(directoryInfo.Root.ToString() + "Users\\" + Environment.UserName + "\\Desktop\\");
            if (dirInfo.Exists)
                browser.setPath(directoryInfo.Root.ToString() + "Users\\" + Environment.UserName + "\\Desktop\\");
            else
                browser.setPath(directoryInfo.Root.ToString() + "Documents and Settings\\" + Environment.UserName + "\\Desktop\\");
            browser.settingBrowser();
        }

        private void linkLabel2_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode.ToString().Equals("Right"))
            {
                listView.Select();
                listView.Focus();
                if (urlBox.Text == "My Computer")
                    browser.setDrives();
                else
                    browser.settingBrowser();
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(Environment.SystemDirectory);
            DirectoryInfo dirInfo = new DirectoryInfo(directoryInfo.Root.ToString() + "Users\\" + Environment.UserName + "\\Documents\\");
            if (dirInfo.Exists)
                browser.setPath(directoryInfo.Root.ToString() + "Users\\" + Environment.UserName + "\\Documents\\");
            else
                browser.setPath(directoryInfo.Root.ToString() + "Documents and Settings\\" + Environment.UserName + "\\My Documents\\");
            browser.settingBrowser();
        }
        private void linkLabel3_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode.ToString().Equals("Right"))
            {
                listView.Select();
                listView.Focus();
                if (urlBox.Text == "My Computer")
                    browser.setDrives();
                else
                    browser.settingBrowser();
            }
        }

        private void linkLabel4_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode.ToString().Equals("Right"))
            {
                listView.Select();
                listView.Focus();
                if (urlBox.Text == "My Computer")
                    browser.setDrives();
                else
                    browser.settingBrowser();
            }
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(Environment.SystemDirectory);

            DirectoryInfo dirInfo = new DirectoryInfo(directoryInfo.Root.ToString() + "Users\\" + Environment.UserName + "\\Music\\");
            if (dirInfo.Exists)
                browser.setPath(directoryInfo.Root.ToString() + "Users\\" + Environment.UserName + "\\Music\\");
            else
                browser.setPath(directoryInfo.Root.ToString() + "Documents and Settings\\" + Environment.UserName + "\\My Documents" + "\\My Music\\");
            browser.settingBrowser();
        }

        private void dateTimer_Tick(object sender, EventArgs e)
        {
            if (this.Opacity <= 1)
                this.Opacity += 0.3;

            DateTime currTime = DateTime.Now;
            day_label.Text = currTime.Day + "/" + currTime.Month + "/" + currTime.Year;
            time_label.Text = currTime.TimeOfDay.Hours + ":" + currTime.TimeOfDay.Minutes + ":" + currTime.TimeOfDay.Seconds;

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            //// 2/11/2011:: When Mongol Dip(MD) Exit make Subachan exit automatically 
            //Process[] pros = Process.GetProcesses();
            //for (int i = 0; i < pros.Count(); i++)
            //{
            //    if (pros[i].ProcessName.ToLower().Contains("javaw"))
            //    {
            //        pros[i].Kill();
            //    }
            //}
            //// End 2/11/2011:: When Mongol Dip(MD) Exit make Subachan exit automatically
            Application.Exit();

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void listView_SelectedIndexChanged(object sender, EventArgs e)
        {
            //String op;
            //op = e.KeyCode.ToString().Equals("+");
            //Console.WriteLine(op);
        }

        private void urlBox_TextChanged(object sender, EventArgs e)
        {
            URL = urlBox.Text;
            new ThesisMain.Form1(URL);
            //MessageBox.Show(URL);
        }

        private void listView_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected)
            {
                new Browser(listView.Items[e.ItemIndex].Text);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            WaitForProcess(); 
        }        
        private void Form1_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            Process[] pros = Process.GetProcesses();
            for (int i = 0; i < pros.Count(); i++)
            {
                if (pros[i].ProcessName.ToLower().Contains("javaw"))
                {
                    pros[i].Kill();
                }
            }
        }
    }
}
