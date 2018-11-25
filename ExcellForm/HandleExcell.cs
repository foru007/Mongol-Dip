using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using SpeechBuilder;
using Gma.UserActivityMonitor;
using System.Windows.Forms;
using System.IO;
using EPocalipse.IFilter;
using ThesisMain;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcellForm
{
    public class HandleExcell
    {
        private Excel.Application excel = null;
        private Excel._Workbook book = null;
        private Excel._Worksheet sheet = null;

        private int ED_mode = 0;
        private int track = 0;

        private ExcelReadThread bookThread;
        private Communication communication;
        private SpeechControl speaker;

        private object fileName = null;
        private object readOnly = false;
        private TextReader reader;
        private object missing = Type.Missing;
        object isVisible = true;

        private Boolean pressCtrl = false;
        private Boolean pressAlt = false;
        private Boolean shift = false;
        private Boolean pressF2 = false;

        String pp = null;
        Process proc = new Process();

        private int x, y;

        private String fName; //string
        private String fullText;
        //  Thread thread = null;
        Thread thread1 = null;

        public static int count = 0;
        private int p = 0;
        private static int d = 0;

        private static int caps = 0, insert = 0;

        private int kj = 0;

        Form2 ob = new Form2();

        private Boolean insrt = false;

        ////private static int ctr_alt_pgup = 0;

        public HandleExcell(String fileName, SpeechControl speaker, Excel.Application excel, int p)
        {
            this.p = p;
            this.excel = excel;
            this.fileName = fileName;
            this.speaker = speaker;
            this.fName = fileName;
            fullText = null;

            //HookManager.KeyDown += HookManager_KeyDown;
            //HookManager.KeyUp -= HookManager_KeyUp;
            //HookManager.KeyDown -= HookManager_KeyDown;
            startHandle();
        }
        public HandleExcell(SpeechControl speaker, Excel.Application excel, int p1)
        {
            this.p = p1;
            this.excel = excel;
            this.speaker = speaker;
            //ob = new Form2();
            //startHandleWithoutFileName();          
        }
        public HandleExcell(SpeechControl speaker, Excel.Application excel)
        {
            this.excel = excel;
            this.speaker = speaker;
            //ob = new Form2();
            //startHandleWithoutFileName();
            NewstartHandle();
        }
        public HandleExcell()
        {
            //Console.WriteLine("On handle Excel");
            NewstartHandle();
        }
        ~HandleExcell()
        {
            HookManager.KeyUp -= HookManager_KeyUp;
            HookManager.KeyDown -= HookManager_KeyDown;
        }
        public void startHandleWithoutFileName()
        {
            if (p == 0) d = 0;
            d++;
            if (d == 1)
            {
                //HookManager.KeyUp += HookManager_KeyUp;
                //HookManager.KeyDown += HookManager_KeyDown;
            }
            try
            {
                excel.Visible = true;
                book = excel.Workbooks.Add(Type.Missing);
                book.Activate();
                excel.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(newExcel_WorkbookBeforeClose);

            }
            catch (COMException)
            {
                MessageBox.Show("Error accessing Word document.");
            }

            //textBoxLog.Text = doc.Content.Text;

            //word.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(oWord_DocumentBeforeClose);
            ////DocOpenTrack = 0;


        }
        public void startHandle()
        {
            if (p == 0) d = 0;
            d++;
            speakStartToEnd();
            if (d == 1)
            {
                HookManager.KeyUp += HookManager_KeyUp;
                HookManager.KeyDown += HookManager_KeyDown;
            }
            try
            {
                Process[] procss = Process.GetProcessesByName("POWERPNT");
                if (procss.Length != 0)
                {
                    foreach (Process procc in procss)
                        procc.Kill();
                }
                excel.Visible = true;
                book = excel.Workbooks.Open(fName, missing, readOnly, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                book.Activate();
                excel.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(excel_WorkbookBeforeClose);
            }
            catch (Exception ex)
            {

            }
        }
        public void startHandleForUnhook()
        {
            HookManager.KeyUp -= HookManager_KeyUp;
            HookManager.KeyDown -= HookManager_KeyDown;
            //excel.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(excel_WorkbookBeforeClose);
            //book = excel.Workbooks;
        }
        public void NewstartHandle()
        {
            try
            {
                HookManager.KeyUp += HookManager_KeyUp;
                HookManager.KeyDown += HookManager_KeyDown;
                //excel.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(excel_WorkbookBeforeClose);
                //book = excel.Workbooks;                
            }
            catch (Exception e)
            {
            }
        }
        public void newExcel_WorkbookBeforeClose(Excel.Workbook document, ref bool Cancel)
        {
            //HookManager.KeyUp -= HookManager_KeyUp;
            //HookManager.KeyDown -= HookManager_KeyDown;
            ////System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            //try
            //{
            //    Process[] procs = Process.GetProcessesByName("EXCEL");
            //    foreach (Process proc in procs)
            //        proc.Kill();
            //}
            //catch (Exception ex)
            //{
            //    //Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            //}
        }
        public void Close()
        {
            try
            {
                if (excel != null)
                {
                    excel.WorkbookBeforeClose -= new Excel.AppEvents_WorkbookBeforeCloseEventHandler(excel_WorkbookBeforeClose);
                    excel.Workbooks.Close();
                    excel.Quit();
                    //object delete = range;
                    //Marshal.ReleaseComObject(ref delete);
                    //delete = sheet;
                    Marshal.ReleaseComObject(sheet);
                    //delete = book;
                    Marshal.ReleaseComObject(book);
                    //delete = books;
                    //ReleaseComObject(ref delete);
                    //delete = excel;
                    Marshal.ReleaseComObject(excel);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch (Exception ex)
            { }
        }
        public void excel_WorkbookBeforeClose(Excel.Workbook document, ref bool Cancel)
        {
            //MessageBox.Show("from excel");
            //book.Save();
            //excel.WorkbookBeforeClose -= new Excel.AppEvents_WorkbookBeforeCloseEventHandler(excel_WorkbookBeforeClose);
            Close();
            HookManager.KeyUp -= HookManager_KeyUp;
            HookManager.KeyDown -= HookManager_KeyDown;
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            //try
            //{
            //    Process[] procs = Process.GetProcessesByName("EXCEL");
            //    foreach (Process proc in procs)
            //        proc.Kill();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            //}

        }
        public void speakStartToEnd()
        {
            //reader = new FilterReader(fName);
            //fullText = reader.ReadToEnd();
            //reader.Dispose();

        }
        public void full()
        {
            //speaker.speak(fullText);
        }
        //private void AlterControl()
        //{
        //    try
        //    {
        //        int length = 15;
        //        String s = communication.getWindowName().ToString();
        //        if (s.Length < length) length = s.Length;
        //        String sx = s.Substring(0, length);
        //        //MessageBox.Show(sx);            
        //        if (sx != "Microsoft Excel") pressAlt = true;
        //    }
        //    catch (Exception ex)
        //    { }

        //}
        public void NarratorRunOrNotCheck()
        {
            //// Process is Running or not
            //Process[] pname = Process.GetProcessesByName("nvda");


            //if (pname.Length == 0)
            //{
            //    try
            //    {
            //        string targetDir = string.Format(@"C:\Resources\NVDA");//this is directory
            //        proc.StartInfo.WorkingDirectory = targetDir;
            //        proc.StartInfo.FileName = "nvda.exe";
            //        proc.StartInfo.Arguments = string.Format("10");//this is argument
            //        proc.StartInfo.CreateNoWindow = false;
            //        proc.Start();

            //    }
            //    catch (Exception ex)
            //    {
            //        //Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            //    }
            //}

            //
        }

        public void NarratorStop()
        {
            //try
            //{
            //    Process[] procs = Process.GetProcessesByName("nvda");
            //    if (procs.Length != 0)
            //    {
            //        foreach (Process proc in procs)
            //            proc.Kill();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    //Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            //}
        }

        private void HookManager_KeyDown(object sender, KeyEventArgs e)
        {
            //Console.WriteLine(excel.Name);
            try
            {

                //excel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                communication = new Communication();
                Thread thread = null;
                //MessageBox.Show(e.KeyData.ToString());
                bookThread = new ExcelReadThread(speaker, excel, book, e.KeyData.ToString());
                Char code = (char)e.KeyCode;
                if (!communication.GetActiveProcess().ToString().Equals("EXCEL") && bookThread != null)
                {
                    kj = 0;
                    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    bookThread.stopAll();
                    pressAlt = false;
                }

                if (communication.GetActiveProcess().ToString().Equals("EXCEL"))
                {
                    Win32API winapi = new Win32API();
                    bool at = winapi.EvaluateCaretPosition();

                    //if (ctr_alt_pgup == 1)
                    //{
                    //    pressAlt = false;
                    //    ctr_alt_pgup = 0;
                    //}

                    if (kj == 0)
                    {
                        HookManager.KeyUp -= HookManager_KeyUp;
                        HookManager.KeyDown -= HookManager_KeyDown;
                        //speaker.speak("winword file testing");
                        HookManager.KeyUp += HookManager_KeyUp;
                        HookManager.KeyDown += HookManager_KeyDown;
                        kj = kj + 1;
                    }

                    if (e.KeyData.ToString().Equals("LControlKey") || e.KeyData.ToString().Equals("RControlKey"))
                    {
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        pressCtrl = true;
                    }
                    else if (e.KeyData.ToString().Equals("LMenu") || e.KeyData.ToString().Equals("RMenu"))
                    {
                        //releaseAlt = true;
                        //NarratorRunOrNotCheck();
                        pressAlt = true;
                    }
                    else if (e.KeyData.ToString().Equals("LShiftKey") || e.KeyData.ToString().Equals("RShiftKey"))
                    {
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        shift = true;
                    }
                    //else if (insrt && e.KeyCode.Equals(Keys.F2))  // Cell Left Text
                    //{
                    //    e.Handled = true;
                    //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //    if (thread != null) { thread.Abort(); thread = null; }
                    //    thread = new Thread(new ThreadStart(bookThread.GetLeftText));
                    //}
                    //else if (insrt && e.KeyCode.Equals(Keys.F3))  // Cell Right Text
                    //{
                    //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //    if (thread != null) { thread.Abort(); thread = null; }
                    //    thread = new Thread(new ThreadStart(bookThread.GetRightText));
                    //}

                    else if (e.KeyData.ToString().Equals("F2"))
                    {
                        //if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        //if (thread != null) { thread.Abort(); thread = null; }
                        //thread = new Thread(new ThreadStart(bookThread.FF2));
                        if (!at)
                        {
                            //Console.WriteLine("edit mode");
                            ED_mode = ED_mode + 1;
                            ED_mode = ED_mode % 2;
                            if (ED_mode == 0)
                            {
                                ob.unhook();
                                track = 0;
                                speaker.speak("Enter Mode");
                                pressF2 = false;
                                if (thread1 != null) { thread1.Abort(); thread1 = null; }
                                if (thread != null) { thread.Abort(); thread = null; }
                                thread = new Thread(new ThreadStart(bookThread.set_track_F2));
                            }
                            else if (ED_mode == 1)
                            {
                                NarratorStop();
                                track = 1;
                                speaker.speak("Edit Mode");
                                pressF2 = true;
                                if (thread1 != null) { thread1.Abort(); thread1 = null; }
                                if (thread != null) { thread.Abort(); thread = null; }
                                ob.hook();
                                thread = new Thread(new ThreadStart(bookThread.setF2_initial));

                            }
                        }

                    }

                    else if (e.KeyData.ToString().Equals("Right") || e.KeyData.ToString().Equals("Left"))
                    {
                        //if (!pressAlt)
                        //    AlterControl();
                        if (!pressF2 && !at)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.operateFull));
                        }

                        else if (pressF2 && !at)////////////////////
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.operateChar));

                        }

                        //else if (pressAlt)
                        //    NarratorRunOrNotCheck();


                    }
                    else if (e.KeyData.ToString().Equals("Up") || e.KeyData.ToString().Equals("Down"))
                    {
                        //if (!pressAlt)
                        //    AlterControl();
                        if (!pressF2 && !at)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.operateFull));
                        }
                        //else if (pressAlt)
                        //{
                        //    NarratorRunOrNotCheck();
                        //}
                    }

                    else if (insrt && (e.KeyData.ToString().Equals("F9")))
                    {
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(bookThread.Test));
                    }
                    else if (insrt && e.KeyCode.Equals(Keys.F12))  // say system time insert+f12
                    {
                        e.Handled = true;
                        //String dat = "Todays Date" + DateTime.Now.ToString("d");
                        //speaker.speak(dat);
                        //String time = "  and  Current Time " + DateTime.Now.ToString("T");
                        //speaker.speak(time);
                    }                    

                    else if (e.KeyValue == 186)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.stop();

                        if (pressCtrl && shift && !at)
                            speaker.speak("insert time" + DateTime.Now.ToString("hh:mm tt"));

                        else if (pressCtrl && !at)
                            speaker.speak("insert date" + DateTime.Now.ToString("d/MM/yyyy"));
                    }

                    else if (e.KeyData.ToString().Equals("Next") || e.KeyData.ToString().Equals("PageUp"))
                    {
                        if (pressCtrl && shift && !at)
                        {
                            //MessageBox.Show("keyData");
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.SheetInstruction));
                        }
                        //else if ((pressCtrl && pressAlt) || (pressCtrl && !pressAlt))
                        //{
                        //    ctr_alt_pgup = 1;
                        //}
                        else if (insrt && e.KeyData.ToString().Equals("PageUp"))
                        {
                            e.Handled = true;
                            thread = new Thread(new ThreadStart(bookThread.GetRightText));
                        }
                    }
                    else if (e.KeyData.ToString().Equals("Tab"))
                    {
                        speaker.speak(e.KeyData.ToString());
                        //if (!pressAlt && !releaseAlt)
                        //    AlterControl();
                        ////MessageBox.Show(pressAlt.ToString())
                        //if (releaseAlt) releaseAlt = false;
                        if (pressF2 && !at)//////////////////kkk
                        {
                            //speaker.speak("Tab");
                            NarratorStop();
                            pressF2 = false;
                            ED_mode = 0;
                            ob.unhook();
                        }
                        else if (!pressF2 && !at)
                        {
                            //speaker.speak("Tab");
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.operateFull));
                        }
                        //else if (pressAlt)
                        //    NarratorRunOrNotCheck();
                    }
                    else if ((e.KeyData.ToString().Equals("E")) || (e.KeyData.ToString().Equals("e")))
                    {
                        if (pressCtrl && !at)
                        {

                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.HeaderInstruction));
                        }
                    }
                    else if ((e.KeyData.ToString().Equals("T")) || (e.KeyData.ToString().Equals("t")))
                    {
                        if (pressCtrl && !at)
                        {
                            e.Handled = true;
                        }
                    }
                    //else if (e.KeyData.ToString().Equals("D1") || e.KeyCode.Equals(Keys.G))
                    //{
                    //    if (pressCtrl && !at)
                    //    {
                    //        pressAlt = true;
                    //        NarratorRunOrNotCheck();
                    //    }
                    //}

                    else if (e.KeyData.ToString().Equals("M") || (e.KeyData.ToString().Equals("m")))
                    {
                        if (pressCtrl && !at)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.operateFull));
                        }
                    }

                    else if (e.KeyData.ToString().Equals("D6"))
                    {
                        if (pressCtrl && !at)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.GetRowText));

                        }
                    }

                    else if (e.KeyData.ToString().Equals("D7"))
                    {
                        if (pressCtrl && !at)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.GetColumnText));
                        }
                    }                    

                    //else if (e.KeyValue == 33)
                    //{
                    //    if (thread != null) { thread.Abort(); thread = null; }
                    //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //    speaker.speak("Page up");
                    //}
                    //else if (e.KeyValue == 34)
                    //{
                    //    if (thread != null) { thread.Abort(); thread = null; }
                    //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //    speaker.speak("Page Down");
                    //}

                    //else if ((e.KeyCode.Equals(Keys.PageDown)) || (e.KeyCode.Equals(Keys.PageUp)))
                    //{
                    //    //MessageBox.Show("keyData");


                    //}
                    else if ((e.KeyCode.Equals(Keys.Space)))
                    {
                        speaker.speak("Space");
                        //if (pressAlt)
                        //{
                        //    NarratorStop();
                        //    pressAlt = false;
                        //}

                        if (shift && !at)
                        {
                            //NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.SelectedCurrentRow));
                        }

                        else if (pressCtrl && !at)
                        {
                            //NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.SelectedCurrentColumn));
                        }

                    }
                    else if (e.KeyValue == 13)   // Enter Key
                    {
                        if (pressF2 && !at)////////////////////////kkk
                        {
                            pressF2 = false;
                            ED_mode = 0;
                            ob.unhook();
                            //MessageBox.Show("uu");
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            ExcelReadThread oob = new ExcelReadThread();
                            oob.set_track_F2();
                        }
                        if (!at)
                        {
                            NarratorStop();
                            if (thread != null) { thread.Abort(); thread = null; }
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            speaker.speak("Enter");
                        }
                        //else if (pressAlt)
                        //{
                        //    pressAlt = false;
                        //}
                    }
                    else if (shift && (e.KeyData.ToString().Equals("D1")))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("exclamation");
                    }
                    else if (shift && (e.KeyData.ToString().Equals("D2")))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("At");
                    }
                    else if (shift && (e.KeyData.ToString().Equals("D3")))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Hash");
                    }
                    else if (shift && (e.KeyData.ToString().Equals("D4")))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Dollar");
                    }
                    else if (shift && (e.KeyData.ToString().Equals("D5")))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("percentage");
                    }
                    else if (shift && (e.KeyData.ToString().Equals("D6")))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("cap");
                    }
                    else if (shift && (e.KeyData.ToString().Equals("D7")))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("ampersand");
                    }
                    else if (shift && (e.KeyData.ToString().Equals("D8")))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Star");
                    }
                    else if (shift && (e.KeyData.ToString().Equals("D9")))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("First bracket Open");
                    }
                    else if (shift && (e.KeyData.ToString().Equals("D0")))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("First bracket Close");
                    }
                    else if (shift && (e.KeyValue == 186))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Colon");
                    }
                    else if (shift && (e.KeyValue == 187))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Plus");
                    }
                    else if (shift && (e.KeyValue == 188))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("less than");
                    }

                    //else if ((pressCtrl && (e.KeyData.ToString().Equals("g"))) || (pressCtrl && (e.KeyData.ToString().Equals("G"))))
                    //{
                    //    pressAlt = true;
                    //    //MessageBox.Show("p");
                    //}

                    else if (shift && (e.KeyValue == 189))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Under score");
                    }
                    else if (shift && (e.KeyValue == 190))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("greater than");
                    }
                    else if (shift && (e.KeyValue == 191))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Question");
                    }
                    else if (shift && (e.KeyValue == 219))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Second bracket open");
                    }
                    else if (shift && (e.KeyValue == 220))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("vertical line");
                    }
                    else if (shift && (e.KeyValue == 221))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Second bracket close");
                    }
                    else if (shift && (e.KeyValue == 222))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("double quotation");
                    }
                    else if (shift && (e.KeyValue == 192))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Equivalency sign");
                    }

                    else if (e.KeyValue >= 48 && e.KeyValue <= 57)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak(e.KeyData.ToString());
                    }

                    else if (e.KeyValue == 187)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Equal");
                    }
                    else if (e.KeyValue == 188)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("comma");
                    }
                    else if (e.KeyValue == 189)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("minus");
                    }
                    else if (e.KeyValue == 190)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("fullstop");
                    }
                    else if (e.KeyValue == 191)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("slash");
                    }
                    else if (e.KeyValue == 219)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("third bracket open");
                    }
                    else if (e.KeyValue == 220)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("back slash");
                    }
                    else if (e.KeyValue == 221)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("third bracket close");
                    }
                    else if (e.KeyValue == 222)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("single quotation");
                    }
                    else if (e.KeyData.ToString().Equals("OMPERIOD"))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Fullstop");
                    }

                    else if (e.KeyData.ToString().Equals("Back"))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Backspace");
                    }
                    else if (e.KeyData.ToString().Equals("Delete"))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Delete");

                    }
                    else if (e.KeyData.ToString().Equals("Escape"))
                    {
                        //x = 0;
                        //y = 0;
                        //x = Cursor.Position.X;
                        //y = Cursor.Position.Y;
                        //speaker.speak("x=" + x.ToString() + "  y=" + y.ToString());

                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Escape");

                        if (pressF2 && !at)///////////////////////kkk
                        {
                            ED_mode = 0;
                            pressF2 = false;
                            ob.unhook();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(bookThread.set_track_F2));
                        }

                        //if (pressAlt)
                        //{
                        //    pressAlt = false;
                        //    NarratorStop();
                        //}

                    }



                    else if (e.KeyValue == 192)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Grave accent");
                    }

                //
                    else if (e.KeyValue == 20)
                    {
                        caps = caps + 1;
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (Control.IsKeyLocked(Keys.CapsLock))
                        {
                            speaker.speak("caps lock turns off");
                        }
                        else
                            speaker.speak("caps lock turns on");
                        //if (caps % 2 == 0) speaker.speak("caps lock off");
                        //else speaker.speak("caps lock on");

                    }

                    else if (e.KeyValue == 36)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Home");
                        if (pressF2)
                            thread = new Thread(new ThreadStart(bookThread.tr_F2));
                        else if (insrt)
                        {
                            e.Handled = true;
                            thread = new Thread(new ThreadStart(bookThread.GetLeftText));
                        }
                    }
                    else if (e.KeyValue == 45)
                    {
                        insrt = true;
                        insert = insert + 1;
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (insert % 2 == 0) speaker.speak("Insert off");
                        else speaker.speak("Insert on");
                    }
                    else if (e.KeyValue == 44)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Print Screen");
                    }
                    else if (e.KeyValue == 19)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Pause Break");
                    }
                    else if (e.KeyValue == 123)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("F12");
                    }
                    else if (e.KeyValue == 91)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Windows");
                    }
                    else if (e.KeyValue == 35)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("End");
                    }
                    else if (e.KeyValue == 93)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Cursor Key");
                    }
                    if (thread != null)
                    {
                        thread.Start();
                    }

                    if (pressF2 && !pressAlt)
                    {
                        if (shift)
                        {
                            bookThread.update1();
                        }
                        else
                        {
                            bookThread.update();
                        }
                    }
                    if ((code >= 'A' && code <= 'Z') || (code >= 'a' && code <= 'z'))
                    {
                        //MessageBox.Show(e.KeyData.ToString());
                        speaker.speak(e.KeyData.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }


        }
        private void HookManager_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                Thread thread = null;
                if (e.KeyData.ToString().Equals("Escape"))
                {
                    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    if (thread != null) { thread.Abort(); thread = null; }
                    NarratorStop();
                    thread = new Thread(new ThreadStart(bookThread.RinbbonUnselect));
                    thread.Start();
                }
                if (e.KeyData.ToString().Equals("LControlKey") || e.KeyData.ToString().Equals("RControlKey"))
                {
                    speaker.speak("Control");
                    pressCtrl = false;
                }

                if (e.KeyData.ToString().Equals("LMenu") || e.KeyData.ToString().Equals("RMenu"))
                {
                    pressAlt = false;
                }

                if (e.KeyData.ToString().Equals("LShiftKey") || e.KeyData.ToString().Equals("RShiftKey"))
                {
                    speaker.speak("Shift");
                    shift = false;
                }

                if (e.KeyData.ToString().Equals("Insert"))
                {
                    insrt = false;
                }
            }
            catch (Exception ex)
            { }
        }
    }
}
