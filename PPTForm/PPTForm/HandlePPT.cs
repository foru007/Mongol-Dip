using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PPT = Microsoft.Office.Interop.PowerPoint;
using SpeechBuilder;
using System.Windows.Forms;
using System.IO;
using ThesisMain;
using System.Threading;
using Gma.UserActivityMonitor;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Configuration;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.Diagnostics;

namespace PPTForm
{
    public class HandlePPT
    {
        private static PPT.Application application = null;
        private PPT._Presentation presentation = null;
        private PPT._Slide slide = null;

        private Communication communication;
        private SpeechControl speaker;

        private PPTReadThread pptThread;

        private object fileName = null;
        private object readOnly = false;
        private TextReader reader;
        private object missing = Type.Missing;
        object isVisible = true;

        private Boolean pressCtrl = false;
        private Boolean pressAlt = false, altPress = false;
        private Boolean releaseAlt = false;
        private Boolean pressF5 = false;
        private Boolean shift = false;

        Thread thread1 = null;

        private String fName; //string
        private String fullText;
        private int p = 0;
        private int x, y, z = 0;
        String pp = null;
        private static int d = 0;
        private static int caps = 0, insert = 0;

        private int kj = 0;

        Process proc = new Process();

        private Boolean insrt = false;

        public HandlePPT(String fileName, SpeechControl speaker, PPT.Application application, int p)
        {
            this.p = p;
            application = application;
            this.fileName = fileName;
            this.speaker = speaker;
            this.fName = fileName;
            fullText = null;

            //HookManager.KeyDown += HookManager_KeyDown;
            startHandle();
        }
        public HandlePPT(SpeechControl speaker, PPT.Application application, int p1)
        {
            this.p = p1;
            application = application;
            this.speaker = speaker;
            //ob = new Form2();
            //startHandleWithoutFileName();

        }
        public HandlePPT(SpeechControl speaker, PPT.Application application1)
        {

            application = application1;
            this.speaker = speaker;
            //ob = new Form2();
            //startHandleWithoutFileName();
            NewstartHandle();

        }

        ~HandlePPT()
        {
            HookManager.KeyUp -= HookManager_KeyUp;
            HookManager.KeyDown -= HookManager_KeyDown;
        }
        public void NewstartHandle()
        {
            try
            {
                HookManager.KeyUp += HookManager_KeyUp;
                HookManager.KeyDown += HookManager_KeyDown;
            }
            catch (Exception e)
            {
            }
        }
        public void startHandle()
        {
            if (p == 0) d = 0;
            d++;
            if (d == 1)
            {
                HookManager.KeyUp += HookManager_KeyUp;
                HookManager.KeyDown += HookManager_KeyDown;
            }
            try
            {
                application.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                presentation = application.Presentations.Open(fName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
                application.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMaximized;
                application.Activate();
                application.PresentationClose += new Microsoft.Office.Interop.PowerPoint.EApplication_PresentationCloseEventHandler(application_PresentationClose);

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
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
                application.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                presentation = application.Presentations.Add(MsoTriState.msoTrue);
                application.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMaximized;
                application.PresentationClose += new Microsoft.Office.Interop.PowerPoint.EApplication_PresentationCloseEventHandler(application_PresentationClose);

            }
            catch (COMException)
            {
                //MessageBox.Show("Error accessing Word document.");
            }

        }
        private void newApplication_PresentationClose(PPT.Presentation p)
        {
            HookManager.KeyUp -= HookManager_KeyUp;
            HookManager.KeyDown -= HookManager_KeyDown;
            try
            {
                Process[] procs = Process.GetProcessesByName("POWERPNT");
                foreach (Process proc in procs)
                    proc.Kill();
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            }

        }
        private void application_PresentationClose(PPT.Presentation p)
        {
            //presentation.Save();
            //p.Save();
            //HookManager.KeyUp -= HookManager_KeyUp;
            //HookManager.KeyDown -= HookManager_KeyDown;
            ////System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
            ////p.Close();
            ////application.Quit();
            //try
            //{
            //    Process[] procs = Process.GetProcessesByName("POWERPNT");
            //    foreach (Process proc in procs)
            //        proc.Kill();
            //}
            //catch (Exception ex)
            //{
            //    //Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            //}
        }
        public void NarratorRunOrNotCheck()
        {
            // Process is Running or not
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

        //private void AlterControl()
        //{
        //    String s = communication.getWindowName().ToString();
        //    String sx = s.Substring(0, 20);
        //    MessageBox.Show(sx);
        //    pp = null;

        //    if (sx != "Microsoft PowerPoint")
        //    {
        //        //MessageBox.Show("true");
        //        pressAlt = true;
        //        z = z + 1;
        //    } 
        //}
        private void AlterControl()
        {
            // edited
            int length = 20;
            String s = communication.getWindowName().ToString();
            if (s.Length < length) length = s.Length;
            String sx = s.Substring(0, length);
            //MessageBox.Show(sx);
            pp = null;
            if (sx != "Microsoft PowerPoint")
            {
                pressAlt = true;
                z = z + 1;
            }
        }
        private void HookManager_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                Console.WriteLine("kj on pp" + kj.ToString());
                communication = new Communication();
                Thread thread = null;
                //MessageBox.Show(e.KeyValue.ToString());
                pptThread = new PPTReadThread(speaker, application, presentation, e.KeyData.ToString());
                Char code = (char)e.KeyCode;

                if (!communication.GetActiveProcess().ToString().Equals("POWERPNT") && pptThread != null)
                {
                    kj = 0;
                    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    pptThread.stopAll();
                    //ob.unhook();
                    pressAlt = false;
                }

                if (communication.GetActiveProcess().ToString().Equals("POWERPNT"))
                {
                    if (kj == 0)
                    {
                        HookManager.KeyUp -= HookManager_KeyUp;
                        HookManager.KeyDown -= HookManager_KeyDown;
                        speaker.speak("powerpont file testing");
                        HookManager.KeyUp += HookManager_KeyUp;
                        HookManager.KeyDown += HookManager_KeyDown;
                        kj = kj + 1;
                    }

                    //speaker.speak("powerpoint file");

                    if (e.KeyData.ToString().Equals("LControlKey") || e.KeyData.ToString().Equals("RControlKey"))
                    {
                        speaker.speak("Control");
                        speaker.stop();
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        pressCtrl = true;
                        if (releaseAlt)
                            releaseAlt = false;
                    }
                    else if (e.KeyData.ToString().Equals("LMenu") || e.KeyData.ToString().Equals("RMenu") || (e.KeyData.ToString().Equals("F10")))
                    {
                        altPress = true;
                        NarratorRunOrNotCheck();
                        releaseAlt = true;
                        if (pressCtrl)
                            releaseAlt = false;
                    }

                    else if (e.KeyData.ToString().Equals("LShiftKey") || e.KeyData.ToString().Equals("RShiftKey"))
                    {
                        speaker.speak("Shift");
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        shift = true;
                    }
                    else if (e.KeyData.ToString().Equals("PageUp") || e.KeyData.ToString().Equals("Next"))
                    {
                        if (e.KeyData.ToString().Equals("PageUp"))
                            speaker.speak("page up");
                        else if (e.KeyData.ToString().Equals("Next"))
                            speaker.speak("page down");

                        if (pressF5 && !pressAlt)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.SlideShow));
                        }
                        else if (!pressF5 && !pressAlt)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.operateCharacter));
                        }
                    }
                    else if (e.KeyData.ToString().Equals("Back"))
                    {
                        if (pressF5 && !pressAlt)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            speaker.speak("Backspace");
                            thread = new Thread(new ThreadStart(pptThread.SlideShow));
                        }
                        else
                        {
                            if (thread != null) { thread.Abort(); thread = null; }
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            speaker.speak("Backspace");
                        }
                    }
                    else if (insrt && (e.KeyData.ToString().Equals("F3")))
                    {
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(pptThread.FontInfo));
                    }
                    else if (insrt && (e.KeyData.ToString().Equals("F4")))
                    {
                        e.Handled = true;
                    }

                    else if (insrt && (e.KeyData.ToString().Equals("F8")))
                    {
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(pptThread.OldChar));
                    }

                    else if (insrt && (e.KeyData.ToString().Equals("F9")))
                    {
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(pptThread.OldWord));
                    }
                    else if (insrt && (e.KeyData.ToString().Equals("F11")))
                    {
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(pptThread.OldPara));
                    }
                    else if (insrt && e.KeyCode.Equals(Keys.F12))  // say system time insert+f12
                    {
                        e.Handled = true;
                        //String dat = "Todays Date" + DateTime.Now.ToString("d");
                        //speaker.speak(dat);
                        //String time = "  and  Current Time " + DateTime.Now.ToString("T");
                        //speaker.speak(time);
                    }
                    else if (insrt && (e.KeyData.ToString().Equals("T")))
                    {
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(pptThread.CurrentPPTName));
                    }

                    else if ((e.KeyData.ToString().Equals("P")))
                    {
                        if (insrt && pressF5 && !pressAlt)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.SlideShowSlideNo));
                        }
                        else if (pressF5 && !pressAlt)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.SlideShow));
                        }
                    }

                    else if (e.KeyData.ToString().Equals("Space"))
                    {
                        if (pressF5 && !pressAlt)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            speaker.stop();
                            speaker.speak("Space");
                            thread = new Thread(new ThreadStart(pptThread.SlideShow));
                        }
                        else
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            speaker.speak("Space");
                        }
                    }
                    //else if (pressAlt && e.KeyData.ToString() != "Down" && e.KeyData.ToString() != "Return")
                    //{
                    //    MessageBox.Show(e.KeyData.ToString());
                    //}

                    else if (e.KeyData.ToString().Equals("Escape"))
                    {
                        if (insrt)
                        {
                            e.Handled = true;
                        }
                        else
                        {
                            if (thread != null) { thread.Abort(); thread = null; }
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            speaker.speak("Escape");
                            NarratorStop();
                            pressAlt = false;
                            pressF5 = false;
                        }
                    }
                    //else if ((e.KeyData.ToString().Equals("F7")))
                    //{
                    //    pressAlt = true;
                    //    NarratorRunOrNotCheck();
                    //}
                    else if ((e.KeyData.ToString().Equals("F5")))
                    {
                        pressF5 = true;
                        pressAlt = false;
                        NarratorStop();

                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(pptThread.ShowControl));
                    }
                    else if (e.KeyData.ToString().Equals("Right") || e.KeyData.ToString().Equals("Left") || e.KeyData.ToString().Equals("Up") || e.KeyData.ToString().Equals("Down"))
                    {
                        //speaker.speak(e.KeyData.ToString());
                        if (!pressF5)
                        {
                            AlterControl();
                            NarratorRunOrNotCheck();
                        }
                        if (pressF5 && !pressAlt)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.SlideShow));
                        }
                        else if (pressCtrl && shift && !pressAlt)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.operateSelection));
                        }

                        else if (pressCtrl && !pressAlt)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.operateWord));
                        }
                        else if (shift && !pressAlt)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.operateSelection));
                        }
                        else if (!pressAlt)
                        {
                            NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.operateCharacter));
                        }

                    }

                    else if (e.KeyData.ToString().Equals("n") || e.KeyData.ToString().Equals("N"))
                    {
                        if (!pressF5 && !pressAlt && pressCtrl)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.ReadWhenPressN));
                            //speaker.speak("You press control plus n to create new powerpoint document");
                            //System.Threading.Thread.Sleep(6000);
                        }
                        if (pressF5 && !pressAlt)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.SlideShow));
                        }
                    }

                    else if (!pressAlt && (e.KeyData.ToString().Equals("Home") || e.KeyData.ToString().Equals("End")))
                    {
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (pressCtrl)
                        {
                            thread = new Thread(new ThreadStart(pptThread.operateCharacter));

                        }
                        if (e.KeyData.ToString().Equals("Home")) speaker.speak("Home");
                        else if (e.KeyData.ToString().Equals("End")) speaker.speak("End");
                    }


                    else if (e.KeyData.ToString().Equals("m") || e.KeyData.ToString().Equals("M"))
                    {
                        if (!pressF5 && !pressAlt && pressCtrl)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.operateCharacter));
                        }
                    }
                    //else if (e.KeyData.ToString().Equals("Up") || e.KeyData.ToString().Equals("Down"))
                    //{
                    //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //    if (thread != null) { thread.Abort(); thread = null; }
                    //    thread = new Thread(new ThreadStart(pptThread.operateFull));

                    //}
                    else if (e.KeyData.ToString().Equals("Tab"))
                    {
                        speaker.speak("Tab");
                        if (releaseAlt) releaseAlt = false;
                        if (!pressAlt)
                            AlterControl();

                        if (!pressAlt && !pressF5)
                        {
                            //NarratorStop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.operateCharacter));
                        }
                        if (!pressAlt && insrt)
                        {
                            e.Handled = true;
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(pptThread.SlideRd));
                        }

                        if (pressAlt)
                        {
                            NarratorRunOrNotCheck();
                        }

                    }

                    else if (e.KeyData.ToString().Equals("Return"))
                    {
                        if (pressAlt == false)
                            AlterControl();
                        //if (pressAlt == false && application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue)
                        //{
                        //    NarratorStop();
                        //    speaker.speak("Enter");
                        //    if (thread != null) { thread.Abort(); thread = null; }
                        //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        //    thread = new Thread(new ThreadStart(pptThread.operateInsideTable));
                        //}
                        if (pressAlt == false)
                        {
                            //MessageBox.Show("lllll");
                            NarratorStop();
                            speaker.speak("Enter");
                            if (thread != null) { thread.Abort(); thread = null; }
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            thread = new Thread(new ThreadStart(pptThread.operateFull));

                        }
                        else if (pressAlt == true)
                        {
                            pressAlt = false;
                        }
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
                    else if (e.KeyValue == 186)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("semicolon");
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


                    else if (e.KeyData.ToString().Equals("Delete"))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        //speaker.speak("Delete");
                        thread = new Thread(new ThreadStart(pptThread.operateCharacter));

                    }
                    //else if (e.KeyData.ToString().Equals("Escape"))
                    //{
                    //    if (thread != null) { thread.Abort(); thread = null; }
                    //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //    speaker.speak("Question");
                    //}

                    //else if (e.KeyValue == 13)   // Enter Key
                    //{
                    //}


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

                    }


                    //else if (e.KeyValue == 36)
                    //{
                    //    if (thread != null) { thread.Abort(); thread = null; }
                    //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //    speaker.speak("Home");
                    //}
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
                    //else if (e.KeyValue == 35)
                    //{
                    //    if (thread != null) { thread.Abort(); thread = null; }
                    //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //    speaker.speak("End");
                    //}
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
                    NarratorStop();
                    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    if (thread != null) { thread.Abort(); thread = null; }
                    thread = new Thread(new ThreadStart(pptThread.RinbbonUnselect));
                    thread.Start();
                }
                if (e.KeyData.ToString().Equals("LControlKey") || e.KeyData.ToString().Equals("RControlKey"))
                {
                    //speaker.speak("Control");
                    pressCtrl = false;
                }

                if (e.KeyData.ToString().Equals("LMenu") || e.KeyData.ToString().Equals("RMenu") || (e.KeyData.ToString().Equals("F10")))
                {
                    altPress = false;
                    if (releaseAlt)
                    {
                        if (pressAlt == true)
                        {
                            NarratorStop();
                            pressAlt = false;
                            //speaker.speak("Alter");  //  false
                            //ob.unhook();
                        }
                        else
                        {
                            //speaker.speak("Alter");  // true
                            pressAlt = true;
                            //count = 0;
                            //docThread.set(0);
                        }
                    }
                    else
                    {
                        pressAlt = false;
                        //releaseAlt = true;
                    }
                    releaseAlt = false;

                    if (!altPress)
                        pressAlt = false;
                }

                if (e.KeyData.ToString().Equals("LShiftKey") || e.KeyData.ToString().Equals("RShiftKey"))
                {
                    //speaker.speak("Shift");
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
