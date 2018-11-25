using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Gma.UserActivityMonitor;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.IO;
using SpeechBuilder;
using ThesisMain;
using EPocalipse.IFilter; // my

namespace DocForm
{
    public class HandleADoc
    {
        private Word.Application word = null;
        private Word._Document doc = null;
        /*
Use the Documents property to return the Documents collection.

Use the Add method to create a new empty document and add it to the Documents collection.

Use the Open method to open a file.

Use Documents(index), where index is the document name or index number to return a single Document object.

The index number represents the position of the document in the Documents collection.
*/

        private object fileName = null;
        private object readOnly = false;
        private object missing = Type.Missing;
        object isVisible = true;
        /*
Missing is used to invoke a method with a default argument.

Only one instance of Missing ever exists.
         */

        private Boolean pressCtrl = false;
        private Boolean pressAlt = false;
        private Boolean shift = false;


        private DocReadThread docThread;
        private SpeechControl speaker;
        private Communication communication;
        private TextReader reader;
        private String fName; //string
        private String fullText;
        Thread thread = null;
        Thread thread1 = null;
        public static int count = 0;
        private int p = 0;
        private int xx = 0;
        private static int d = 0;
        private Form2 ob;
        private static int caps = 0, insert = 0;
        Process proc = new Process();
        String pp = null;
        private int kj = 0;
        // private Boolean releaseAlt = false;
        private Boolean insrt = false;
        //  private Boolean alt = false;
        private Boolean ctrl = false;
        public Boolean at = false;


        Process[] procs = Process.GetProcessesByName("WINWORD");

        public HandleADoc(String fileName, SpeechControl speaker, Word.Application word1, int p1)
        {
            this.p = p1;
            word = word1;
            this.fileName = fileName;
            this.speaker = speaker;
            reader = null;
            this.fName = fileName;
            fullText = null;
            ob = new Form2();
            //HookManager.KeyDown += HookManager_KeyDown;
            startHandle();
        }
        public HandleADoc(SpeechControl speaker, Word.Application word, int p1)
        {
            this.p = p1;
            this.word = word;
            this.speaker = speaker;
            //ob = new Form2();
            startHandleWithoutFileName();

        }
        public HandleADoc(SpeechControl speaker, Word.Application word)
        {
            //this.p = p1;
            this.word = word;
            this.speaker = speaker;
            //ob = new Form2();
            //startHandleWithoutFileName();
            NewstartHandle();
        }
        public HandleADoc()
        {
        }

        ~HandleADoc()
        {
            HookManager.KeyUp -= HookManager_KeyUp;
            HookManager.KeyDown -= HookManager_KeyDown;
        }

        /*aditi
        public void setFileName( String fName )
        {
            this.fileName = fName;
            this.fName = fName;
            Console.WriteLine("***************  ****************", fName);
            
        }*/
        public void AsstForWord()
        {
            speaker.stop();
            String ss = "Instruction of how to operate  Word Document Press  Ctrl + Numpad 0 or Ctrl + D7. মাইক্রোসফট ওয়ার্ড ডকুমেন্ট চালানোর নির্দেশিকা জানতে কন্ট্রোল এবং নামপ্যাড ০ অথবা  কন্ট্রোল এবং ডি সেভেন বাটন চাপ দিন। To Listen One Character From Left Press Left Arrow. বাম পাশের একটি অক্ষর পরার জন্য বাম তীর চাপ দিন।  To Listen One Character From Right Press  Right Arrow. ডান পাশের একটি অক্ষর পরার জন্য ডান তীর চাপ দিন।  To Listen One Word From left Press  Ctrl + Right Arrow. ডান পাশের একটি অক্ষর পরার জন্য ডান তীর চাপ দিন। To Listen One Word From Right Press  Ctrl + Left Arrow. বাম পাশের একটি শব্দ পরার জন্য কনট্রোল এবং বাম তীর চাপ দিন। To Listen One Up Paragraph Press  Ctrl + Up Arrow. বাম পাশের একটি প্যারাগ্রাফ পরার জন্য কনট্রোল এবং উপর তীর চাপ দিন। To Listen One Down Paragraph Press Ctrl + Down Arrow. ডান পাশের একটি প্যারাগ্রাফ পরার জন্য কনট্রোল এবং নিচ তীর চাপ দিন। To Listen Current Page Number, Current Text Font Size, Current Text Font Name Instruction Press F2. বর্তমান পেইজ নাম্বার, টেক্সট ফন্ট সাইজ, টেক্সট ফন্ট নাম জানতে এফ ২ বাটন চাপ দিন। To Detect Table Cell Position Press Tab/Up. টেবিল সেল এর অবস্থান জানতে ট্যাব বাটন চাপ দিন। Line Number of Document Press Up or Down Key. বর্তমান লাইন নাম্বার জানতে উপর নিচ বাটন চাপ দিন। To create new office Document Press Ctrl + n. নতুন ডকুমেন্ট তৈরি করতে কন্ট্রোল এবং এন বাটন চাপ দিন। To Save Currently Working Word document Press Ctrl + S. ডকুমেন্ট  সেইভ করতে কন্ট্রোল এবং এফ ফউর বাটন চাপ দিন। To Save Current working Word document and close Press Altr + F4. ডকুমেন্ট  সেইভ করতে এবং বন্ধ করতে অলটার এবং এফ ফউর বাটন চাপ দিন।";
            speaker.speak(ss);

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
                word.Visible = true;
                doc = word.Documents.Add(ref missing, ref missing, ref missing, ref isVisible);
                doc.Activate();
                word.Activate();
                word.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(newWord_DocumentBeforeClose);

            }
            catch (COMException)
            {
                //MessageBox.Show("Error accessing Word document.");
            }
            //textBoxLog.Text = doc.Content.Text;
            //word.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(oWord_DocumentBeforeClose);
            ////DocOpenTrack = 0;

        }
        public void NewstartHandle()
        {
            try
            {
                HookManager.KeyUp += HookManager_KeyUp;
                HookManager.KeyDown += HookManager_KeyDown;
                word.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(oWord_DocumentBeforeClose);
            }
            catch (Exception e)
            {
            }
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
                Process[] procs = Process.GetProcessesByName("EXCEL");
                if (procs.Length != 0)
                {
                    foreach (Process proc in procs)
                        proc.Kill();
                }

                Process[] procss = Process.GetProcessesByName("POWERPNT");
                if (procss.Length != 0)
                {
                    foreach (Process procc in procss)
                        procc.Kill();
                }
                word.Visible = true;
                doc = word.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                doc.Activate();
                word.Activate();
            }
            catch (COMException)
            {
                MessageBox.Show("Error accessing Word document.");
            }

            //textBoxLog.Text = doc.Content.Text;
            //word.DocumentBeforeClose +=new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(oWord_DocumentBeforeClose);

        }
        public void speakStartToEnd()
        {
            reader = new FilterReader(fName);
            fullText = reader.ReadToEnd();
            reader.Dispose();
        }
        private void newWord_DocumentBeforeClose(Word.Document document, ref bool Cancel)
        {
            //    HookManager.KeyUp -= HookManager_KeyUp;
            //    HookManager.KeyDown -= HookManager_KeyDown;
            //    try
            //    {
            //        Process[] procs = Process.GetProcessesByName("WINWORD");
            //        foreach (Process proc in procs)
            //            proc.Kill();
            //    }
            //    catch (Exception ex)
            //    {
            //        //Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            //    }
        }
        public void Close()
        {
            if (word != null)
            {
                word.DocumentBeforeClose -= new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(oWord_DocumentBeforeClose);
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(word);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        private void WindowSelectionChange(Word.Selection s)
        {
            //MessageBox.Show(s.Text.ToString());
            HookManager.KeyUp += HookManager_KeyUp;
            HookManager.KeyDown += HookManager_KeyDown;
            word.WindowSelectionChange -= new ApplicationEvents4_WindowSelectionChangeEventHandler(WindowSelectionChange);
        }
        private void oWord_DocumentBeforeClose(Word.Document document, ref bool Cancel)
        {
            int a = word.Documents.Count;
            if (a == 1)
            {
                HookManager.KeyUp -= HookManager_KeyUp;
                HookManager.KeyDown -= HookManager_KeyDown;
                word.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(WindowSelectionChange);
            }
            else if (a > 1)
            {
                HookManager.KeyUp -= HookManager_KeyUp;
                HookManager.KeyDown -= HookManager_KeyDown;
                //speaker.speak("winword file testing");
                HookManager.KeyUp += HookManager_KeyUp;
                HookManager.KeyDown += HookManager_KeyDown;
                kj = 1;
            }

        }
        //private void AlterControl()
        //{
        //    //pp = null;
        //    int length = 14;
        //    String s = communication.getWindowName().ToString();
        //    //MessageBox.Show(s);
        //    if (s.Length < length) length = s.Length;
        //    char[] a = s.ToCharArray();

        //    for (int i = s.Length - 1; i >= 0; i--)
        //    {
        //        pp = pp + a[i];
        //    }
        //    String sx = pp.Substring(0, length);
        //    //char[] b = sx.ToCharArray();
        //    pp = null;
        //    //for (int i = sx.Length - 1; i >= 0; i--)
        //    //{
        //    //    pp = pp + b[i];
        //    //}
        //    //MessageBox.Show(sx);
        //    if (sx != "droW tfosorciM") pressAlt = true;
        //    //else if (sx == "droW tfosorciM") pressAlt = false;
        //}

        public void full()
        {
            speaker.speak(fullText);
        }
        public void setAt()
        {
            at = true;
        }
        private void HookManager_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                communication = new Communication();
                thread = null;
                //word = (Word.Application)Microsoft.VisualBasic.Interaction.GetObject("", "Word.Application");
                docThread = new DocReadThread(speaker, word, doc, e.KeyData.ToString());
                //MessageBox.Show(e.KeyValue.ToString());
                Char code = (char)e.KeyCode;

                if (!communication.GetActiveProcess().ToString().Equals("WINWORD") && docThread != null)
                {
                    kj = 0;
                    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    docThread.stopAll();
                    pressAlt = false;
                }
                if (communication.GetActiveProcess().ToString().Equals("WINWORD"))
                {
                    //Win32API winapi = new Win32API();
                    //at = winapi.EvaluateCaretPosition();

                    if (kj == 0)
                    {
                        HookManager.KeyUp -= HookManager_KeyUp;
                        HookManager.KeyDown -= HookManager_KeyDown;
                        HookManager.KeyUp += HookManager_KeyUp;
                        HookManager.KeyDown += HookManager_KeyDown;
                        kj = kj + 1;
                    }
                    if (e.KeyData.ToString().Equals("LControlKey") || e.KeyData.ToString().Equals("RControlKey"))
                    {
                        ctrl = true;
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        //speaker.speak("Control");
                        pressCtrl = true;
                        //if (releaseAlt)
                        //    releaseAlt = false;
                    }
                    else if (e.KeyData.ToString().Equals("LMenu") || e.KeyData.ToString().Equals("RMenu"))
                    {
                        //alt = true;
                        //releaseAlt = true;
                        //if (pressCtrl)
                        //    releaseAlt = false;
                        pressAlt = true;

                    }
                    else if (e.KeyData.ToString().Equals("LShiftKey") || e.KeyData.ToString().Equals("RShiftKey"))
                    {
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        shift = true;
                    }
                    else if (e.KeyData.ToString().Equals("Right") || e.KeyData.ToString().Equals("Left"))
                    {
                        //  MessageBox.Show("at=" + at + " pressalt=" + pressAlt);
                        //if (releaseAlt) { releaseAlt = false; }
                        speaker.stop();
                        //speaker.speak(e.KeyData.ToString());
                        //if (!pressAlt)
                        //    AlterControl();
                        //NarratorRunOrNotCheck();
                        //if (pressCtrl && pressAlt)
                        //{
                        //    //speaker.speak("pressCtrl && pressAlt");
                        //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        //    if (thread != null) { thread.Abort(); thread = null; }
                        //    //NarratorRunOrNotCheck();
                        //    NarratorStop();
                        //    thread = new Thread(new ThreadStart(docThread.operateSentence));
                        //}
                        if (pressCtrl && shift && !at)    //11
                        {
                            //speaker.speak("pressCtrl && not pressAlt");                            
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.operateSelectedWordPara));
                        }
                        else if (pressCtrl && pressAlt)
                        {
                            speaker.stop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.slectPreviousAndNextCell));
                        }
                        else if (pressCtrl && !at)        //7      update
                        {
                            //speaker.speak("pressCtrl && not pressAlt");                            
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.operateWord));
                        }


                        else if (shift && !at)            // shift selected text to speech
                        {
                            //speaker.speak("shift");
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            //speaker.speak("hi i am a good boy");
                            //NarratorRunOrNotCheck();
                            thread = new Thread(new ThreadStart(docThread.operateSelect));
                        }
                        else if (!at)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            //NarratorRunOrNotCheck();                       
                            thread = new Thread(new ThreadStart(docThread.operateChar));
                        }
                    }
                    else if (pressCtrl && shift && (e.KeyData.ToString().Equals("Up") || e.KeyData.ToString().Equals("Down")))  //11
                    {
                        //speaker.speak("pressCtrl && not pressAlt");                            
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(docThread.operateSelectedWordPara));
                    }
                    else if (pressCtrl && pressAlt && (e.KeyData.ToString().Equals("Up") || e.KeyData.ToString().Equals("Down")))
                    {
                        speaker.stop();
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(docThread.slectPreviousAndNextRow));
                    }
                    else if (pressCtrl && (e.KeyData.ToString().Equals("Up") || e.KeyData.ToString().Equals("Down")))
                    {
                        e.Handled = true;
                        speaker.stop();
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(docThread.operateParagraph));
                    }
                    else if ((e.KeyData.ToString().Equals("Up")))
                    {
                        //if (releaseAlt) releaseAlt = false;
                        speaker.stop();
                        //speaker.speak(e.KeyData.ToString());
                        //if (!pressAlt)
                        //    AlterControl();
                        if (pressAlt)
                        {
                            e.Handled = true;
                            speaker.stop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.operatePriorSentence));
                            //releaseAlt = false;
                        }
                        else if (insrt && !at)    // say line insert+up arrow
                        {
                            e.Handled = true;
                            if (!at)
                            {
                                speaker.stop();
                                if (thread1 != null) { thread1.Abort(); thread1 = null; }
                                if (thread != null) { thread.Abort(); thread = null; }
                                thread = new Thread(new ThreadStart(docThread.SayLine));
                            }
                        }
                        else if (shift && !at)            // shift selected text to speech
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.operateSelect));
                        }
                        else if (!at)
                        {
                            speaker.stop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.operateChar));
                        }

                    }
                    else if (e.KeyData.ToString().Equals("Down"))
                    {
                        //if (releaseAlt) releaseAlt = false;
                        speaker.stop();
                        //if (!pressAlt)
                        //    AlterControl();
                        if (pressAlt)
                        {
                            e.Handled = true;
                            speaker.stop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.operatePriorSentence));
                            //releaseAlt = false;
                        }
                        else if (insrt && !at)           // say selected text shift+insert+down arrow 
                        {
                            e.Handled = true;
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.ShiftSelectText));
                        }
                        else if (shift && !at)            // shift selected text to speech
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.operateSelect));
                        }
                        else if (!at)
                        {
                            speaker.stop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.operateChar));
                        }
                    }
                    else if (e.KeyData.ToString().Equals("Tab"))
                    {
                        speaker.speak(" Tab ");
                        if (insrt)    // say window prompt and text insert+tab
                        {
                            e.Handled = true;
                            if (!at)
                            {
                                speaker.stop();
                                if (thread1 != null) { thread1.Abort(); thread1 = null; }
                                if (thread != null) { thread.Abort(); thread = null; }
                                thread = new Thread(new ThreadStart(docThread.operateChar));
                            }
                        }
                        else
                        {

                            //speaker.speak(e.KeyData.ToString());
                            //if (releaseAlt) releaseAlt = false;
                            //if (!pressAlt)
                            //    AlterControl();
                            if (!at)
                            {
                                //NarratorStop();
                                pressAlt = false;
                                if (thread1 != null) { thread1.Abort(); thread1 = null; }
                                if (thread != null) { thread.Abort(); thread = null; }
                                thread = new Thread(new ThreadStart(docThread.TabPress));
                            }
                        }
                    }

                    else if (insrt && e.KeyData.ToString().Equals("F1"))
                    {
                        e.Handled = true;
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(docThread.PageSetUp));
                    }
                    else if (insrt && e.KeyData.ToString().Equals("C"))
                    {
                        e.Handled = true;
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(docThread.TableCurCellInfo));
                    }
                    else if (insrt && e.KeyData.ToString().Equals("F"))
                    {
                        e.Handled = true;
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(docThread.Instraction));
                    }
                    else if (insrt && e.KeyData.ToString().Equals("D5"))
                    {
                        e.Handled = true;
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(docThread.Instraction));
                    }
                    else if (insrt && e.KeyCode.Equals(Keys.F12)) // say system time insert+f12
                    {
                        e.Handled = true;
                        //String dat = "Todays Date" + DateTime.Now.ToString("d");
                        //speaker.speak(dat);
                        //String time = "  and  Current Time " + DateTime.Now.ToString("T");
                        //speaker.speak(time);
                    }
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////end
                    else if (insrt && (e.KeyData.ToString().Equals("T")))
                    {
                        e.Handled = true;
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        if (thread != null) { thread.Abort(); thread = null; }
                        thread = new Thread(new ThreadStart(docThread.readTitle));
                    }
                    else if (pressCtrl && (e.KeyData.ToString().Equals("P")))   //////*********("NumPad1")))
                    {
                        //if (!pressAlt)
                        //{
                        //    pressAlt = true;
                        //}
                    }
                    /////////////////////////////////////////////
                    else if ((pressCtrl && e.KeyCode.ToString().Equals("x")) || (pressCtrl && e.KeyCode.ToString().Equals("X")))
                    {

                        try
                        {
                            Uri uri = new Uri(fileName.ToString());
                            string filename = Path.GetFileName(uri.LocalPath);
                            speaker.speak(filename);
                        }

                        catch
                        {
                            speaker.speak("you open a new document");
                        }
                    }
                    ///////////////////////////////////
                    else if ((pressCtrl && e.KeyCode.ToString().Equals("t")) || (pressCtrl && e.KeyCode.ToString().Equals("T")))
                    {

                        // word.Selection.Range.Information(wdFirstCharacterLineNumber);


                        // word.Selection.Range.ComputeStatistics()
                        // office line number and page number get.

                        // word.ActiveDocument.Range(0,word.Selection.Start).ComputeStatistics()


                        //int a = word.ActiveDocument.Range(0,Selection)
                        //    //word.Selection.Range.ComputeStatistics(word.
                        //    //word.Selection.Range.ComputeStatistics(Word.WdStatistic.wdStatisticLines);
                        //MessageBox.Show(a.ToString());

                        //word.ActiveDocument.Range(0,Selection)

                        //Selection.Range.Information(wdFirstCharacterLineNumber);

                    }

                    else if ((pressCtrl && (e.KeyCode.ToString().Equals("NumPad0"))) || (pressCtrl && (e.KeyCode.ToString().Equals("D7"))))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        thread1 = new Thread(new ThreadStart(AsstForWord));
                        thread1.Start();
                        ////speaker.speak("press control One for listening the whole document at a time");
                        //speaker.speak("press left key to know one character from left");
                        //speaker.speak("press right key to know one character from right");
                        //speaker.speak("Press control and right key for listening one word right");
                        //speaker.speak("Press control and left key for listening one word left");
                        ////speaker.speak("Press control alter right for one sentence right");
                        ////speaker.speak("Press control alter left key for one sentence left");
                        //speaker.speak("Press control up key for listening one paragraph up");
                        //speaker.speak("Press control D1 key for open new microsoft word document");
                        //speaker.speak("Press control D2 key for open new noteped document");
                        //speaker.speak("Press control D3 key for open Mail sending and asccess window");
                        //speaker.speak("To terminate this application press alter f4");

                    }
                    //else if (e.KeyData.ToString().Equals("F4"))
                    //{
                    //    if (thread != null) { thread.Abort(); thread = null; }
                    //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //    if (!pressAlt)
                    //    {
                    //        pressAlt = false;
                    //        alt = false;
                    //    }
                    //}
                    else if (e.KeyData.ToString().Equals("F4"))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        //speaker.speak("Space");
                        if (pressAlt)
                        {
                            pressAlt = false;
                            //AlterControl();
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
                    else if (insrt && (e.KeyData.ToString().Equals("D3")))   //8
                    {
                        e.Handled = true;
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        thread = new Thread(new ThreadStart(docThread.PgLineNo));
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
                    else if (e.KeyData.ToString().Equals("D1"))
                    {
                        speaker.speak(e.KeyData.ToString());
                        if (pressAlt)
                        {
                            e.Handled = true;
                            speaker.stop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.columnTitle));
                            //releaseAlt = false;
                        }
                    }
                    else if (e.KeyData.ToString().Equals("D7"))
                    {
                        speaker.speak(e.KeyData.ToString());
                        if (pressAlt)
                        {
                            e.Handled = true;
                            speaker.stop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.rowTitle));
                            //releaseAlt = false;
                        }
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

                    else if (e.KeyData.ToString().Equals("Back"))
                    {
                        //e.Handled = true;
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        thread = new Thread(new ThreadStart(docThread.backSpace));
                        //speaker.speak("Backspace");
                    }
                    else if (e.KeyData.ToString().Equals("Delete"))
                    {
                        if (insrt)
                        {
                            e.Handled = true;
                            if (thread != null) { thread.Abort(); thread = null; }
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            thread = new Thread(new ThreadStart(docThread.LineNumber));
                        }
                        else
                        {
                            if (thread != null) { thread.Abort(); thread = null; }
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            thread = new Thread(new ThreadStart(docThread.delete));
                            speaker.speak("Delete");
                        }
                    }
                    else if (e.KeyData.ToString().Equals("Escape"))
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Escape");
                        if (pressAlt)
                        {
                            pressAlt = false;
                        }

                    }

                    else if (e.KeyValue == 13)   // Enter Key
                    {
                        if (pressAlt == false)
                        {
                            if (thread != null) { thread.Abort(); thread = null; }
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            speaker.speak("Enter");
                        }
                        else
                        {
                            pressAlt = false;
                        }
                        //MessageBox.Show("enter"+pressAlt.ToString());
                    }


                    if (e.KeyValue == 192)
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Grave accent");
                    }

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
                    else if (e.KeyValue == 33)  // pageup
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Page up");
                        if (pressCtrl)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            speaker.speak(" previous page");
                            thread = new Thread(new ThreadStart(docThread.PageNumber));
                        }
                    }
                    else if (e.KeyValue == 34)   //5
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Page Down");
                        if (pressCtrl)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            speaker.speak(" next page");
                            thread = new Thread(new ThreadStart(docThread.PageNumber));
                        }
                    }
                    else if (e.KeyValue == 36)   //5
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        if (thread1 != null) { thread1.Abort(); thread1 = null; }
                        speaker.speak("Home");
                        if (pressAlt && pressCtrl)
                        {
                            e.Handled = true;
                            speaker.stop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.slectFirstCell));
                            //releaseAlt = false;
                        }
                        else if (pressCtrl)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.goTopOfPage));
                        }
                        else if (shift && !at) // Shift + Home/ End Selected portion Echo  //11
                        {
                            //speaker.speak("pressCtrl && not pressAlt");                            
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.operateSelectedWordPara));
                        }
                        else if (!at) // Shift + Home/ End Selected portion Echo                          //3
                        {
                            //speaker.speak("pressCtrl && not pressAlt");                            
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.PressHome));
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
                        if (pressAlt && pressCtrl)
                        {
                            e.Handled = true;
                            speaker.stop();
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.slectLastCell));
                            //releaseAlt = false;
                        }
                        else if (pressCtrl)
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.goBottomOfPage));
                        }
                        else if (shift && !at)  // Shift + Home/ End Selected portion Echo   //11
                        {
                            //speaker.speak("pressCtrl && not pressAlt");                            
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.operateSelectedWordPara));
                        }
                        else if (insrt)  // say top line of window insert+end
                        {
                            e.Handled = true;
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.readTitle));
                        }
                        else if (!at)  // say top line of window insert+end          //3
                        {
                            if (thread1 != null) { thread1.Abort(); thread1 = null; }
                            if (thread != null) { thread.Abort(); thread = null; }
                            thread = new Thread(new ThreadStart(docThread.PressEnd));
                        }
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
                    if ((code >= 'A' && code <= 'Z') || (code >= 'a' && code <= 'z'))
                    {
                        //MessageBox.Show(e.KeyData.ToString());

                        if (Control.IsKeyLocked(Keys.CapsLock))
                        {
                            if (!shift)
                            {
                                speaker.stop();
                                speaker.speak(" Capital  " + e.KeyData.ToString());
                            }
                            else
                            {
                                speaker.stop();
                                speaker.speak(e.KeyData.ToString());
                            }
                        }
                        else if (shift)
                        {
                            speaker.stop();
                            speaker.speak(" Capital  " + e.KeyData.ToString());
                        }
                        else
                        {
                            speaker.stop();
                            speaker.speak(e.KeyData.ToString());
                        }
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
                Win32API winapi = new Win32API();
                at = winapi.EvaluateCaretPosition();
                thread = null;
                //if ((e.KeyData.ToString().Equals("End")))   //////*********("NumPad1")))
                //    {
                //        if (pressCtrl)
                //        {
                //            //speaker.speak(fullText);
                //            if (thread1 != null) { thread1.Abort(); thread1 = null; }                                        
                //            //docThread.set_fulltext(fullText);                    
                //            if (thread != null) { thread.Abort(); thread = null; }
                //            thread = new Thread(new ThreadStart(docThread.textread));
                //            thread.Start();
                //        }
                //    }
                //thread = new Thread(new ThreadStart(docThread.test));
                //thread.Start();
                if (e.KeyData.ToString().Equals("Escape"))
                {
                    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    if (thread != null) { thread.Abort(); thread = null; }
                    //NarratorRunOrNotCheck();
                    thread = new Thread(new ThreadStart(docThread.RinbbonUnselect));
                    thread.Start();
                }
                if (e.KeyData.ToString().Equals("LControlKey") || e.KeyData.ToString().Equals("RControlKey"))
                {
                    //speaker.speak("Control");
                    pressCtrl = false;
                    ctrl = false;
                }

                if (e.KeyData.ToString().Equals("LMenu") || e.KeyData.ToString().Equals("RMenu"))
                {
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


