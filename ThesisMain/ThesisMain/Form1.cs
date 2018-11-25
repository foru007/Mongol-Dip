using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Reflection;
using System.IO;
using Gma.UserActivityMonitor;
using SpeechBuilder;
using DocForm;
using System.Threading;
using RichTextEditor;
//using NetPopMimeClient;
using sendEmail;
//using Search;

namespace ThesisMain
{
    public partial class Form1 : Form
    {
        private Communication communication;
        private ImageAnalysis bitmapImage;
        private SpeechControl speaker;
        private Soket socket;
        private String activeWindow;
        private String sourcePath;
        private String destinationPath;
        private String sourceItem;
        static String Url;
        Thread thread1 = null;

        Process proc = new Process();
        private Boolean releaseAlt = false;
        private Boolean ctrl = false;
        private Boolean alter = false;
        //private Boolean shift = false;
        private Boolean insrt = false; 
        private static int ShutDownMonitor;
        private static int tempo = 44100;

        private String FN;

        public Form1(SpeechControl speaker)
        {
            InitializeComponent();

            bitmapImage = new ImageAnalysis();
            communication = new Communication();
            //speaker = new SpeechControl();
            this.speaker = speaker;

            activeWindow = null;
            HookManager.KeyUp += HookManager_KeyUp;
            HookManager.KeyDown += HookManager_KeyDown;

            sourcePath = null;
            destinationPath = null;
            sourceItem = null;
            //bitmapImage.GetScreenShot();
            //MessageBox.Show(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\tessdata");

        }
        public Form1(String u)
        {
            Url = u;
            //MessageBox.Show("GetTM"+Url);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            HookManager.KeyUp += HookManager_KeyUp;
            HookManager.KeyDown += HookManager_KeyDown;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            HookManager.KeyUp -= HookManager_KeyUp;
            HookManager.KeyDown -= HookManager_KeyDown;
        }
        private void HookManager_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {

                if (!isControllingVolume(e.KeyCode.ToString()))
                {
                    speaker.stop();
                }
                
            }
            catch (Exception)
            {
                //MessageBox.Show("Keydown exception of first try");
            }

            //When browser is selected
            /**************************************************************************/
            try
            {

                if (communication.getWindowName().ToString().Equals("Mongol Dip  Browser"))
                {
                    
                    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //if (ctrl && e.KeyCode.ToString().Equals("D1") || ctrl && e.KeyCode.ToString().Equals("NumPad1"))
                    ////////?????  numpad1 ???? ///??? new doc opening
                    //{
                    //    int p = 0;
                    //    //MessageBox.Show(fName);
                    //    //DT = new Thread(new ThreadStart(ddoc));
                    //    //DT.Start();
                    //    //Keys.Alt.Equals(false);
                    //    //if (word == null) word = new Microsoft.Office.Interop.Word.Application();
                    //    Process[] pros = Process.GetProcesses();
                    //    for (int i = 0; i < pros.Count(); i++)
                    //    {
                    //        if (pros[i].ProcessName.ToLower().Contains("winword"))
                    //        {
                    //            p++;
                    //            //if(p>1)pros[i].Kill();
                    //            //MessageBox.Show(p.ToString());
                    //        }
                    //    }
                    //    //MessageBox.Show(p.ToString());
                    //    if (p == 0)
                    //    {
                    //        word = new Microsoft.Office.Interop.Word.Application();
                    //        p = -1;
                    //    }

                    //    DocForm.DocForm wordApp = new DocForm.DocForm(speaker);
                    //    HandleADoc handleDoc = new HandleADoc(speaker, word, p);
                    //    handleDoc.startHandleWithoutFileName();
                    //    ///Thread thread = new Thread(new ThreadStart(handleDoc.startHandleWithoutFileName));
                    //    //thread.Start();
                    //}  
                    //else if (insrt && e.KeyCode.ToString().Equals("F4"))
                    //{
                    //    speaker.speak("Do You Really want to ShutDown Computer press y key to confirm shutdown now");
                    //    ShutDownMonitor = 1;
                    //}                                       
                    else if (ctrl && e.KeyCode.ToString().Equals("R"))
                    {
                        //assaf7.Events tr = new assaf7.Events();
                        //tr.Show();
                    }                   
                    //else if (ctrl && e.KeyCode.ToString().Equals("T"))
                    ////////?????  numpad2 ???? ///??? new text editor opening
                    //{
                    //    frmMain notepad = new frmMain(speaker);
                    //    notepad.Show();
                    //}
                    //else if (ctrl && e.KeyCode.ToString().Equals("M"))
                    ////////?????  numpad3 ???? ///??? mailing...but problem ase
                    //{
                    //    sendEmail.Form1 mailAgent = new sendEmail.Form1(speaker);
                    //    mailAgent.Show();
                    //}
                    else if (ctrl && e.KeyCode.ToString().Equals("U"))
                    //////?????  numpad3 ???? ///??? mailing...but problem ase
                    {
                        speaker.speak(Url.ToString());
                    }
                    //else if (e.KeyCode.ToString().Equals("mplus") || ctrl && e.KeyCode.ToString().Equals("D9"))
                    ////////?????  Add ???? ///??? volume increasing
                    //{
                    //    speaker.volume += 5;
                    //}
                    //else if (e.KeyCode.ToString().Equals("mminus") || ctrl && e.KeyCode.ToString().Equals("D8"))    //////?????  Subtract ???? ///??? volume decreasing
                    //{
                    //    speaker.volume -= 5;

                    //}
                    //else if (e.KeyCode.ToString().Equals("Multiply") || ctrl && e.KeyCode.ToString().Equals("D7"))   //////?????  Multiply ???? ///??? speaker speed increasing
                    //{
                    //    speaker.speed++;
                    //}
                    //else if (e.KeyCode.ToString().Equals("Divide") || ctrl && e.KeyCode.ToString().Equals("D6"))  //////?????  Divide ???? ///??? speaker speed decreasing
                    //{
                    //    speaker.speed--;
                    //}
                    //else if (e.KeyCode.ToString().Equals("Space"))  ////????? space ???? and ???? ///??? voice Pause & Resume ???.
                    //{
                    //    speaker.pauseAndResume();
                    //}
                    else if (ctrl && e.KeyCode.ToString().Equals("A")) //////?????  X ????
                    {

                    }
                    //else if (ctrl && e.KeyCode.ToString().Equals("S")) //////?????  V ????
                    //{
                    //    socket = new Soket();
                    //    if (tempo == 44100) tempo = 56000;
                    //    else tempo =44100;
                    //    socket.mmm(tempo.ToString());
                    //}

                    //else if (e.KeyCode.ToString().Equals("F1"))  //////?????  numpad0 ????
                    //{
                    //    speaker.stop();
                    //    RichTextEditor.Form1 AA = new RichTextEditor.Form1(speaker);
                    //    if (thread1 != null) { thread1.Abort(); thread1 = null; }
                    //    thread1 = new Thread(new ThreadStart(AA.AsstForBroser));
                    //    thread1.Start();
                    //    /*
                    //    speaker.speak("press control and 1 for opening a new microsoft word document");
                    //    speaker.speak("press control and 2 for opening a new text document");
                    //    speaker.speak("press control and 3 for sending or receiving mail");
                    //    speaker.speak("press Add or press control Digit9 for volume increase");
                    //    speaker.speak("press Subtract or press control Digit8 for volume decrease");
                    //    speaker.speak("press Multiply or press control Digit7 for speed up");
                    //    speaker.speak("press Divide or press control Digit6 for speed down ");
                    //    speaker.speak("press space bar for pause and resume");
                    //     */
                    //}
                    if (e.KeyData.ToString() == "Divide")
                    {
                        speaker.speak("Divide");
                    }
                    else if (e.KeyData.ToString() == "Multiply")
                    {
                        speaker.speak("Multiply");
                    }
                    else if (e.KeyData.ToString() == "Add")
                    {
                        speaker.speak("Add");
                    }
                    else if (e.KeyData.ToString() == "Subtrack")
                    {
                        speaker.speak("Subtrack");
                    }
                }
                else
                {
                    //if(e.KeyData.ToString()=="Return")
                    //    MessageBox.Show("u press enter");
                }
                
            }

            catch(Exception ex) 
            { 
                //MessageBox.Show("Keydown exception of first try"); 
            }
            /****************************************************************************/

            if (e.KeyCode.ToString().Equals("LControlKey") || e.KeyCode.ToString().Equals("RControlKey"))
            {
                ctrl = true;
            }
            else if (e.KeyCode.ToString().Equals("Insert")) ///////////////////////////////////////
            {
                insrt = true;
            }

            else if (e.KeyCode.ToString().Equals("LMenu") || e.KeyCode.ToString().Equals("RMenu"))
            {
                releaseAlt = true;
                alter = true;
                //NarratorRunOrNotCheck();   
            }

        }
        //browser selected hotkey closed

        private void HookManager_KeyUp(object sender, KeyEventArgs e) ///// HookManager ?? ???? ??????
        {            
            int i = 0;
            if (e.KeyData.ToString() == "Return")
            {                              
                //Process[] procs = Process.GetProcessesByName("WINWORD");
                //if (procs.Length != 0)
                //{
                //    DocForm.DocForm docForm = new DocForm.DocForm();
                //}
            }
            
            else if (e.KeyCode.ToString().Equals("Tab"))
            {
                if (releaseAlt) releaseAlt = false;                
            }
            else if (e.KeyCode.ToString().Equals("LControlKey") || e.KeyCode.ToString().Equals("RControlKey"))
            {
                ctrl = false;
                //speaker.speak("Control Key");

            }
            else if (e.KeyCode.ToString().Equals("Insert")) //////////////////////////////////////////////
            {
                insrt = false;
            }

            else if (e.KeyCode.ToString().Equals("LMenu") || e.KeyCode.ToString().Equals("RMenu"))
            {                
                releaseAlt = false;
                //speaker.speak("Alter Key");
                alter = false;
            }
            
            else if (insrt && e.KeyCode.Equals(Keys.F12))  // say system time insert+f12
            {
                speaker.stop();
                String dat = "Todays Date" + DateTime.Now.ToString("d");
                speaker.speak(dat);
                String time = "  and  Current Time " + DateTime.Now.ToString("T");
                speaker.speak(time);
            }
            else if (insrt && e.KeyCode.Equals(Keys.F4))  // say system time insert+f12
            {
                speaker.speak("Do you Really Want to Shut Down Mongol Dip");
                //System.Threading.Thread.Sleep(3000);
                DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "Do you Really Want to Shut Down Mongol Dip", "Shut Down Mongol Dip", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    Process[] procs = Process.GetProcessesByName("Thesis.vshost");  
                    foreach (Process proc in procs)
                        proc.Kill();

                }
                else if (dialogResult == DialogResult.No)
                {
                    return;
                }
            }

            else if (ctrl && e.KeyCode.ToString().Equals("T"))
            //////?????  numpad2 ???? ///??? new text editor opening
            {
                speaker.speak("You open a new text editor");
                //System.Threading.Thread.Sleep(3000);
                frmMain notepad = new frmMain(speaker);
                notepad.Show();
            }
            else if (ctrl && e.KeyCode.ToString().Equals("M") && !communication.GetActiveProcess().ToString().Equals("WINWORD") && !communication.GetActiveProcess().ToString().Equals("EXCEL") && !communication.GetActiveProcess().ToString().Equals("POWERPNT"))
            //////?????  numpad3 ???? ///??? mailing...but problem ase
            {
                speaker.speak("You open Mail Window");
                //System.Threading.Thread.Sleep(2000);
                sendEmail.Form1 mailAgent = new sendEmail.Form1(speaker);
                mailAgent.Show();
                mailAgent.Activate();
            }

            else if (e.KeyCode.ToString().Equals("mplus") || alter && ctrl && e.KeyCode.ToString().Equals("D9") &&  !communication.GetActiveProcess().ToString().Equals("WINWORD") && !communication.GetActiveProcess().ToString().Equals("EXCEL") && !communication.GetActiveProcess().ToString().Equals("POWERPNT"))
            //////?????  Add ???? ///??? volume increasing
            {
                speaker.volume += 5;
            }
            else if (e.KeyCode.ToString().Equals("mminus") || alter && ctrl && e.KeyCode.ToString().Equals("D8") && !communication.GetActiveProcess().ToString().Equals("WINWORD") && !communication.GetActiveProcess().ToString().Equals("EXCEL") && !communication.GetActiveProcess().ToString().Equals("POWERPNT"))    //////?????  Subtract ???? ///??? volume decreasing
            {
                speaker.volume -= 5;

            }
            else if (e.KeyCode.ToString().Equals("Multiply") || alter && ctrl && e.KeyCode.ToString().Equals("PageUp"))   //////?????  Multiply ???? ///??? speaker speed increasing
            {
                speaker.speed++;
                speaker.speak("Faster");
            }
            else if (e.KeyCode.ToString().Equals("Divide") || alter && ctrl && e.KeyCode.ToString().Equals("Next"))  //////?????  Divide ???? ///??? speaker speed decreasing
            {
                speaker.speed--;
                speaker.speak("Slower");
            }
            else if (e.KeyCode.ToString().Equals("Space") && !communication.getWindowName().ToString().Equals("Text Editor") && !communication.GetActiveProcess().ToString().Equals("WINWORD") && !communication.GetActiveProcess().ToString().Equals("EXCEL") && !communication.GetActiveProcess().ToString().Equals("POWERPNT"))  ////????? space ???? and ???? ///??? voice Pause & Resume ???.
            {
                speaker.pauseAndResume();
            }

            else if (insrt && e.KeyCode.ToString().Equals("Escape"))
            {
                speaker.speed = -1;
                speaker.volume = 90;
                speaker.speak("Screen Refreshed");
            }

            else if (ctrl && e.KeyCode.ToString().Equals("K") && !communication.GetActiveProcess().ToString().Equals("WINWORD") && !communication.GetActiveProcess().ToString().Equals("EXCEL") && !communication.GetActiveProcess().ToString().Equals("POWERPNT")) //////?????  V ????
            {
                socket = new Soket();
                if (tempo == 44100) tempo = 56000;
                else tempo = 44100;
                socket.mmm(tempo.ToString());
            }


            try
            {
                if (communication.getWindowName().ToString().Equals("Text Editor"))
                {
                    Char code = (char)e.KeyCode;

                    if ((code >= 'A' && code <= 'Z') || (code >= 'a' && code <= 'z'))
                    {
                        speaker.speak(e.KeyData.ToString());
                    }
                }
                else if (!communication.GetActiveProcess().ToString().Equals("WINWORD") && !communication.GetActiveProcess().Equals("EXCEL") && !communication.GetActiveProcess().Equals("POWERPNT"))
                {
                    if (!isControllingVolume(e.KeyCode.ToString()))
                    {
                        if (e.KeyData.ToString().Equals("Capital"))
                        {
                            if (Control.IsKeyLocked(Keys.CapsLock))
                            {
                                speaker.speak("caps lock turns off");
                            }
                            else
                                speaker.speak("caps lock turns on");
                        }
                        else if (!e.KeyCode.ToString().Equals("LMenu") && !e.KeyCode.ToString().Equals("RMenu") && !e.KeyCode.ToString().Equals("LControlKey") && !e.KeyCode.ToString().Equals("RControlKey"))
                        {
                            speaker.speak(e.KeyData.ToString());
                            //if (!e.KeyData.ToString().Equals("Tab"))
                            //System.Threading.Thread.Sleep(250);
                        }
                    }
                }
            }
            catch (Exception ex) { }           

            //try
            //{
            //    if (communication.isClientRegion())
            //    {
            //        if (!communication.GetActiveProcess().Equals("NOTEPAD")
            //            && !communication.GetActiveProcess().Equals("WINWORD")
            //            && !communication.GetActiveProcess().Equals("EXCEL")
            //            && !communication.GetActiveProcess().Equals("POWERPNT")
            //            && !communication.getWindowName().Equals("Text Editor"))
            //        {
            //            //speaker.speak(communication.getSelecedText().ToString());
            //            //Console.WriteLine("ffffffffffffffffffff "+communication.getSelecedText().ToString());
            //        }

            //        //TextCaptureLib.CTextCapture TcServer;
            //        //TextCaptureLib.CInputWindow iWindow;

            //        //TcServer = new TextCaptureLib.CTextCapture();
            //        //iWindow = new TextCaptureLib.CInputWindow();
            //        //iWindow.WindowFromPos(0, 0);
            //        //speaker.speak();
            //        //MessageBox.Show(TcServer.GetText(iWindow));

            //    }
            //    else
            //    {
            //        //MessageBox.Show(img.getIconText( graphics ));
            //        //speaker.speak(bitmapImage.getIconText());

            //    }
            //    //speaker.speak(bitmapImage.getIconText());
            //}
            //catch (Exception)
            //{
            //    //MessageBox.Show("Keyup exception");
            //}            
        }


        private void cutAndPaste()
        {
            String sourceDirectory = sourcePath + sourceItem + "//";
            String sourceFile = sourcePath + sourceItem;

            if (Directory.Exists(sourceDirectory))
            {
                Directory.Move(sourceDirectory, destinationPath);
            }
            else if (File.Exists(sourceFile))
            {
                File.Move(sourceFile, destinationPath);
            }
            return;
        }

        //private void CheckActiveWindow_Tick(object sender, EventArgs e)
        //{
        //    String tempWindow = communication.getWindowName();

        //    if (tempWindow != activeWindow)
        //    {
        //        activeWindow = tempWindow;
        //        speaker.stop();
        //        Process[] procs = Process.GetProcessesByName("nvda");

        //        if (communication.isClientRegion() && procs.Length == 0)
        //        {
        //            speaker.speak(communication.getWindowName());

        //        }
        //        else
        //        {
        //            //speaker.speak("Windows explorer browser");
        //            //communication.BringWindowToTop("Mongol Dip Browser", false);
        //        }
        //    }
        //}

        private Boolean isControllingVolume(String key)
        {
            if (key.Equals("Add") || key.Equals("Subtract")
                || key.Equals("Multiply") || key.Equals("Divide")
                || key.Equals("Space"))
            {
                return true;
            }
            return false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
