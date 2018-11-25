using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Speech;
using System.Speech.Synthesis;
using System.Windows.Forms;
using System.Threading;


namespace SpeechBuilder
{
    /// <summary>
    /// এই ক্লাসটি speaker কে control করার জন্য ব্যবহার করা হয়েছে
    /// volume বাড়ানো-কমানো
    /// আম
    /// পড়ার speed বাড়ানো-কমানো
    /// speaker select করা
    /// pause এবং resume করা
    /// যেকোন মুহূতে reading থামিয়ে দেয্য
    /// </summary>
    public class SpeechControl
    {
        public static int TT = 0;
        public String temp = null;
        public Boolean STOP = false;
        SpeechSynthesizer speaker;
        Soket soket;
        char[] sp = new char[2];

        String Bang = "";

        //constructor
        //***************************************************//
        public SpeechControl()
        {
            soket = new Soket();
            //this.soket = soket;
            speaker = new SpeechSynthesizer();
            speaker.Rate = -1;
            speaker.Volume = 90;
        }
        //***************************************************//

        /// <summary>
        /// for speaker selection
        /// </summary>
        /// <param name="speakerName"></param>
        //***************************************************//
        public void selectSpeaker(String speakerName)
        {
            speaker.SelectVoice("Microsoft Mike");
        }
        //***************************************************//
        /// <summary>
        /// যেকোন moment এ pause and resume
        /// </summary>
        /// 
        //***************************************************//
        public void pauseAndResume()
        {
            if (speaker.State == SynthesizerState.Speaking)
            {
                speaker.Pause();
            }
            else if (speaker.State == SynthesizerState.Paused)
            {
                speaker.Resume();
            }
        }
        public void Pause()
        {
            speaker.Pause();

        }
        public void Resume()
        {
            speaker.Resume();
        }
        public void Test(object sender, SpeakCompletedEventArgs e)
        {
            //System.Threading.Thread.Sleep(500);
            TT = 0;
            //MessageBox.Show("dddddddddd");
        }
        //***************************************************//
        /// <summary>
        /// যে word তি পড়তে হবে
        /// </summary>
        /// <param name="word"></param>
        //***************************************************//

        public void speak(String word)
        {
            //Console.WriteLine("Word::=" + word);
            try
            {
                STOP = false;
                String EngS = "";
                String BangS = "";

                if (word == null || word.Equals(""))
                    return;

                else if (word.Length >= 1)
                {
                    int c = 0;
                    String Nword = word + ' ' + "Null";
                    Nword.Trim();
                    String[] ss = Nword.Split(' ');
                    foreach (String s in ss)
                    {
                        if (s.Trim() != "")
                        {
                            //Console.WriteLine("Split::=" + s);
                            if (s.Trim() != "Null")
                            {
                                char[] wc = s.ToCharArray();
                                //Console.WriteLine(wc[0]);
                                if (wc[0] >= 0 && wc[0] <= 127)
                                {
                                    EngS += s + ' ';
                                    //Console.WriteLine("Read ENG::=" + EngS);
                                    if (BangS != "")
                                    {
                                        Bang = BangS;
                                        BangS.Trim();
                                        //Console.WriteLine("Read Bangla::=" + BangS);
                                        //soket.mm("1");
                                        temp = soket.mm(BangS);                                                                                
                                        Console.WriteLine(temp);
                                        if (temp=="STOPPED") {break;}                                        
                                        //String nn = BangS.Length.ToString();
                                        //int n = int.Parse(nn) * 185;
                                        //System.Threading.Thread.Sleep(n);
                                        BangS = "";
                                    }
                                    c = 1;
                                }
                                else
                                {
                                    BangS += s + ' ';
                                    Bang += s + ' ';
                                    //Console.WriteLine("BNG::=" + BangS);
                                    if (EngS != "")
                                    {
                                        TT = 100000;
                                        EngS.Trim();
                                        //Console.WriteLine("Read ENG::=" + EngS);                                        
                                        speaker.SpeakAsync(EngS);
                                        speaker.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(Test);                                                                                                          
                                        int i = 0;
                                        while (TT != 0)
                                        {
                                            i++;
                                        }
                                        if (STOP) break;
                                        //String ee = EngS.Length.ToString();
                                        //int T = int.Parse(ee) * 73;                                        
                                        //System.Threading.Thread.Sleep(TT);
                                        EngS = "";
                                        //Console.WriteLine("TEST");
                                    }
                                    c = 2;
                                    //BangN = "the nak tkek ndfkk";
                                    //Console.WriteLine(BangS);
                                }
                            }
                            else
                            {
                                //Console.WriteLine("Found Null");
                                //Console.WriteLine("Read ENG::=" + EngS);                                
                                if (c == 1 && EngS != "")
                                {
                                    speaker.SpeakAsync(EngS);
                                    //speaker.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(Test);
                                }
                                else if (c == 2 && BangS != "")
                                {
                                    Bang = BangS;
                                    //Console.WriteLine("Found Null:: Read Bangla"); 
                                    //soket.mm("1");
                                    soket.mm(BangS);                                    
                                }
                            }
                        }

                        //soket.mm(BangS);
                    }                    

                }
                else speaker.SpeakAsync(word);
            }
            catch (Exception ex) { }
        }
        //***************************************************//

        /// <summary>
        /// যেকোন moment এ stop করিয়ে দেয়া
        /// </summary>
        //***************************************************//
        public void stop()
        {
            if (Bang != "")
            {
                soket.mmm("stop");
                Bang = "";
            }
            
            if (speaker.State == SynthesizerState.Speaking)
            {
                speaker.SpeakAsyncCancelAll();
                STOP = true;
            }

        }
        //***************************************************//
        /// <summary>
        /// volume বাড়ানো বা কমানো
        /// </summary>
        //***************************************************//
        public int volume
        {
            get
            {
                return speaker.Volume;
            }
            set
            {
                if (value < 1)
                {
                    speaker.Volume = 1;
                }
                else if (value > 98)
                {
                    speaker.Volume = 99;
                }
                else
                {
                    speaker.Volume = value;
                }

            }
        }
        //***************************************************//
        /// <summary>
        /// speed বাড়ানো বা কমানো
        /// </summary>
        //***************************************************//
        public int speed
        {
            get
            {
                return speaker.Rate;
            }
            set
            {
                //speaker.Rate = value;

                if (value < -5)
                {
                    speaker.Rate = -5;
                }
                else if (value > 9)
                {
                    speaker.Rate = 9;
                }
                else
                {
                    speaker.Rate = value;
                }
            }
        }
        //***************************************************//
    }
    //End of class
}
