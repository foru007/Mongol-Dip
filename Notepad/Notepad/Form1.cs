using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SpeechBuilder;
using System.Threading;

namespace RichTextEditor
{
    public partial class Form1 : Form
    {
       private SpeechControl speaker;
        public Form1(SpeechControl speaker)
        {
           InitializeComponent();
           this.speaker = speaker;
        }
        // 30/10/2011 :: Method for Give Bangla Audio Assistance
        public void bangla()
        {
           //Monitor.Enter(this);
            String ss = "Instruction of how to operate  Notepad & PDF Document press Ctrl + Numpad 0 or Ctrl + D0. নোটপ্যাড এবং পিডিএফ চালানোর নির্দেশিকা জানতে কন্ট্রোল এবং নামপ্যাড ০ অথবা  কন্ট্রোল এবং ডি ০ বাটন চাপ দিন। To Listen One Character From Left Press  Left Arrow. বাম পাশের একটি অক্ষর পরার জন্য বাম তীর চাপ দিন।  To Listen One Character From Right Press  Right Arrow. ডান পাশের একটি অক্ষর পরার জন্য ডান তীর চাপ দিন।  To Listen One Word From left Press  Ctrl + Right Arrow. ডান পাশের একটি শব্দ পরার জন্য কনট্রোল এবং ডান তীর চাপ দিন। To Listen One Word From Right Press  Ctrl + Left Arrow. বাম পাশের একটি শব্দ পরার জন্য কনট্রোল এবং বাম তীর চাপ দিন। To Listen Full Text Press Ctrl + End. সম্পূর্ণ লেখা পরার জন্য কনট্রোল এবং এনড চাপ দিন। To Save document Press Altr + F4. ডকুমেন্ট  সেইভ করতে অলটার এবং এফ ফউর চাপ দিন।";
            speaker.speak(ss);
           //speaker.stop();
           //Monitor.Pulse(this);
           //Monitor.Exit(this);

        }
        public void AsstForBroser()
        {
            speaker.stop();
            String ss = "Instruction of how to operate  Mongol Dip Browser press F1 key. মঙ্গলদ্বীপ চালানোর নির্দেশিকা জানতে এফ ওয়ান বাটন চাপ দিন। To know any keyboard key name Press the desire key. কোনো  কি এর নাম জানতে ওই বাটন চাপ দিন। To Select File/Folder from current link. Press Right/left /up/down key. বর্তমান অবস্থান থেকে  ফাইল ফোল্ডার এর নাম জানতে ডান, বাম, উপর, নিচ এর বাটন চাপ দিন। To Open/Execute/Enter current selected File/Folder Press Enter. কোনো ফাইল, ফোল্ডার খুলতে এন্টার বাটন চাপ দিন। To Browse Hard Drive such as My computer, My document, Desktop, Music Press Tab. কোনো ড্রাইভে যেতে ট্যাব বাটন চাপ দিন। To Terminate Any Opening Application Press Altr + F4. কোনো খোলা ফাইল বন্ধ করতে অলটার  এবং  এফ ফউর বাটন চাপ দিন। To Hear Today’s Date & Time Press  Insert + F12. আজকের দিন তারিখ  এবং বর্তমান সময় জানার জন্য ইনসার্ট এবং  এফ ১২  চাপ দিন। To Shut-Down Computer Press  Insert + F4. কম্পিউটার  বন্ধ করার জন্য ইনসার্ট এবং  এফ ফউর  চাপ দিন। To Open New Microsoft Word Document Press Ctrl + W. নতুন ওয়ার্ড ডকুমেন্ট খোলার জন্য কনট্রোল এবং  ডবলিও চাপ দিন। To Open New Text Editor Document Press Ctrl + T. নতুন টেক্সট ডকুমেন্ট খোলার জন্য কনট্রোল এবং টি চাপ দিন। To Open New Excel Document Press Ctrl + E. নতুন এক্সেল ডকুমেন্ট খোলার জন্য কনট্রোল এবং ই চাপ দিন। To Open New PowerPoint Document Press Ctrl + P. নতুন পাওয়ার পয়েন্ট ডকুমেন্ট খোলার জন্য কনট্রোল এবং পি চাপ দিন। To Open E-mail Sending & Receiving Window Press Ctrl + M. মেইল পাঠনো এবং গ্রহণ করার জন্য কনট্রোল এবং এম চাপ দিন To open Task Reminder Press Ctrl + R. টাস্ক রিমাইনডার খোলার জন্য কনট্রোল এবং আর চাপ দিন। Copy File/Folder Press  Ctrl + C. ফাইল কপি করতে কনট্রোল এবং সি চাপ দিন। Paste File/Folder Press  Ctrl + V. কপি করা ফাইল রাখতে কনট্রোল এবং সি চাপ দিন। Edit File Folder Name From Mongol Dip Browser Press F2. ফাইল ফোল্ডার এর নাম পরিবর্তন করতে এফ টু চাপ দিন। Delete File/Folder Press  Delete. কোনো ফাইল ফোল্ডার মুছে ফেলতে ডিলিট বাটন চাপ দিন। To Increase Speaker Volume Press  Ctrl + D9 or  Press +. স্পিকারের শব্দ বাড়াতে কন্ট্রোল এবং ডি ৯ অথবা কন্ট্রোল এবং প্লাস বাটন চাপ দিন। To Decrease Speaker Volume Press  Ctrl + D8 or  Press -. স্পিকারের শব্দ বাড়াতে কন্ট্রোল এবং ডি ৮ অথবা কন্ট্রোল এবং মাইনাস বাটন চাপ দিন। To Increase Speech Speed Press  Ctrl + D7 or  Press *. তাড়াতাড়ি পড়তে কন্ট্রোল এবং ডি ৭ অথবা কন্ট্রোল এবং স্টার বাটন চাপ দিন। To Decrease Speech Speed Press  Ctrl + D6  or  Press /. আস্তে পড়তে কন্ট্রোল এবং ডি ৬ অথবা কন্ট্রোল এবং ডিভাইড বাটন চাপ দিন। To Add a new Folder Press  Ctrl + F. নতুন ফোল্ডার তৈরি করতে কন্ট্রোল এবং এফ চাপ দিন। To on-off Speaker Sound Press Space key. স্পিকার সাউন্ড বন্ধ করতে ও খুলতে স্পেস চাপ দিন।";
            speaker.speak(ss);

        }
        //End :: Method Give Bangla Audio Assistance     
       
    }
}
