using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PPT = Microsoft.Office.Interop.PowerPoint;
using SpeechBuilder;

namespace PPTForm
{
    public partial class Form1 : Form
    {
        private String fileName = null;
        private PPT.Application ppt=null;
        int p = 0;

        public Form1(String fileName, SpeechControl speaker, PPT.Application ppt, int p)
        {
            this.p = p;
            ppt = ppt;
            this.fileName = fileName;
            //MessageBox.Show(fileName);
            InitializeComponent();

            new HandlePPT(fileName, speaker, ppt, p);

        }
        public Form1(PPT.Application ppt, SpeechControl speaker)
        {
            this.ppt = ppt;
           
            InitializeComponent();
            //MessageBox.Show("on pptForm"+ppt.Name);
            new HandlePPT(speaker, ppt);

        }
        public Form1()
        {
            InitializeComponent();
        }
    }
}
