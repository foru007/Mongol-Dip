using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SpeechBuilder;

namespace ExcellForm
{
    public partial class Form1 : Form
    {
        private String fileName = null;
        private Excel.Application excel;
        int p = 0;

        public Form1(String fileName, SpeechControl speaker, Excel.Application excel, int p)
        {
            this.p = p;
            this.excel = excel;
            this.fileName = fileName;
            //MessageBox.Show(fileName);
            InitializeComponent();

            new HandleExcell(fileName, speaker,excel, p);

        }
        public Form1(Excel.Application excel,SpeechControl speaker)
        {
            this.excel = excel;            
            //MessageBox.Show(excel.Name);
            InitializeComponent();
            //Console.WriteLine(excel.Name);
            new HandleExcell(speaker, excel);

        }
        public Form1()
        {
            InitializeComponent();
            //new HandleExcell();
        }
       
    }
}
