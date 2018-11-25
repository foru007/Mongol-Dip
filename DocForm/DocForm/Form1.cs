using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SpeechBuilder;
using Word = Microsoft.Office.Interop.Word;


namespace DocForm
{
    public partial class DocForm : Form
    {
        SpeechControl speaker;
        private String fileName = null;
        private Word.Application word = null;
        int p = 0;

        public DocForm(String fileName, SpeechControl speaker, Word.Application word, int p)
        {
            this.p = p;
            this.word = word;
            this.fileName = fileName;
            //MessageBox.Show(fileName);
            InitializeComponent();

            new HandleADoc(fileName, speaker, word, p);

        }
        public DocForm(SpeechControl speaker)
        {
            InitializeComponent();

        }
        public DocForm(Word.Application word, SpeechControl speaker)
        {
            InitializeComponent();
            this.word = word;
            //this.fileName = fileName;
            //MessageBox.Show(word.Name);
            //speaker = new SpeechControl();
            new HandleADoc(speaker,word);
        }

        private void DocForm_Load(object sender, EventArgs e)
        {
        }

    }
}
