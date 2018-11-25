using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SpeechBuilder;

namespace RichTextEditor
{
    public partial class Form2 : Form
    {
        private SpeechControl speaker;
        public Form2(SpeechControl speaker)
        {
            InitializeComponent();
            this.speaker = speaker;
        }
    }
}
