using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Windows.Forms;
using ExtendedRichTextBox;
using SpeechBuilder;
using System.IO;
using System.Threading;
using System.Reflection;
using Word;
using System.Diagnostics;

namespace RichTextEditor
{

    public partial class frmMain : System.Windows.Forms.Form
    {
        //String ss=null;
        //String ee= null;
        String FT = null;
        //int T=0;
        private SpeechControl speaker;
        //private Form1 f1;
        private Boolean pressCtrl = false;
        private Boolean pressShift = false;
        Thread thread = null;
        string fN;

        //ExtendedRichTextBox.RichTextBoxPrintCtrl rtbDoc;
        // constructor
        public frmMain(SpeechControl speaker)
        {
            fN = null;
            if (thread != null) { thread.Abort(); thread = null; }
            InitializeComponent();
            currentFile = "";
            this.Text = "Text Editor";
            this.speaker = speaker;
            //loadAFile(@"C:\Documents and Settings\Alamgir\Desktop\refrences.txt");

        }
      
        #region "Declaration"

        private string currentFile;
        private int checkPrint;

        #endregion



        #region "Menu Methods"

        public void ReadFullText()
        {

            speaker.speak(FT);
            //speaker.stop();
            //if (thread != null) { thread.Abort(); }
        }
        public void ReadSelectedText()
        {

            //speaker.speak(rtbDoc.SelectedText.ToString());
            //speaker.stop();
            //if (thread != null) { thread.Abort(); }
        }
        private void NewToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (rtbDoc.Modified == true)
                {
                    System.Windows.Forms.DialogResult answer;
                    answer = MessageBox.Show("Save current document before creating new document?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer == System.Windows.Forms.DialogResult.No)
                    {
                        currentFile = "";
                        this.Text = "Editor: New Document";
                        rtbDoc.Modified = false;
                        rtbDoc.Clear();
                        return;
                    }
                    else
                    {
                        SaveToolStripMenuItem_Click(this, new EventArgs());
                        rtbDoc.Modified = false;
                        rtbDoc.Clear();
                        currentFile = "";
                        this.Text = "Editor: New Document";
                        return;
                    }
                }
                else
                {
                    currentFile = "";
                    this.Text = "Editor: New Document";
                    rtbDoc.Modified = false;
                    rtbDoc.Clear();
                    return;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void OpenToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (rtbDoc.Modified == true)
                {
                    System.Windows.Forms.DialogResult answer;
                    answer = MessageBox.Show("Save current file before opening another document?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer == System.Windows.Forms.DialogResult.No)
                    {
                        rtbDoc.Modified = false;
                        OpenFile();
                    }
                    else
                    {
                        SaveToolStripMenuItem_Click(this, new EventArgs());
                        OpenFile();
                    }
                }
                else
                {
                    OpenFile();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void OpenFile()
        {
            try
            {
                OpenFileDialog1.Title = "RTE - Open File";
                OpenFileDialog1.DefaultExt = "rtf";
                OpenFileDialog1.Filter = "Text Files|*.txt|Rich Text Files|*.rtf|HTML Files|*.htm|All Files|*.*";
                OpenFileDialog1.FilterIndex = 1;
                OpenFileDialog1.FileName = string.Empty;

                if (OpenFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    if (OpenFileDialog1.FileName == "")
                    {
                        return;
                    }

                    string strExt;
                    strExt = System.IO.Path.GetExtension(OpenFileDialog1.FileName);
                    strExt = strExt.ToUpper();

                    if (strExt == ".RTF")
                    {
                        rtbDoc.LoadFile(OpenFileDialog1.FileName, RichTextBoxStreamType.RichText);
                    }
                    else
                    {
                        System.IO.StreamReader txtReader;
                        txtReader = new System.IO.StreamReader(OpenFileDialog1.FileName);
                        rtbDoc.Text = txtReader.ReadToEnd();
                        txtReader.Close();
                        txtReader = null;
                        rtbDoc.SelectionStart = 0;
                        rtbDoc.SelectionLength = 0;
                    }

                    currentFile = OpenFileDialog1.FileName;
                    rtbDoc.Modified = false;
                    this.Text = "Editor: " + currentFile.ToString();
                }
                else
                {
                    MessageBox.Show("Open File request cancelled by user.", "Cancelled");
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void SaveToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (currentFile == string.Empty)
                {
                    SaveAsToolStripMenuItem_Click(this, e);
                    return;
                }

                try
                {
                    string strExt;
                    strExt = System.IO.Path.GetExtension(currentFile);
                    strExt = strExt.ToUpper();
                    if (strExt == ".RTF")
                    {
                        rtbDoc.SaveFile(currentFile);
                    }
                    else
                    {
                        System.IO.StreamWriter txtWriter;
                        txtWriter = new System.IO.StreamWriter(currentFile);
                        txtWriter.Write(rtbDoc.Text);
                        txtWriter.Close();
                        txtWriter = null;
                        rtbDoc.SelectionStart = 0;
                        rtbDoc.SelectionLength = 0;
                    }
                    this.Text = "Editor: " + currentFile.ToString();
                    rtbDoc.Modified = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), "File Save Error");
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }


        }


        private void SaveAsToolStripMenuItem_Click(object sender, System.EventArgs e)
        {

            try
            {
                SaveFileDialog1.Title = "RTE - Save File";
                SaveFileDialog1.DefaultExt = "rtf";
                SaveFileDialog1.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*";
                SaveFileDialog1.FilterIndex = 1;

                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    if (SaveFileDialog1.FileName == "")
                    {
                        return;
                    }

                    string strExt;
                    strExt = System.IO.Path.GetExtension(SaveFileDialog1.FileName);
                    strExt = strExt.ToUpper();

                    if (strExt == ".RTF")
                    {
                        rtbDoc.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.RichText);
                    }
                    else
                    {
                        System.IO.StreamWriter txtWriter;
                        txtWriter = new System.IO.StreamWriter(SaveFileDialog1.FileName);
                        txtWriter.Write(rtbDoc.Text);
                        txtWriter.Close();
                        txtWriter = null;
                        rtbDoc.SelectionStart = 0;
                        rtbDoc.SelectionLength = 0;
                    }

                    currentFile = SaveFileDialog1.FileName;
                    rtbDoc.Modified = false;
                    this.Text = "Editor: " + currentFile.ToString();
                    MessageBox.Show(currentFile.ToString() + " saved.", "File Save");
                }
                else
                {
                    MessageBox.Show("Save File request cancelled by user.", "Cancelled");
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void ExitToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (rtbDoc.Modified == true)
                {
                    System.Windows.Forms.DialogResult answer;
                    answer = MessageBox.Show("Save this document before closing?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer == System.Windows.Forms.DialogResult.Yes)
                    {
                        return;
                    }
                    else
                    {
                        rtbDoc.Modified = false;
                        System.Windows.Forms.Application.Exit();
                    }
                }
                else
                {
                    rtbDoc.Modified = false;
                    System.Windows.Forms.Application.Exit();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void SelectAllToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.SelectAll();
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to select all document content.", "RTE - Select", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private void CopyToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.Copy();
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to copy document content.", "RTE - Copy", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private void CutToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.Cut();
            }
            catch
            {
                MessageBox.Show("Unable to cut document content.", "RTE - Cut", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private void PasteToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.Paste();
            }
            catch
            {
                MessageBox.Show("Unable to copy clipboard content to document.", "RTE - Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private void SelectFontToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                {
                    FontDialog1.Font = rtbDoc.SelectionFont;
                }
                else
                {
                    FontDialog1.Font = null;
                }
                FontDialog1.ShowApply = true;
                if (FontDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    rtbDoc.SelectionFont = FontDialog1.Font;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void FontColorToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                ColorDialog1.Color = rtbDoc.ForeColor;
                if (ColorDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    rtbDoc.SelectionColor = ColorDialog1.Color;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void BoldToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                {
                    System.Drawing.Font currentFont = rtbDoc.SelectionFont;
                    System.Drawing.FontStyle newFontStyle;

                    newFontStyle = rtbDoc.SelectionFont.Style ^ FontStyle.Bold;

                    rtbDoc.SelectionFont = new System.Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle);

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void ItalicToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                {
                    System.Drawing.Font currentFont = rtbDoc.SelectionFont;
                    System.Drawing.FontStyle newFontStyle;

                    newFontStyle = rtbDoc.SelectionFont.Style ^ FontStyle.Italic;

                    rtbDoc.SelectionFont = new System.Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }





        private void UnderlineToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                {
                    System.Drawing.Font currentFont = rtbDoc.SelectionFont;
                    System.Drawing.FontStyle newFontStyle;

                    newFontStyle = rtbDoc.SelectionFont.Style ^ FontStyle.Underline;

                    rtbDoc.SelectionFont = new System.Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }





        private void NormalToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                {
                    System.Drawing.Font currentFont = rtbDoc.SelectionFont;
                    System.Drawing.FontStyle newFontStyle;
                    newFontStyle = FontStyle.Regular;
                    rtbDoc.SelectionFont = new System.Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void PageColorToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                ColorDialog1.Color = rtbDoc.BackColor;
                if (ColorDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    rtbDoc.BackColor = ColorDialog1.Color;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void mnuUndo_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (rtbDoc.CanUndo)
                {
                    rtbDoc.Undo();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void mnuRedo_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (rtbDoc.CanRedo)
                {
                    rtbDoc.Redo();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void LeftToolStripMenuItem_Click_1(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.SelectionAlignment = HorizontalAlignment.Left;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void CenterToolStripMenuItem_Click_1(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.SelectionAlignment = HorizontalAlignment.Center;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void RightToolStripMenuItem_Click_1(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.SelectionAlignment = HorizontalAlignment.Right;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void AddBulletsToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.BulletIndent = 10;
                rtbDoc.SelectionBullet = true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void RemoveBulletsToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.SelectionBullet = false;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void mnuIndent0_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.SelectionIndent = 0;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void mnuIndent5_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.SelectionIndent = 5;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void mnuIndent10_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.SelectionIndent = 10;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void mnuIndent15_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.SelectionIndent = 15;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void mnuIndent20_Click(object sender, System.EventArgs e)
        {
            try
            {
                rtbDoc.SelectionIndent = 20;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }

        }




        private void FindToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                frmFind f = new frmFind(this);
                f.Show();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void FindAndReplaceToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                frmReplace f = new frmReplace(this);
                f.Show();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void PreviewToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                PrintPreviewDialog1.Document = PrintDocument1;
                PrintPreviewDialog1.ShowDialog();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void PrintToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                PrintDialog1.Document = PrintDocument1;
                if (PrintDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    PrintDocument1.Print();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void mnuPageSetup_Click(object sender, System.EventArgs e)
        {
            try
            {
                PageSetupDialog1.Document = PrintDocument1;
                PageSetupDialog1.ShowDialog();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void InsertImageToolStripMenuItem_Click(object sender, System.EventArgs e)
        {

            OpenFileDialog1.Title = "RTE - Insert Image File";
            OpenFileDialog1.DefaultExt = "rtf";
            OpenFileDialog1.Filter = "Bitmap Files|*.bmp|JPEG Files|*.jpg|GIF Files|*.gif";
            OpenFileDialog1.FilterIndex = 1;
            OpenFileDialog1.ShowDialog();

            if (OpenFileDialog1.FileName == "")
            {
                return;
            }

            try
            {
                string strImagePath = OpenFileDialog1.FileName;
                Image img;
                img = Image.FromFile(strImagePath);
                Clipboard.SetDataObject(img);
                DataFormats.Format df;
                df = DataFormats.GetFormat(DataFormats.Bitmap);
                if (this.rtbDoc.CanPaste(df))
                {
                    this.rtbDoc.Paste(df);
                }
            }
            catch
            {
                MessageBox.Show("Unable to insert image format selected.", "RTE - Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void rtbDoc_SelectionChanged(object sender, EventArgs e)
        {
            //tbrBold.Checked = rtbDoc.SelectionFont.Bold;
            //tbrItalic.Checked = rtbDoc.SelectionFont.Italic;
            //tbrUnderline.Checked = rtbDoc.SelectionFont.Underline;
        }



        #endregion




        #region Toolbar Methods


        private void tbrSave_Click(object sender, System.EventArgs e)
        {
            SaveToolStripMenuItem_Click(this, e);
        }


        private void tbrOpen_Click(object sender, System.EventArgs e)
        {
            OpenToolStripMenuItem_Click(this, e);
        }


        private void tbrNew_Click(object sender, System.EventArgs e)
        {
            NewToolStripMenuItem_Click(this, e);
        }


        private void tbrBold_Click(object sender, System.EventArgs e)
        {
            BoldToolStripMenuItem_Click(this, e);
        }


        private void tbrItalic_Click(object sender, System.EventArgs e)
        {
            ItalicToolStripMenuItem_Click(this, e);
        }


        private void tbrUnderline_Click(object sender, System.EventArgs e)
        {
            UnderlineToolStripMenuItem_Click(this, e);
        }


        private void tbrFont_Click(object sender, System.EventArgs e)
        {
            SelectFontToolStripMenuItem_Click(this, e);
        }


        private void tbrLeft_Click(object sender, System.EventArgs e)
        {
            rtbDoc.SelectionAlignment = HorizontalAlignment.Left;
        }


        private void tbrCenter_Click(object sender, System.EventArgs e)
        {
            rtbDoc.SelectionAlignment = HorizontalAlignment.Center;
        }


        private void tbrRight_Click(object sender, System.EventArgs e)
        {
            rtbDoc.SelectionAlignment = HorizontalAlignment.Right;
        }


        private void tbrFind_Click(object sender, System.EventArgs e)
        {
            frmFind f = new frmFind(this);
            f.Show();
        }


        private void tspColor_Click(object sender, EventArgs e)
        {
            FontColorToolStripMenuItem_Click(this, new EventArgs());
        }




        #endregion




        #region Printing


        private void PrintDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {

            checkPrint = 0;

        }



        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            checkPrint = rtbDoc.Print(checkPrint, rtbDoc.TextLength, e);

            if (checkPrint < rtbDoc.TextLength)
            {
                e.HasMorePages = true;
            }
            else
            {
                e.HasMorePages = false;
            }

        }





        #endregion




        #region Form Closing Handler


        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (rtbDoc.Modified == true)
                {
                    System.Windows.Forms.DialogResult answer;
                    answer = MessageBox.Show("Save current document before exiting?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer == System.Windows.Forms.DialogResult.No)
                    {
                        rtbDoc.Modified = false;
                        rtbDoc.Clear();
                        return;
                    }
                    else
                    {
                        SaveToolStripMenuItem_Click(this, new EventArgs());
                    }
                }
                else
                {
                    rtbDoc.Clear();
                }
                currentFile = "";
                this.Text = "Editor: New Document";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        #endregion

        #region keydow and keypress

        private void rtbDoc_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                //MessageBox.Show(e.KeyCode.ToString());
                if (e.KeyCode.ToString().Equals("ShiftKey"))
                {
                    speaker.speak("Shift");
                    pressShift = true;
                }
                else if (e.KeyCode.ToString().Equals("Tab"))
                {
                    speaker.speak("Tab");
                }
                else if (e.KeyCode.ToString().Equals("Escape"))
                {
                    speaker.speak("Escape");
                }
                else if (e.KeyCode.ToString().Equals("Capital"))
                {
                    speaker.speak("Capital");
                }
                else if (e.KeyCode.ToString().Equals("Insert"))
                {
                    speaker.speak("Insert");
                }
                else if (e.KeyCode.ToString().Equals("Delete"))
                {
                    speaker.speak("Delete");
                }
                else if (e.KeyCode.ToString().Equals("Home"))
                {
                    speaker.speak("Home");
                }
                else if (e.KeyCode.ToString().Equals("PageUp"))
                {
                    speaker.speak("PageUp");
                }
                else if (e.KeyCode.ToString().Equals("Next"))
                {
                    speaker.speak("Next");
                }
                else if (e.KeyCode.ToString().Equals("ControlKey"))
                {
                    speaker.speak("Control Key");
                    pressCtrl = true;
                }
                else if (e.KeyCode.ToString().Equals("Back"))
                {
                    speaker.speak("Back Space");
                }
                else if ((pressCtrl && e.KeyCode.ToString().Equals("t")) || (pressCtrl && e.KeyCode.ToString().Equals("T")))   // say title
                {
                    if (fN == null)
                        speaker.speak("you open a new text document");

                    speaker.speak(fN);
                }

                //if ((pressCtrl && e.KeyCode.ToString().Equals("x")) || (pressCtrl && e.KeyCode.ToString().Equals("X")))   // say title
                //{

                //    MessageBox.Show(rtbDoc.GetLineFromCharIndex(rtbDoc.SelectionStart).ToString());
                //}
                else if (pressCtrl && e.KeyCode.ToString().Equals("Right"))                          //read 1 word
                {
                    if (thread != null) { thread.Abort(); thread = null; }
                    int cursorPosition = rtbDoc.SelectionStart;
                    int nextSpace = rtbDoc.Text.IndexOf(' ', cursorPosition);
                    int selectionStart = 0;
                    string trimmedString = string.Empty;
                    // Strip everything after the next space...
                    if (nextSpace != -1)
                    {
                        trimmedString = rtbDoc.Text.Substring(0, nextSpace);

                    }
                    else
                    {
                        trimmedString = rtbDoc.Text;

                    }

                    if (trimmedString.LastIndexOf(' ') != -1)
                    {
                        selectionStart = 1 + trimmedString.LastIndexOf(' ');
                        trimmedString = trimmedString.Substring(1 + trimmedString.LastIndexOf(' '));

                    }
                    rtbDoc.Select(rtbDoc.SelectionStart, trimmedString.Length);
                    speaker.speak(rtbDoc.SelectedText.ToString());
                    //speaker.stop();
                    //if (thread != null) { thread.Abort(); }

                }

                else if (pressCtrl && pressShift && e.KeyCode.ToString().Equals("Right"))  // read sentence upto fullstop
                {
                    if (thread != null) { thread.Abort(); thread = null; }
                    int cursorPosition = rtbDoc.SelectionStart;
                    int nextSpace = rtbDoc.Text.IndexOf(".", cursorPosition);
                    int selectionStart = 0;
                    string trimmedString = string.Empty;
                    // Strip everything after the next space...
                    if (nextSpace != -1)
                    {
                        trimmedString = rtbDoc.Text.Substring(0, nextSpace);
                    }
                    else
                    {
                        trimmedString = rtbDoc.Text;
                    }
                    if (trimmedString.LastIndexOf(".") != -1)
                    {
                        selectionStart = 1 + trimmedString.LastIndexOf(".");
                        trimmedString = trimmedString.Substring(1 + trimmedString.LastIndexOf("."));
                    }
                    rtbDoc.Select(rtbDoc.SelectionStart, trimmedString.Length);
                    speaker.speak(rtbDoc.SelectedText.ToString());
                    speaker.stop();
                    //if (thread != null) { thread.Abort(); }

                }
                else if (pressCtrl && e.KeyCode.ToString().Equals("Down"))                // read 2nd para to rest para
                {
                    if (thread != null) { thread.Abort(); thread = null; }
                    int cursorPosition = rtbDoc.SelectionStart;
                    int nextSpace = rtbDoc.Text.IndexOf("\n", cursorPosition);
                    int selectionStart = 0;
                    string trimmedString = string.Empty;
                    // Strip everything after the next space...
                    if (nextSpace != -1)
                    {
                        trimmedString = rtbDoc.Text.Substring(0, nextSpace);
                    }
                    else
                    {
                        trimmedString = rtbDoc.Text;
                    }

                    if (trimmedString.LastIndexOf("\n") != -2)
                    {
                        //MessageBox.Show("ffgggfffgg");
                        selectionStart = 1 + trimmedString.LastIndexOf("\n");
                        trimmedString = trimmedString.Substring(1 + trimmedString.LastIndexOf("\n"));
                        rtbDoc.Select(rtbDoc.SelectionStart, trimmedString.Length);
                        speaker.speak(rtbDoc.SelectedText.ToString());

                    }

                }
                else if (pressCtrl && e.KeyCode.ToString().Equals("End"))                      // read full text
                {
                    FT = rtbDoc.Text;
                    if (thread != null) { thread.Abort(); thread = null; }
                    thread = new Thread(new ThreadStart(ReadFullText));
                    thread.Start();
                }

                else if (e.KeyCode.ToString().Equals("End"))
                {
                    speaker.speak("End");
                }

                // 30/10/2011 :: Give Bangla Audio Assistance           
                else if ((pressCtrl && e.KeyCode.ToString().Equals("NumPad0")) || pressCtrl && e.KeyCode.ToString().Equals("D0"))  /////*****("NumPad0"))
                {

                    Form1 ff = new Form1(speaker);

                    //// press control one for reading full text সম্পূর্ণ লেখা পরার জন্য কনট্রোল এবং এক চাপ দিন। press control right arrow for reading A word from right side and left arrow for left side  ডান পাশের একটি শব্ধ পরার জন্য কনট্রোল এবং ডান তীর অথবা বাম পাশের জন্য বাম তীর চাপ দিন।  press left arrow to listen one character from left  বাম পাশের একটি অক্ষর পরার জন্য বাম তীর চাপ দিন।  press right arrow to listen one character from right ডান পাশের একটি অক্ষর পরার জন্য ডান তীর চাপ দিন।
                    // 3/11/2011 :: Solve Bangla Audio Assistance problem1
                    if (thread != null) { thread.Abort(); thread = null; }
                    thread = new Thread(new ThreadStart(ff.bangla));
                    thread.Start();
                    // End :: Solve Bangla Audio Assistance problem1 

                }
                //End :: Give Bangla Audio Assistance
                else if (e.KeyData.ToString().Equals("Right"))  //tir                              // read 1 character
                {
                    if (thread != null) { thread.Abort(); thread = null; }
                    rtbDoc.Select(rtbDoc.SelectionStart, 1); //1 char porba
                    //MessageBox.Show(rtbDoc.SelectedText.ToString());
                    speaker.stop();

                    if (rtbDoc.SelectedText.ToString() == " ")
                        speaker.speak("Space");

                    if (rtbDoc.SelectedText.ToString() == "!")
                        speaker.speak("Exclamation mark");
                    //else if (rtbDoc.SelectedText.ToString() == """)
                    //    speaker.speak("Double quotes");
                    //else if (rtbDoc.SelectedText.ToString() == """)
                    //    speaker.speak("Double quotes");
                    else if (rtbDoc.SelectedText.ToString() == "(")
                        speaker.speak("first bracker Open");
                    else if (rtbDoc.SelectedText.ToString() == ")")
                        speaker.speak("first bracker Close");
                    else if (rtbDoc.SelectedText.ToString() == ",")
                        speaker.speak("Comma");
                    else if (rtbDoc.SelectedText.ToString() == "-")
                        speaker.speak("Hyphen");
                    else if (rtbDoc.SelectedText.ToString() == ".")
                        speaker.speak(" full stop");
                    else if (rtbDoc.SelectedText.ToString() == ":")
                        speaker.speak("Colon");
                    else if (rtbDoc.SelectedText.ToString() == ";")
                        speaker.speak("Semicolon");
                    else if (rtbDoc.SelectedText.ToString() == "?")
                        speaker.speak("Question mark");
                    else if (rtbDoc.SelectedText.ToString() == "[")
                        speaker.speak("third bracket Open");
                    else if (rtbDoc.SelectedText.ToString() == "]")
                        speaker.speak("third bracket Close");
                    else if (rtbDoc.SelectedText.ToString() == "`")
                        speaker.speak("Grave accent");
                    else if (rtbDoc.SelectedText.ToString() == "{")
                        speaker.speak("Second bracket Open");
                    else if (rtbDoc.SelectedText.ToString() == "}")
                        speaker.speak("Second bracket Close");
                    else if (rtbDoc.SelectedText.ToString() == "~")
                        speaker.speak("Equivalency sign ");
                    else if (rtbDoc.SelectedText.ToString() == "'")
                        speaker.speak("Single quote");
                    else
                        speaker.speak(rtbDoc.SelectedText.ToString());
                    rtbDoc.DeselectAll();
                }
                //MessageBox.Show(e.KeyCode.ToString());
                /*     if (pressCtrl && (e.KeyData.ToString().Equals("Up") || e.KeyData.ToString().Equals("Down")))
                     {

                         speaker.speak(rtbDoc.Text);
                         //speaker.speak("press control one for reading full text");
                         //speaker.speak("press left arrow to listen one character from left");
                        // speaker.speak("press right arrow to listen one character from right");
                         //speaker.speak("press alter f4 to terminate this application");
                     }
                */
            }
            catch (Exception ex)
            { }
        }

        private void rtbDoc_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (pressShift && e.KeyCode.ToString().Equals("Right"))
                {
                    if (thread != null) { thread.Abort(); thread = null; }
                    speaker.speak(rtbDoc.SelectedText.ToString());


                }
                else if (pressShift && e.KeyCode.ToString().Equals("Left"))
                {
                    if (thread != null) { thread.Abort(); thread = null; }
                    speaker.speak(rtbDoc.SelectedText.ToString());
                }
                else if (e.KeyCode.ToString().Equals("ControlKey"))
                {
                    //if (thread != null) { thread.Abort(); thread = null; }
                    pressCtrl = false;
                }
                else if (e.KeyCode.ToString().Equals("ShiftKey"))
                {
                    //if (thread != null) { thread.Abort(); thread = null; }
                    pressShift = false;

                }
                else if (pressCtrl && e.KeyCode.ToString().Equals("Left"))                       // read previous word
                {
                    if (thread != null) { thread.Abort(); thread = null; }
                    int cursorPosition = rtbDoc.SelectionStart;
                    int nextSpace = rtbDoc.Text.IndexOf(' ', cursorPosition);
                    int selectionStart = 0;
                    string trimmedString = string.Empty;
                    // Strip everything after the next space...
                    if (nextSpace != -1)
                    {
                        trimmedString = rtbDoc.Text.Substring(0, nextSpace);

                    }
                    else
                    {
                        trimmedString = rtbDoc.Text;

                    }


                    if (trimmedString.LastIndexOf(' ') != -1)
                    {
                        selectionStart = 1 + trimmedString.LastIndexOf(' ');
                        trimmedString = trimmedString.Substring(1 + trimmedString.LastIndexOf(' '));

                    }


                    rtbDoc.Select(rtbDoc.SelectionStart, trimmedString.Length);
                    speaker.speak(rtbDoc.SelectedText.ToString());
                    //speaker.stop();
                    rtbDoc.DeselectAll();
                    //if (thread != null) { thread.Abort(); }

                }
                else if (pressCtrl && e.KeyCode.ToString().Equals("Up"))                          // read 1st para & also use to read previous para
                {
                    if (thread != null) { thread.Abort(); thread = null; }
                    int cursorPosition = rtbDoc.SelectionStart;
                    int nextSpace = rtbDoc.Text.IndexOf("\n", cursorPosition);
                    int selectionStart = 0;
                    string trimmedString = string.Empty;
                    // Strip everything after the next space...
                    if (nextSpace != -1)
                    {
                        trimmedString = rtbDoc.Text.Substring(0, nextSpace);


                    }
                    else
                    {
                        trimmedString = rtbDoc.Text;
                    }


                    if (trimmedString.LastIndexOf("\n") != -1)
                    {
                        selectionStart = 1 + trimmedString.LastIndexOf("\n");
                        trimmedString = trimmedString.Substring(1 + trimmedString.LastIndexOf("\n"));
                    }
                    rtbDoc.Select(rtbDoc.SelectionStart, trimmedString.Length);
                    speaker.speak(rtbDoc.SelectedText.ToString());
                    //speaker.stop();
                    rtbDoc.DeselectAll();
                    //if (thread != null) { thread.Abort(); }

                }
                else if (e.KeyCode.ToString().Equals("Up"))                // read 2nd para to rest para
                {
                    try
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        int cursorPosition = rtbDoc.SelectionStart;
                        int lineIndex = rtbDoc.GetLineFromCharIndex(cursorPosition);
                        string lineText = rtbDoc.Lines[lineIndex];
                        speaker.speak(lineText.ToString());
                    }
                    catch (Exception ex) { }
                }
                else if (e.KeyCode.ToString().Equals("Down"))                // read 2nd para to rest para
                {
                    try
                    {
                        if (thread != null) { thread.Abort(); thread = null; }
                        int cursorPosition = rtbDoc.SelectionStart;
                        int lineIndex = rtbDoc.GetLineFromCharIndex(cursorPosition);
                        string lineText = rtbDoc.Lines[lineIndex];
                        speaker.speak(lineText.ToString());
                    }
                    catch (Exception ex) { }
                }
                else if (e.KeyData.ToString().Equals("Left"))                                    //read 1 char from back
                {
                    if (thread != null) { thread.Abort(); thread = null; }
                    rtbDoc.Select(rtbDoc.SelectionStart, 1);
                    //MessageBox.Show(rtbDoc.SelectedText.ToString());
                    speaker.stop();

                    if (rtbDoc.SelectedText.ToString() == " ")
                        speaker.speak("Space");
                    if (rtbDoc.SelectedText.ToString() == "!")
                        speaker.speak("Exclamation mark");
                    //else if (rtbDoc.SelectedText.ToString() == """)
                    //    speaker.speak("Double quotes");
                    //else if (rtbDoc.SelectedText.ToString() == """)
                    //    speaker.speak("Double quotes");
                    else if (rtbDoc.SelectedText.ToString() == "(")
                        speaker.speak("first bracker Open");
                    else if (rtbDoc.SelectedText.ToString() == ")")
                        speaker.speak("first bracker Close");
                    else if (rtbDoc.SelectedText.ToString() == ",")
                        speaker.speak("Comma");
                    else if (rtbDoc.SelectedText.ToString() == "-")
                        speaker.speak("Hyphen");
                    else if (rtbDoc.SelectedText.ToString() == ".")
                        speaker.speak(" full stop");
                    else if (rtbDoc.SelectedText.ToString() == ":")
                        speaker.speak("Colon");
                    else if (rtbDoc.SelectedText.ToString() == ";")
                        speaker.speak("Semicolon");
                    else if (rtbDoc.SelectedText.ToString() == "?")
                        speaker.speak("Question mark");
                    else if (rtbDoc.SelectedText.ToString() == "[")
                        speaker.speak("third bracket Open");
                    else if (rtbDoc.SelectedText.ToString() == "]")
                        speaker.speak("third bracket Close");
                    else if (rtbDoc.SelectedText.ToString() == "`")
                        speaker.speak("Grave accent");
                    else if (rtbDoc.SelectedText.ToString() == "{")
                        speaker.speak("Second bracket Open");
                    else if (rtbDoc.SelectedText.ToString() == "}")
                        speaker.speak("Second bracket Close");
                    else if (rtbDoc.SelectedText.ToString() == "~")
                        speaker.speak("Equivalency sign ");
                    else if (rtbDoc.SelectedText.ToString() == "'")
                        speaker.speak("Single quote");

                    else
                        speaker.speak(rtbDoc.SelectedText.ToString());

                    rtbDoc.DeselectAll();
                }
            }
            catch (Exception ex)
            { }
        }

        #endregion

        public void pdfLoadFile(String fileName)
        {
            if (thread != null) { thread.Abort(); thread = null; }
            string strExt;
            pressCtrl = false;

            FileInfo fileInfo = new FileInfo(fileName);
            fN = null;
            Uri uri = new Uri(fileInfo.ToString());
            fN = Path.GetFileName(uri.LocalPath);
        }

        public void loadAFile(String fileName)
        {
            if (thread != null) { thread.Abort(); thread = null; }
            string strExt;
            pressCtrl = false;

            FileInfo fileInfo = new FileInfo(fileName);
            //MessageBox.Show(fileInfo.ToString());

            //
            fN = null;
            Uri uri = new Uri(fileInfo.ToString());
            fN = Path.GetFileName(uri.LocalPath);
            //speaker.speak(fN);


            //
            strExt = fileInfo.Extension.ToString();
            strExt = strExt.ToUpper();

            if (strExt == ".RTF")
            {
                rtbDoc.LoadFile(OpenFileDialog1.FileName, RichTextBoxStreamType.RichText);
            }
            else
            {
                //System.IO.StreamReader txtReader;
                //txtReader = new System.IO.StreamReader(fileName);
                StreamReader sr = new StreamReader(fileName, Encoding.UTF8);
                String line = sr.ReadToEnd();
                rtbDoc.Text = line;
                sr.Close();
                //txtReader.Close();
                //txtReader = null;
                rtbDoc.SelectionStart = 0;
                rtbDoc.SelectionLength = 0;
            }
        }

        private void SaveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void rtbDoc_TextChanged(object sender, EventArgs e)
        {
            if (thread != null) { thread.Abort(); thread = null; }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            speaker.speak("checking");
            fSpellCheck(rtbDoc);

        }


        public void fSpellCheck(RichTextBoxPrintCtrl rtbDoc)
        {

            int iErrorCount = 0;
            Word.Application app = new Word.Application();

            if (rtbDoc.Text.Length > 0)
            {
                app.Visible = false;

                object template = Missing.Value;
                object newTemplate = Missing.Value;
                object documentType = Missing.Value;
                object visible = true;
                object optional = Missing.Value;


                _Document doc = app.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);//pass korcha
                doc.Words.First.InsertBefore(rtbDoc.Text);


                Word.ProofreadingErrors we = doc.SpellingErrors;//adi//A collection of spelling and grammatical errors for the specified document or range. There is no ProofreadingError object; instead, each item in the ProofreadingErrors collection is a Range object that represents one spelling or grammatical error.


                iErrorCount = we.Count;

                speaker.speak("Error numbers" + iErrorCount.ToString());//adi//returns a string representation of that object(irrorcount).//bondho corla count korana

                //doc.CheckSpelling(ref optional, ref optional, ref optional, ref optional,
                //ref optional, ref optional, ref optional,
                //ref optional, ref optional, ref optional, ref optional, ref optional);//adi


                object first = 0;
                object last = doc.Characters.Count - 1;

                rtbDoc.Text = doc.Range(ref first, ref last).Text;//value pass korcha main method a. ref na hoa example: int first mana value accept korcha
            }
            object saveChanges = false;
            object originalFormat = Missing.Value;
            object routeDocument = Missing.Value;
            app.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
        }

        private void rtbDoc_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                //MessageBox.Show("fffffff");
                if ((e.KeyChar >= 33 && e.KeyChar <= 64) || (e.KeyChar >= 91 && e.KeyChar <= 96) || (e.KeyChar >= 123 && e.KeyChar <= 126))
                    speaker.speak(e.KeyChar.ToString());
                if (e.KeyChar == '!')
                    speaker.speak("Exclamation mark");
                else if (e.KeyChar == '"')
                    speaker.speak("Double quotes");
                else if (e.KeyChar == '(')
                    speaker.speak("first bracker Open");
                else if (e.KeyChar == ')')
                    speaker.speak("first bracker Close");
                else if (e.KeyChar == ',')
                    speaker.speak("Comma");
                else if (e.KeyChar == '-')
                    speaker.speak("Hyphen");
                else if (e.KeyChar == '.')
                    speaker.speak(" full stop");
                else if (e.KeyChar == ':')
                    speaker.speak("Colon");
                else if (e.KeyChar == ';')
                    speaker.speak("Semicolon");
                else if (e.KeyChar == '?')
                    speaker.speak("Question mark");
                else if (e.KeyChar == '[')
                    speaker.speak("Third bracket Open");
                else if (e.KeyChar == ']')
                    speaker.speak("third bracket Close");
                else if (e.KeyChar == '`')
                    speaker.speak("Grave accent");
                else if (e.KeyChar == '{')
                    speaker.speak("Second brace Open");
                else if (e.KeyChar == '}')
                    speaker.speak("Second brace Close");
                else if (e.KeyChar == '~')
                    speaker.speak("Equivalency sign");
                else if (e.KeyChar == 39)
                    speaker.speak("Single quote");
                else if (e.KeyChar == ' ')
                    speaker.speak("Space");
                //else speaker.speak(e.ToString());
            }
            catch (Exception ex)
            { }

        }

        private void FontDialog1_Apply(object sender, EventArgs e)
        {

        }


    }


}
