using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Text.RegularExpressions;
using System.Collections;
using SpeechBuilder;
using RichTextEditor;
using PaglaPlayer;
using org.pdfbox.pdmodel;
using org.pdfbox.util;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PPT = Microsoft.Office.Interop.PowerPoint;


namespace Search
{
    public partial class Form1 : Form
    {
        private int occupiedBuffer = 0;
        private static Word.Application word = null;
        private static Excel.Application excel = null;
        private static PPT.Application ppt = null;
        private SpeechControl speaker;
        private Boolean flag;
        private ListViewItem currentSelectedItem;
        private ArrayList Dirs = new ArrayList();
        private ThreadList thrdList = new ThreadList();
        private string ContainingFolder = "";
        private String SearchPattern = "";
        private String SearchForText = "";
        private bool CaseSensitive = false;
        private String dr;
        private String ur;
        Thread oT = null;
        Process proc = new Process();
        private String FN;
        private int countSearchResult;

        public Form1(SpeechControl speaker)
        {
            InitializeComponent();
            this.speaker = speaker;
            //ur = u;
            //MessageBox.Show(ur);
            //urlBox.Text = u;
            //String url = Thesis.Form1.URL;
                  
        }
        public Form1(Word.Application word1)
        {
            word = word1;

        }
        public Form1(Excel.Application excel1)
        {
            excel = excel1;

        }
        public Form1(PPT.Application ppt1)
        {
            ppt = ppt1;
        }
       
        public void runDoc()
        {
            try
            {
                Monitor.Enter(this);
                //System.Threading.Thread.Sleep(3000);

                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    int p = 0;
                    Process[] pros = Process.GetProcesses();
                    for (int i = 0; i < pros.Count(); i++)
                    {
                        if (pros[i].ProcessName.ToLower().Contains("winword"))
                        {
                            p++;
                            //if(p>1)pros[i].Kill();                      
                        }
                    }
                    if (p == 0)
                    {
                        word = new Microsoft.Office.Interop.Word.Application();
                        
                    }
                    DocForm.DocForm docForm = new DocForm.DocForm(FN, speaker, word, p);

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);

            }
            catch (Exception )
            {

            }
        }
        public void runPPT()
        {
            try
            {
                Monitor.Enter(this);
                //System.Threading.Thread.Sleep(3000);

                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;
                    int p = 0;

                    Process[] pros = Process.GetProcesses();
                    for (int i = 0; i < pros.Count(); i++)
                    {
                        if (pros[i].ProcessName.ToLower().Contains("powerpnt"))
                        {
                            p++;
                            //if(p>1)pros[i].Kill();
                            //MessageBox.Show(p.ToString());
                        }
                    }
                    //MessageBox.Show(p.ToString());
                    if (p == 0)
                    {
                        ppt = new Microsoft.Office.Interop.PowerPoint.Application();
                    }
                    PPTForm.Form1 PpT1 = new PPTForm.Form1(FN, speaker, ppt, p);

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);

            }
            catch (Exception ex)
            {

            }
        }
        public void runExcel()
        {
            try
            {
                //System.Threading.Thread.Sleep(3000);
                Monitor.Enter(this);

                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);

                }
                else
                {

                    occupiedBuffer = 1;

                    int p = 0;
                    Process[] pros = Process.GetProcesses();
                    for (int i = 0; i < pros.Count(); i++)
                    {
                        if (pros[i].ProcessName.ToLower().Contains("excel"))
                        {
                            p++;
                            //if(p>1)pros[i].Kill();
                            //MessageBox.Show(p.ToString());
                        }
                    }
                    //MessageBox.Show(p.ToString());
                    if (p == 0)
                    {
                        excel = new Microsoft.Office.Interop.Excel.Application();
                    }

                    ExcellForm.Form1 ExF1 = new ExcellForm.Form1(FN, speaker, excel, p);

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);

            }
            catch (Exception ex)
            {

            }
        }
        public void executeApp(String fName)
        {
            Thread thread = null;      
            FileInfo file = new FileInfo(fName);
            frmMain notepad = new frmMain(speaker);
            
            String extension = file.Extension.ToLower();

            if (extension.Equals(".txt") || extension.Equals(".rtf") || extension.Equals(".RTF"))
            {
                //MessageBox.Show("Text");
                notepad.loadAFile(fName);
                notepad.Show();

            }
            //else if (extension.Equals(".pdf"))
            //{
            //    PdfRead.PdfRead pdfRead = new PdfRead.PdfRead(fName, speaker);
            //}
            else if (file.Extension.Equals(".doc") || file.Extension.Equals(".docx"))
            {
                FN = fName;                
                thread = new Thread(new ThreadStart(runDoc));
                thread.Start();
                thread.Join();

            }
            else if (file.Extension.Equals(".xls") || file.Extension.Equals(".xlsx"))
            {
                FN = fName;
                thread = new Thread(new ThreadStart(runExcel));
                thread.Start();
                thread.Join();
            }
            else if (file.Extension.Equals(".ppt") || file.Extension.Equals(".pptx"))
            {
                FN = fName;
                thread = new Thread(new ThreadStart(runPPT));
                thread.Start();
                thread.Join();
            }
            else if (isAMediaFile(file.Extension))
            {
                PaglaPlayerPane player = new PaglaPlayerPane();
                player.startPlaying(getAllPlayList(fName));
                player.Show();
            }
            //else if (file.Extension.Equals(".htm")
            //    || file.Extension.Equals(".html")
            //    || file.Extension.Equals(".pdf"))
            //{
            //    if (reader != null)
            //    {
            //        reader.Peek();
            //    }
            //    Console.WriteLine(fName);
            //    reader = new FilterReader("C:\\Documents and Settings\\Faisal\\Desktop\\Faisal.pdf"); ///////************ pdf
            //    string ss = reader.ReadToEnd();
            //    reader.Close();
            //    reader = null;
            //    MessageBox.Show(ss);
            //    //speaker.speak(stripper.getText(doc));
            //    //f2.richTextBox1.Text = ss;
            //    //f2.Show();

   //    //MessageBox.Show(ss);
            //    //speaker.speak(ss);

   //}
            else if (file.Extension.Equals(".pdf") || file.Extension.Equals(".html") || file.Extension.Equals(".htm"))
            {
                try
                {
                    //String fe = "E:\\Faisal\\Resume_of_faisal.pdf";
                    PDDocument d = PDDocument.load(fName);
                    PDFTextStripper stripper = new PDFTextStripper();
                    //MessageBox.Show(stripper.getText(d));
                    notepad.rtbDoc.Text = stripper.getText(d);
                    notepad.Show();
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
            else if (file.Extension.Equals(".exe"))
            {
                Process.Start(fName);
            }
        }
        private Boolean isAMediaFile(String extension)
        {
            extension = extension.ToLower();

            if (extension.Equals(".mp3") || extension.Equals(".wmv")
                || extension.Equals(".wav") || extension.Equals(".avi")
                || extension.Equals(".wma")
                )
            {
                return true;
            }
            return false;
        }
        private String[] getAllPlayList(String fileName)
        {
            FileInfo fileInfo = new FileInfo(fileName);
            DirectoryInfo folder = new DirectoryInfo(fileInfo.Directory.FullName);
            //MessageBox.Show(folder.FullName);
            String[] allFiles = Directory.GetFiles(folder.FullName);

            String[] playList = new String[500];


            int i = 0;
            foreach (String file in allFiles)
            {
                fileInfo = new FileInfo(file);

                String extension = fileInfo.Extension;
                if (isAMediaFile(extension))
                {
                    playList[i++] = fileInfo.FullName;
                    //MessageBox.Show( fileInfo.FullName );
                }
            }
            return playList;
        }


       // my

        public delegate void AddListBoxItemDelegate(String Text);

        public delegate void UpdateThreadStatusDelegate(String thrdName, SearchThreadState sts);

        public void UpdateThreadStatus(String thrdName, SearchThreadState sts)
        {
            SearchThread st = thrdList.Item(thrdName);
            st.state = sts;
            countSearchResult = listView1.Items.Count;
            if (countSearchResult == 0)
                speaker.speak("Search Completed no file found for the searching");
            else
                speaker.speak("Search Completed searching result total " + countSearchResult + " file found");
        }
        public void seturl(string s)
        {
            urlBox.Text = s;
            //MessageBox.Show(s);
            //urlBox.Text = s;
        }
        public void AddListBoxItem(String Text)
        {
            // I use Monitor to synchronize access 
            // to the file founded list
            listView1.Show();
            Monitor.Enter(listView1);

            listView1.Items.Add(Text);

            Monitor.Exit(listView1);
        }

        private void searchbutton1_Click(object sender, EventArgs e)
        {

            // empty thread list
            for (int i = thrdList.ItemCount() - 1; i >= 0; i--)
            {
                thrdList.RemoveItem(i);
            }

            // clear the file founded list
            listView1.Items.Clear();
            ContainingFolder = "";

            // get the search pattern
            // or use a default
            SearchPattern = "*" + search.Text.Trim() + "*";
            if (SearchPattern.Length == 0)
            {
                SearchPattern = "*.*";
            }



            // clear the Dirs arraylist
            Dirs.Clear();

            // check if each selected drive exists
            //foreach (int Index in disksListBox.CheckedIndices)
            //{
            // chek if drive is ready
            String Dir = urlBox.Text; //disksListBox.Items[Index].ToString().Substring(0, 2);
            Dir += @"\";
            //if (CheckExists(Dir))
            //{
            //    Dirs.Add(Dir);
            //}
            //}

            // I use 1 thread for each dir to scan
            //foreach (String Dir in Dirs)
            //{
            
            string thrdName = "Thread" + ((int)(thrdList.ItemCount() + 1)).ToString();
            FileSearch fs = new FileSearch(Dir, SearchPattern, SearchForText, CaseSensitive, this, thrdName);
            oT = new Thread(new ThreadStart(fs.SearchDir));

            oT.Name = thrdName;

            SearchThread st = new SearchThread();
            st.searchdir = Dir;
            st.name = oT.Name;
            st.thrd = oT;
            st.state = SearchThreadState.ready;
            thrdList.AddItem(st);
            if (search.Text != "")
            {
                oT.Start();                
            }

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //urlBox.Text=Thesis.Form1.URL.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Thesis.Form1 F1 = new Thesis.Form1(speaker);
            ////urlBox.Text = F1.urlBox.Text;
            //MessageBox.Show(F1.urlBox.Text);
            oT.Abort();
            
        }

        
        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected)
            {
                dr = listView1.Items[e.ItemIndex].Text;
            }
            //MessageBox.Show(dr);
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListView listView1 = sender as ListView;
            currentSelectedItem = (listView1.SelectedItems.Count > 0 ? listView1.SelectedItems[0] : null);


            if (currentSelectedItem != null)
            {
                //MessageBox.Show(currentSelectedItem.Text);
                speaker.stop();
                speaker.speak(currentSelectedItem.Text);

            }
        }
        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {            
            if (e.KeyCode.Equals(Keys.Enter))
            {
                executeApp(dr);
                this.Close();
                
            }
        }

        private void search_Enter(object sender, EventArgs e)
        {
            speaker.speak("Write down the file text what you search");
        }

        private void button2_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {            
           
        }

        private void searchbutton1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {           
            if (!e.KeyCode.ToString().Equals("Return"))
            {
                if(oT!=null)
                    oT.Abort();
            }

            if (e.KeyCode.ToString().Equals("Tab"))
            {
                if (listView1.Items.Count != 0)
                {
                    listView1.Items[0].Focused = true;
                    listView1.Items[0].Selected = true;
                }
            }
        }       
        
    }
}
