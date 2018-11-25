using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Collections;
using System.IO;
using DocForm;
using ExcellForm;
using PPTForm;
//using PdfRead;
using PaglaPlayer;
using System.Threading;
using SpeechBuilder;
using RichTextEditor;
using EPocalipse.IFilter;
using System.Diagnostics;
using Thesis;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PPT = Microsoft.Office.Interop.PowerPoint;
//using org.pdfbox.pdmodel;
//using org.pdfbox.util;

namespace ManualFolderBrowser
{
    public class Browser
    {
        private int occupiedBuffer = 0;
        private Word.Application word = null;
        private Excel.Application excel = null;
        private PPT.Application ppt = null;
        private ListView listView;
        private ImageList imageList;
        private ListViewItem[] listViewItem;
        private ListViewItem currentSelectedItem;
        private SpeechControl speaker;
        private static String selectedItemName = null;
        private static String SourceDir = null;
        private static String DesDir = null;
        private static String sourceFile = null;
        private static String destFile = null;
        private static String nme = null;
        private String Old = null;
        private String New = null;
        private Boolean pressF2 = false;
        private Boolean ctrl = false;
        private Boolean insrt = false;
        private static int ShutDownMonitor;
        Thread thread1 = null;

        Process proc = new Process();

        private TextReader reader;

        private int noOfItem;
        private String path;
        private String FN;
        private String[] files;
        private String[] folders;
        private TextBox urlBox;

        private int pp = 0;

        float maxbytes = 0, fileSz = 0;
        int copied = 0;
        int total = 0;

        int d5,d10,d20,d30,d40,d50,d60,d70,d80,d90;


        public Browser(ListView listView, SpeechControl speaker, TextBox urlBox)
        {
            imageList = new ImageList();
            this.listView = listView;
            listView.Select();
            //speaker = new SpeechControl();
            this.speaker = speaker;
            this.urlBox = urlBox;
            reader = null;
            //word = new Microsoft.Office.Interop.Word.Application();
            //excel = new Microsoft.Office.Interop.Excel.Application();
            //ppt = new Microsoft.Office.Interop.PowerPoint.Application();            
            init();
        }        
        public Browser(String s)
        {
            //Process[] procs = Process.GetProcessesByName("nvda");
            //if (procs.Length != 0)
            //{
            //    foreach (Process proc in procs)
            //        proc.Kill();
            //}
            selectedItemName = s;
            //MessageBox.Show(selectedItemName); 
        }
        private void init()
        {
            /*
             * initialize the list view
             */

            //*************************************************//
            listView.View = View.LargeIcon;
            listView.LargeImageList = imageList;
            listView.UseCompatibleStateImageBehavior = false;
            listView.TabIndex = 0;
            listView.SelectedIndexChanged += new System.EventHandler(listView_SelectedIndexChanged);
            listView.KeyDown += new KeyEventHandler(listView_KeyDown);
            listView.KeyUp+=new KeyEventHandler(listView_KeyUp);
            listView.MouseDoubleClick += new MouseEventHandler(listView_MouseDoubleClick);
            //listView.Sorting = SortOrder.Ascending;
            //**************************************************//

            /*
             * end of initalize for the listview
             */


            /*default path*/
            //********************************//

            DirectoryInfo dirInfo = new DirectoryInfo(getSystemDrive() + "Users\\" + Environment.UserName + "\\Desktop\\");
            if(dirInfo.Exists)
                path = getSystemDrive() + "Users\\" + Environment.UserName + "\\Desktop\\";
            else
                path = getSystemDrive() + "Documents and Settings\\" + getCurrentUserName() + "\\Desktop\\";
            //MessageBox.Show(dirInfo.Exists.ToString());
            urlBox.Text = path;

            //********************************//

            /*
             * Starting listview item = 0;
             */
            //****************************//
            noOfItem = 0;
            //****************************//

            /*
             * starting selected item is null
             */
            //******************************//
            currentSelectedItem = null;
            //******************************//

            /*
             * Adding images corresponding filetypes and folder
             */
            //*********************************************//
            imageList.Images.Add(Image.FromFile("folder.jpg"));
            imageList.Images.Add(Image.FromFile("pdf.jpg"));
            imageList.Images.Add(Image.FromFile("doc.jpg"));
            imageList.Images.Add(Image.FromFile("notepad.jpg"));
            imageList.Images.Add(Image.FromFile("unknown.jpg"));
            imageList.Images.Add(Image.FromFile("wordpad.jpg"));
            imageList.Images.Add(Image.FromFile("html.jpg"));

            imageList.Images.Add(Image.FromFile("hdd.jpg"));
            imageList.Images.Add(Image.FromFile("cd.jpg"));
            imageList.Images.Add(Image.FromFile("removable.jpg"));

            imageList.Images.Add(Image.FromFile("exe.jpg"));
            imageList.Images.Add(Image.FromFile("media.jpg"));

            //*********************************************//

            /*
             * icon size
             */
            //*************************************//
            imageList.ImageSize = new Size(35, 35);
            //*************************************//

            settingBrowser();

            return;
        }

        public void setPath(String path)
        {
            this.path = path;
        }
        public String getCurrentPath()
        {
            return this.path;
        }
        public String getSelectedItem()
        {
            return currentSelectedItem.Text;
        }
        public static void copyDirectory(string Src, string Dst) //folder copy method
        {
            String[] Files;

            if (Dst[Dst.Length - 1] != Path.DirectorySeparatorChar) Dst += Path.DirectorySeparatorChar;
            if (!Directory.Exists(Dst)) Directory.CreateDirectory(Dst);
            Files = Directory.GetFileSystemEntries(Src);
            foreach (string Element in Files)
            {
                // Sub directories
                if (Directory.Exists(Element)) copyDirectory(Element, Dst + Path.GetFileName(Element));
                // Files in directory
                else File.Copy(Element, Dst + Path.GetFileName(Element), true);
            }
        }

        /// <summary>
        /// resetting the browser
        /// </summary>
        /// 
        public void settingBrowser()
        {

            /*
             * getting all files and folders
             */
            //*****************************//
            folders = getFolders();
            files = getFiles();
            //*****************************//

            listView.Clear();

            listViewItem = new ListViewItem[noOfItem + 1];


            /*
             * set all the files and folders in the browser
             */
            //****************************//
            setFoldersAndFiles(folders, files);
            //****************************//
            urlBox.Text = path;
            if (pp == 0)
                speaker.speak(path);
            pp = pp + 1;
            return;
        }

        /// <summary>
        /// event handler method for list view 
        /// it's monitor the keyboard key pressing
        /// and stored current selected item
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListView listView = sender as ListView;
            currentSelectedItem = (listView.SelectedItems.Count > 0 ? listView.SelectedItems[0] : null);

            if (currentSelectedItem != null)
            {
                //MessageBox.Show(currentSelectedItem.Text);
                speaker.stop();

                try
                {
                    DriveInfo driveInfo = new DriveInfo(currentSelectedItem.Text);

                    if (driveInfo.DriveType.ToString().Equals("CDRom"))
                    {
                        //speaker.speak(currentSelectedItem.Text + " " + "CD Rom");
                    }
                    else if (driveInfo.DriveType.ToString().Equals("Fixed"))
                    {
                        //speaker.speak(currentSelectedItem.Text + " " + "Hard disk drive");
                    }
                    else
                    {
                        //speaker.speak(currentSelectedItem.Text + " " + driveInfo.DriveType);
                    }
                }
                catch (Exception)
                {
                    //speaker.speak(currentSelectedItem.Text);
                }
            }
            return;
        }

        /// <summary>
        /// this is a keydown event handler
        /// when a user presses enter or back space then it will do it's job
        /// when user presses enter then it check whether it is folder or not
        /// it folder then it take all files and folders inside the folder and
        /// reset the browser for setting all the folder and files
        /// when user presses back space then it return back to it's parent folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void listView_KeyDown(object sender, KeyEventArgs e)
        {
            //MessageBox.Show(e.KeyCode.GetHashCode().ToString());
            if (thread1 != null) { thread1.Abort(); thread1 = null; }
            //if (ShutDownMonitor == 1)
            //{
            //    if (e.KeyCode.ToString().Equals("Y"))/////////
            //    {
            //        System.Diagnostics.Process.Start("shutdown.exe", "-s -t 0");
            //    }
            //    else ShutDownMonitor = 0;
            //}
            if (e.KeyCode.Equals(Keys.ControlKey)) ctrl = true;
            if (e.KeyCode.Equals(Keys.Insert)) insrt = true; 
            else if (e.KeyCode.Equals(Keys.Enter))
            {
                //checking that the path is folder or not if yes then go inside
                //if files then execute it
                /*************************************************************/
                try
                {
                    if (Directory.Exists(path + currentSelectedItem.Text + "\\"))
                    {
                        path += currentSelectedItem.Text + "\\";
                        //MessageBox.Show(path);
                        settingBrowser();

                    }
                    else if (File.Exists(path + currentSelectedItem.Text + "\\"))
                    {
                        //Console.WriteLine("path" + path + currentSelectedItem.Text);
                        //MessageBox.Show(currentSelectedItem.Text);
                        String s = currentSelectedItem.Text.Substring(0, 2);
                        if (s != "~$")                        
                            executeApp(path + currentSelectedItem.Text);
                    }
                    else
                    {

                        DriveInfo driveInfo = new DriveInfo(currentSelectedItem.Text);

                        if (driveInfo.IsReady)
                        {
                            path = currentSelectedItem.Text;
                            urlBox.Text = path;

                            settingBrowser();
                        }

                    }
                }
                catch
                {
                    speaker.speak("Enter");
                }
                /***************************************************************/
            }
           
            else if (e.KeyCode.Equals(Keys.F1))
            {
                speaker.stop();
                RichTextEditor.Form1 AA = new RichTextEditor.Form1(speaker);
                if (thread1 != null) { thread1.Abort(); thread1 = null; }
                thread1 = new Thread(new ThreadStart(AA.AsstForBroser));
                thread1.Start();
            }
            else if (e.KeyCode.Equals(Keys.Delete))
            {
                try
                {
                    String url = urlBox.Text;
                    String Dir = url + selectedItemName;
                    FileInfo file = new FileInfo(Dir);
                    String extension = file.Extension.ToLower();
                    String tt = "Do You Really want to Delete " + selectedItemName + " from your computer";

                    if (selectedItemName == null)
                    {
                        return;
                    }

                    if ((extension.Equals(".tmp")) || (extension.Equals(".xls")) || (extension.Equals(".xlsx")) || (extension.Equals(".mpeg")) || (extension.Equals(".dat")) || (extension.Equals(".cam")) || (extension.Equals(".flv")) || (extension.Equals(".3gp")) || (extension.Equals(".gif")) || (extension.Equals(".otf")) || (extension.Equals(".ttc")) || (extension.Equals(".ttf")) || (extension.Equals(".xml")) || (extension.Equals(".xhtml")) || (extension.Equals(".dotx")) || (extension.Equals(".dot")) || (extension.Equals(".tar")) || (extension.Equals(".rar")) || (extension.Equals(".iso")) || (extension.Equals(".jpz")) || (extension.Equals(".gz")) || (extension.Equals(".jar")) || (extension.Equals(".bmp")) || (extension.Equals(".7z")) || (extension.Equals(".pptx")) || (extension.Equals(".docx")) || (extension.Equals(".exe")) || (extension.Equals(".html")) || (extension.Equals(".htm")) || (extension.Equals(".rtf")) || (extension.Equals(".txt")) || (extension.Equals(".ppt")) || (extension.Equals(".doc")) || (extension.Equals(".pdf")) || (extension.Equals(".mp3")) || extension.Equals(".wmv") || extension.Equals(".wav") || extension.Equals(".avi") || extension.Equals(".wma") || extension.Equals(".dll"))
                    {

                        speaker.speak(tt);
                        try
                        {
                            if (MessageBox.Show("", "Do You Really want to Delete " + selectedItemName + " from your computer", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                File.Delete(@Dir);
                                speaker.speak("File Delete completed");
                            }
                            else return;
                        }
                        catch (Exception ex)
                        {
                            speaker.speak("This file is already use please close the file to delete the file");
                            //System.Threading.Thread.Sleep(4000);
                        }
                        settingBrowser();
                    }

                    else
                    {
                        speaker.speak(tt);
                        if (MessageBox.Show("", "Do You Really want to Delete " + selectedItemName + " from your computer", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.IO.Directory.Delete(Dir, true);
                            settingBrowser();
                            speaker.speak("Folder Delete completed");
                        }
                        else return;
                    }

                }
                catch (Exception ex)
                {

                }

            }
              
                // URL info
            else if (ctrl && (e.KeyCode.Equals(Keys.U)))
            {
                speaker.speak(urlBox.ToString());
            }

                // File Info
            else if (ctrl && (e.KeyCode.Equals(Keys.I)))
            {
                try
                {
                    //////////////Driver Size                   

                    if (selectedItemName.Length == 3 && selectedItemName.Substring(1, 2) == ":\\")
                    {
                        float totalSz = 0, FreeSp = 0;
                        foreach (System.IO.DriveInfo label in System.IO.DriveInfo.GetDrives())
                        {
                            if (label.IsReady && label.Name == selectedItemName)
                            {
                                totalSz = label.TotalSize;
                                FreeSp = label.TotalFreeSpace;
                            }
                        }

                        totalSz = totalSz / (1024 * 1024 * 1024);
                        speaker.speak(" Driver " + selectedItemName + " total size " + totalSz.ToString("0.00") + " giga byte ");

                        FreeSp = FreeSp / (1024 * 1024 * 1024);

                        if (FreeSp < 1)
                        {
                            FreeSp = FreeSp * 1024;
                            speaker.speak(" Free Space " + FreeSp.ToString("0.00") + " mega byte");
                        }
                        else
                        {
                            speaker.speak(" Free Space " + FreeSp.ToString("0.00") + " giga byte");
                        }
                        return;
                    }

                    ///////////// DrSz

                    fileSz = 0;
                    String Directr = urlBox.Text;
                    String nme1 = selectedItemName;
                    String link = Directr + "\\" + nme1;

                    DirectoryInfo dinfo = new DirectoryInfo(link);
                    String Di = dinfo.Extension.ToLower();
                    //MessageBox.Show(Di);
                    if ((Di.Equals(".xlsx")) || (Di.Equals(".xls")) || (Di.Equals(".mpeg")) || (Di.Equals(".dat")) || (Di.Equals(".cam")) || (Di.Equals(".flv")) || (Di.Equals(".3gp")) || (Di.Equals(".gif")) || (Di.Equals(".otf")) || (Di.Equals(".ttc")) || (Di.Equals(".ttf")) || (Di.Equals(".xml")) || (Di.Equals(".xhtml")) || (Di.Equals(".dotx")) || (Di.Equals(".dot")) || (Di.Equals(".tar")) || (Di.Equals(".rar")) || (Di.Equals(".iso")) || (Di.Equals(".jpz")) || (Di.Equals(".gz")) || (Di.Equals(".jar")) || (Di.Equals(".bmp")) || (Di.Equals(".7z")) || (Di.Equals(".pptx")) || (Di.Equals(".docx")) || (Di.Equals(".exe")) || (Di.Equals(".html")) || (Di.Equals(".htm")) || (Di.Equals(".rtf")) || (Di.Equals(".txt")) || (Di.Equals(".ppt")) || (Di.Equals(".doc")) || (Di.Equals(".pdf")) || (Di.Equals(".mp3")) || Di.Equals(".wmv") || Di.Equals(".wav") || Di.Equals(".avi") || Di.Equals(".wma") || Di.Equals(".dll"))
                    {
                        FileInfo f = new FileInfo(link);
                        float s1 = f.Length;
                        // MessageBox.Show(s1.ToString());
                        float t;
                        t = s1 / 1024; // kb

                        if (t >= 1024)
                        {
                            t = t / 1024;
                            speaker.speak("File size " + t.ToString("0.00") + " Mega byte");
                        }
                        else if (t < 1)
                        {
                            t = t * 1024;
                            speaker.speak("File size " + t.ToString() + " bytes");
                        }
                        else
                            speaker.speak("File size " + t.ToString("0.00") + " kilo byte");

                    }

                    else
                    {
                        FileGetSize(dinfo);

                        float t;
                        t = fileSz;

                        t = fileSz / 1024;
                        // MessageBox.Show(t.ToString());
                        if (t >= 1024 * 1024)
                        {
                            t = t / (1024 * 1024);
                            speaker.speak("Folder size " + t.ToString("0.00") + " giga byte");
                        }
                        else if (t >= 1024)
                        {
                            t = t / 1024;
                            speaker.speak("Folder size " + t.ToString("0.00") + " Mega byte");
                        }
                        else
                        {
                            speaker.speak("Folder size " + t.ToString("0.00") + " kilo byte");
                        }

                    }
                }
                catch (Exception ex)
                { }

            }

            else if (ctrl && (e.KeyCode.Equals(Keys.C)))
            {
                //MessageBox.Show("this is C");
                String Dir = urlBox.Text;
                SourceDir = Dir;
                nme = selectedItemName;
                //ctrl = false;
                speaker.speak("You Press Control + c to copy" + nme);

                d5 = 0; d10 = 0; d20 = 0; d30 = 0; d40 = 0; d50 = 0; d60 = 0; d70 = 0; d80 = 0; d90 = 0;
            }
            else if (ctrl && (e.KeyCode.Equals(Keys.V)))
            {

                Thread t = new Thread(NewThread);
                t.Start();
            }

            else if ((e.KeyCode.Equals(Keys.F5)))
            {
                settingBrowser();
            }

            else if (ctrl && (e.KeyCode.Equals(Keys.F)))
            {
                //New Folder Add


                String url = urlBox.Text;
                String Dir = url + @"\New Folder";

                System.IO.DirectoryInfo dirInfo = new DirectoryInfo(Dir);

                if (!dirInfo.Exists) //checks if the directory already exists or not  
                {
                    dirInfo.Create();
                    settingBrowser();
                    for (int x = 0; x < listView.Items.Count; x++)
                    {
                        if (listView.Items[x].Text == "New Folder")
                        {
                            listView.Items[x].Focused = true;
                            listView.Items[x].Selected = true;
                        }
                    }
                    speaker.speak(" Add in the Location " + url);
                    return;
                }

                else
                {
                    for (int i = 2; ; i++)
                    {
                        String cc = url + @"\New Folder(" + i + ")";
                        System.IO.DirectoryInfo dirInfo1 = new DirectoryInfo(cc);
                        if (!dirInfo1.Exists) //checks if the directory already exists or not  
                        {
                            dirInfo1.Create();
                            settingBrowser();
                            for (int x = 0; x < listView.Items.Count; x++)
                            {
                                if (listView.Items[x].Text == "New Folder(" + i + ")")
                                {
                                    listView.Items[x].Focused = true;
                                    listView.Items[x].Selected = true;
                                }
                            }
                            speaker.speak(" Add in the Location " + url);
                            break;
                        }

                    }
                }

                //speaker.speak("There Has a Folder Name New Folder in the directory so please change the folder name to create new folder");                 

            }
            else if (e.KeyCode.Equals(Keys.F2))
            {
                try
                {
                    pressF2 = true;
                    Old = path + currentSelectedItem.Text;

                    ListViewItem item = listView.SelectedItems[0];
                    speaker.speak("Edit Mode");
                    item.BeginEdit();
                }

                catch (Exception)
                { }

                //if(item.Checked)
                //{
                //String New = path + " " + currentSelectedItem.Text;
                //MessageBox.Show(old);
                //
                //System.IO.File.Move()                
            }
            else if (e.KeyCode.Equals(Keys.Back))
            {
                DirectoryInfo dInfo = new DirectoryInfo(path);
                //to check the path is drive, if yes then stop resetting browser
                /***************************************************************/
                if (path.Length != dInfo.Name.Length)
                {
                    path = path.Remove(path.Length - dInfo.Name.Length - 1);
                    settingBrowser();
                    urlBox.Text = path;
                }
                /***************************************************************/
                else
                {
                    setDrives();
                }

            }
        }

        public void Copy1(string sourceDirectory, string targetDirectory)
        {
            try
            {
                DirectoryInfo diSource = new DirectoryInfo(sourceDirectory);
                DirectoryInfo diTarget = new DirectoryInfo(targetDirectory);
                //Gets size of all files present in source folder.
                GetSize(diSource, diTarget);
                maxbytes = maxbytes / 1024;


                //float kk = maxbytes / (1024 * 1024);

                //if (kk > 2.5)
                //{
                //    speaker.speak("File Size is Big");
                //    return;
                //}

                //MessageBox.Show(maxbytes.ToString());

                //progressBar1.Maximum = (int)maxbytes;

                //MessageBox.Show(maxbytes.ToString());

                CopyAll(diSource, diTarget);
            }
            catch (Exception ex)
            { }

        }

        public void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            try
            {
                int k = 5;

                if (Directory.Exists(target.FullName) == false)
                {
                    Directory.CreateDirectory(target.FullName);
                }
                foreach (FileInfo fi in source.GetFiles())
                {

                    fi.CopyTo(Path.Combine(target.ToString(), fi.Name), true);

                    total += (int)fi.Length;

                    copied += (int)fi.Length;
                    copied /= 1024;
                    //progressBar1.Step = copied;

                    //progressBar1.PerformStep();
                    //label1.Text = (total / 1048576).ToString() + "MB of " + (maxbytes / 1024).ToString() + "MB copied";
                    float total_sz = maxbytes / 1024;
                    float copy_complete = total / 1048576;

                    float compare = (copy_complete / total_sz) * 100;


                    if (compare > 90 && compare < 90.7 && d90 == 0) { MessageBox.Show("90% complete" + d90); k = 0; d90 = 1; }
                    else if (compare > 80 && compare < 80.7 && d80 == 0) { speaker.speak("80% complete"); k = 0; d80 = 1; }
                    else if (compare > 70 && compare < 70.7 && d70 == 0) { speaker.speak("70% complete"); k = 0; d70 = 1; }
                    else if (compare > 60 && compare < 60.7 && d60 == 0) { speaker.speak("60% complete"); k = 0; d60 = 1; }
                    else if (compare > 50 && compare < 50.7 && d50 == 0) { speaker.speak("50% complete"); k = 0; d50 = 1; }
                    else if (compare > 40 && compare < 40.7 && d40 == 0) { speaker.speak("40% complete"); k = 0; d40 = 1; }
                    else if (compare > 30 && compare < 30.7 && d30 == 0) { speaker.speak("30% complete"); k = 0; d30 = 1; }
                    else if (compare > 20 && compare < 20.7 && d20 == 0) { speaker.speak("20% complete"); k = 0; d20 = 1; }
                    else if (compare > 10 && compare < 10.7 && d10 == 0) { speaker.speak("10% complete"); k = 0; d10 = 1; }
                    else if (compare > 5 && compare < 5.7 && d5 == 0) { speaker.speak("5% complete"); k = 0; d5 = 1; }
                    if (k == 0)
                    {
                        speaker.speak("please wait");
                        k = 1;
                    }

                    //MessageBox.Show("total_sz" + total_sz + "  copy_complete" + copy_complete + " compare=" + compare.ToString());


                    //label1.Refresh();
                }
                foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
                {
                    DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                    CopyAll(diSourceSubDir, nextTargetSubDir);
                }
            }
            catch (Exception ex)
            { }
        }

        public void GetSize(DirectoryInfo source, DirectoryInfo target)
        {
            try
            {
                if (Directory.Exists(target.FullName) == false)
                {
                    Directory.CreateDirectory(target.FullName);
                }
                foreach (FileInfo fi in source.GetFiles())
                {
                    maxbytes += (int)fi.Length;//Size of File
                }
                foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
                {
                    DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                    GetSize(diSourceSubDir, nextTargetSubDir);

                }
            }
            catch (Exception ex)
            {
            }
        }

        public void FileGetSize(DirectoryInfo source)
        {
            try
            {
                foreach (FileInfo fi in source.GetFiles())
                {
                    fileSz += (int)fi.Length;//Size of File
                }
                foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
                {
                    FileGetSize(diSourceSubDir);
                }
            }
            catch (Exception ex)
            { 
            }
        }

        private void listView_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.Equals(Keys.ControlKey)) ctrl = false;
            if (e.KeyCode.Equals(Keys.Insert)) insrt = false;
            if (e.KeyCode.Equals(Keys.Enter))
            {
                if (pressF2)
                {
                    try
                    {
                        FileInfo file = null;
                        FileInfo file1 = null;
                        file = new FileInfo(Old);
                        String oldEx = file.Extension.ToLower().ToString();
                        //MessageBox.Show(oldEx);
                        New = path + currentSelectedItem.Text;
                        file1 = new FileInfo(New);
                        String newEx = file1.Extension.ToLower().ToString();
                        
                        if (oldEx == "") //folder
                        {
                            if (file1.Name.ToString() != "")
                            {
                                System.IO.Directory.Move(@Old, @New);
                                //MessageBox.Show(file.Name.ToString());
                                settingBrowser();

                                for (int x = 0; x < listView.Items.Count; x++)
                                {
                                    if (listView.Items[x].Text == file1.Name.ToString())
                                    {
                                        listView.Items[x].Focused = true;
                                        listView.Items[x].Selected = true;
                                    }
                                }
                            }
                            else
                            {                                
                                settingBrowser();
                                for (int x = 0; x < listView.Items.Count; x++)
                                {
                                    if (listView.Items[x].Text == file.Name.ToString())
                                    {
                                        listView.Items[x].Focused = true;
                                        listView.Items[x].Selected = true;
                                    }
                                }
                                speaker.speak("Please Enter Name to Update Folder");
                            }
                        
                        }
                        else if (oldEx == newEx)  //file
                        {
                            System.IO.File.Move(@Old, @New);
                            settingBrowser();
                            for (int x = 0; x < listView.Items.Count; x++)
                            {
                                if (listView.Items[x].Text == currentSelectedItem.Text)
                                {
                                    listView.Items[x].Focused = true;
                                    listView.Items[x].Selected = true;
                                }
                            }
                        }
                        else
                        {                            
                            settingBrowser();                            
                            for (int x = 0; x < listView.Items.Count; x++)
                            {                               
                                if (listView.Items[x].Text == file.Name.ToString())
                                {
                                    listView.Items[x].Focused = true;
                                    listView.Items[x].Selected = true;
                                }
                            }
                            speaker.speak("File Format not match");
                        }
                    }
                    catch (Exception)
                    { }
                }            
                                
            }
        }

        private void listView_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            //MessageBox.Show("the");
        }

        private void listView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (Directory.Exists(path + currentSelectedItem.Text + "\\"))
            {
                path += currentSelectedItem.Text + "\\";
                settingBrowser();
                urlBox.Text = path;
            }
            else if (File.Exists(path + currentSelectedItem.Text + "\\"))
            {
                executeApp(path + currentSelectedItem.Text);
            }
            else
            {
                DriveInfo driveInfo = new DriveInfo(currentSelectedItem.Text);

                if (driveInfo.IsReady)
                {
                    path = currentSelectedItem.Text;
                    settingBrowser();
                    urlBox.Text = path;
                }

            }
        }

        /*
             * setting Drive names and image into listview
             */
        //****************************************************//
        public void setDrives()
        {
            String[] drives = Environment.GetLogicalDrives();

            listView.Clear();
            listViewItem = new ListViewItem[drives.Length + 1];

            int i = 0;
            foreach (String drive in drives)
            {
                DriveInfo driveInfo = new DriveInfo(drive);

                listViewItem[i] = new ListViewItem(driveInfo.Name);
                listView.Items.Add(listViewItem[i]);

                if (driveInfo.DriveType.ToString().Equals("Fixed"))
                {
                    listViewItem[i].ImageIndex = (int)TAG.HDD;
                }
                else if (driveInfo.DriveType.ToString().Equals("CDRom"))
                {
                    listViewItem[i].ImageIndex = (int)TAG.CD;
                }
                else
                {
                    listViewItem[i].ImageIndex = (int)TAG.REMOVABLE;
                }
                i++;
            }
            urlBox.Text = "My Computer";
            speaker.speak("My Computer");
            return;

        }
        //****************************************************//

        private void setFoldersAndFiles(String[] folders, String[] files)
        {

            int i = 0;
            /*
             * setting folder names and image into listview
             */
            //*************************************************************//
            foreach (String folder in folders)
            {
                listViewItem[i] = new ListViewItem(getFolderName(folder));

                listView.Items.Add(listViewItem[i]);
                listViewItem[i].ImageIndex = (int)TAG.FOLDER;
                i++;
            }
            //*************************************************************//

            /*
             * setting file names and corresponding images into listView
             */
            //**************************************************************//
            foreach (String file in files)
            {
                listViewItem[i] = new ListViewItem(getFileName(file));
                listView.Items.Add(listViewItem[i]);
                listViewItem[i].ImageIndex = getFileFormat(file);
                i++;
            }
            //**************************************************************//
            return;
        }
        /// <summary>
        /// to know the format of a file
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>tag index</returns>

        private int getFileFormat(String fileName)
        {
            FileInfo fileInfo = new FileInfo(fileName);
            String extension = fileInfo.Extension.ToLower();

            if (extension == ".pdf")
            {
                return (int)TAG.PDF;
            }
            else if (extension == ".doc")
            {
                return (int)TAG.DOC;
            }
            else if (extension == ".docx")
            {
                return (int)TAG.DOCX;
            }
            else if (extension == ".txt")
            {
                return (int)TAG.NOTEPAD;
            }
            else if (extension == ".rtf")
            {
                return (int)TAG.WORDPAD;
            }
            else if (extension == ".html" || extension == ".htm")
            {
                return (int)TAG.HTML;
            }
            else if (extension == ".exe")
            {
                return (int)TAG.EXECUTABLE;
            }
            else if (extension.Equals(".mp3") || extension.Equals(".wmv")
                || extension.Equals(".wav") || extension.Equals(".avi")
                || extension.Equals(".wma"))
            {
                return (int)TAG.MEDIA;
            }
            else return (int)TAG.UNKNOWN;
        }

        /*
         * To know only the file name
         */
        //***************************************//
        private String getFileName(String file)
        {
            FileInfo fileInfo = new FileInfo(file);
            return fileInfo.Name;
        }
        //***************************************//

        /*
         * to know only the last folder name from a path
         */
        //*********************************************//
        private String getFolderName(String folder)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(folder);
            return directoryInfo.Name;
        }
        //********************************************

        /// <summary>
        /// get all the folders from a directory and count the no of folder
        /// </summary>
        /// <returns> all the folders </returns>
        private String[] getFolders()
        {
            String[] folders = Directory.GetDirectories(path);
            noOfItem = folders.Length;
            return folders;
        }
        /// <summary>
        /// get all the files from a directory and count no of files
        /// </summary>
        /// <returns> all the files </returns>
        private String[] getFiles()
        {
            String[] files = Directory.GetFiles(path);
            noOfItem += files.Length;
            return files;
        }

        /// <summary>
        /// getting the drive name where user installed windows
        /// </summary>
        /// <returns></returns>
        private String getSystemDrive()
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(Environment.SystemDirectory);
            return directoryInfo.Root.ToString();
        }
        /// <summary>
        /// get the user name
        /// </summary>
        /// <returns></returns>
        private String getCurrentUserName()
        {
            return Environment.UserName;
        }


        /// <summary>
        /// using an enumerator for tagging constant and using image indexing
        /// </summary>

        private enum TAG
        {
            FOLDER = 0,
            PDF = 1,
            DOC = 2,
            DOCX = 2,
            NOTEPAD = 3,
            UNKNOWN = 4,
            WORDPAD = 5,
            HTML = 6,
            HDD = 7,
            CD = 8,
            REMOVABLE = 9,
            EXECUTABLE = 10,
            MEDIA = 11
        }
        //public void runNarrator()
        //{
            //try
            //{
            //    Monitor.Enter(this);
            //    if (occupiedBuffer == 1)
            //    {
            //        Monitor.Wait(this);
            //    }
            //    else
            //    {

            //        occupiedBuffer = 1;
            //        Process[] pname = Process.GetProcessesByName("nvda");
            //        if (pname.Length == 0)
            //        {
            //            string targetDir = string.Format(@"C:\Resources\NVDA");//this is directory
            //            proc.StartInfo.WorkingDirectory = targetDir;
            //            proc.StartInfo.FileName = "nvda.exe";
            //            proc.StartInfo.Arguments = string.Format("10");//this is argument
            //            proc.StartInfo.CreateNoWindow = false;
            //            proc.Start();
            //        }
            //        occupiedBuffer = 0;                   

            //    }
            //    Monitor.Pulse(this);
            //    Monitor.Exit(this);

            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            //}
        //}

        private void NewThread()
        {
            //MessageBox.Show(SourceDir);
            //String url = urlBox.Text;
            if (SourceDir != null)
            {
                String Dir = urlBox.Text;
                DesDir = Dir;
                FileInfo file = new FileInfo(SourceDir + nme);
                String extension = file.Extension.ToLower();
                if ((extension.Equals(".xlsx")) || (extension.Equals(".xls")) || (extension.Equals(".mpeg")) || (extension.Equals(".dat")) || (extension.Equals(".cam")) || (extension.Equals(".flv")) || (extension.Equals(".3gp")) || (extension.Equals(".gif")) || (extension.Equals(".otf")) || (extension.Equals(".ttc")) || (extension.Equals(".ttf")) || (extension.Equals(".xml")) || (extension.Equals(".xhtml")) || (extension.Equals(".dotx")) || (extension.Equals(".dot")) || (extension.Equals(".tar")) || (extension.Equals(".rar")) || (extension.Equals(".iso")) || (extension.Equals(".jpz")) || (extension.Equals(".gz")) || (extension.Equals(".jar")) || (extension.Equals(".bmp")) || (extension.Equals(".7z")) || (extension.Equals(".pptx")) || (extension.Equals(".docx")) || (extension.Equals(".exe")) || (extension.Equals(".html")) || (extension.Equals(".htm")) || (extension.Equals(".rtf")) || (extension.Equals(".txt")) || (extension.Equals(".ppt")) || (extension.Equals(".doc")) || (extension.Equals(".pdf")) || (extension.Equals(".mp3")) || extension.Equals(".wmv") || extension.Equals(".wav") || extension.Equals(".avi") || extension.Equals(".wma") || extension.Equals(".dll"))
                {
                    if (SourceDir == DesDir)
                    {
                        speaker.speak("Source and Destination directory same. File already exists in the directory.");
                        return;
                    }

                    speaker.speak("File" + nme + "Copy Starting");
                    sourceFile = System.IO.Path.Combine(SourceDir, nme);
                    destFile = System.IO.Path.Combine(DesDir, nme);
                    if (!System.IO.Directory.Exists(DesDir))
                    {
                        System.IO.Directory.CreateDirectory(DesDir);
                    }
                    System.IO.File.Copy(sourceFile, destFile, true);
                    //settingBrowser();
                    speaker.speak("File Copy Completed");
                }
                else
                {
                    try
                    {
                        speaker.speak("copy starting");
                        Copy1(SourceDir + nme, DesDir + nme);

                        speaker.speak("Copy Done");

                        //speaker.speak("Folder" + nme + "Copy Starting");
                        //copyDirectory(SourceDir + nme, DesDir + nme);
                        //settingBrowser();
                        //speaker.speak("Folder Copy Completed");
                        //settingBrowser();
                    }
                    catch (Exception Ex)
                    {
                        Console.Error.WriteLine(Ex.Message);
                    }
                }
                //ctrl = false;
                SourceDir = null;
            }

            else
            {
                //ctrl = false;
                speaker.speak("Please Select a file or folder to copy");
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
        public void runDoc()
        {
            try
            {
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
            //thread = new Thread(new ThreadStart(runNarrator));
            //thread.Start();            
            FileInfo file = new FileInfo(fName);
            frmMain notepad = new frmMain(speaker);
            String extension = file.Extension.ToLower();

            if (extension.Equals(".txt") || extension.Equals(".rtf") || extension.Equals(".RTF"))
            {
                //Process[] procs = Process.GetProcessesByName("nvda");
                //if (procs.Length != 0)
                //{
                //    foreach (Process proc in procs)
                //        proc.Kill();
                //}
                //MessageBox.Show("Text");
                notepad.loadAFile(fName);
                notepad.Show();

            }
            else if (file.Extension.Equals(".doc") || file.Extension.Equals(".docx"))
            {
                FN = fName;
                Process.Start(FN);
                //thread = new Thread(new ThreadStart(runDoc));

                //thread.Start();
                //thread.Join();

            }
            else if (file.Extension.Equals(".xls") || file.Extension.Equals(".xlsx"))
            {

                FN = fName;
                Process.Start(FN);
                //thread = new Thread(new ThreadStart(runExcel));
                //thread.Start();
                //thread.Join();

            }
            else if (file.Extension.Equals(".ppt") || file.Extension.Equals(".pptx"))
            {
                FN = fName;
                Process.Start(FN);
                //thread = new Thread(new ThreadStart(runPPT));
                //thread.Start();
                //thread.Join();

            }
            else if (isAMediaFile(file.Extension))
            {
                PaglaPlayerPane player = new PaglaPlayerPane();
                player.startPlaying(getAllPlayList(fName));
                player.Show();
            }
            else if (file.Extension.Equals(".pdf") || file.Extension.Equals(".html") || file.Extension.Equals(".htm"))
            {
                //Process[] procs = Process.GetProcessesByName("nvda");
                //if (procs.Length != 0)
                //{
                //    foreach (Process proc in procs)
                //        proc.Kill();
                //}
                try
                {
                    //PDDocument d = PDDocument.load(fName);
                    //PDFTextStripper stripper = new PDFTextStripper();
                    //MessageBox.Show(stripper.getText(d));
                    //notepad.rtbDoc.Text = stripper.getText(d);
                    //notepad.Show();

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

    }
    /*end Browser class*/
}
