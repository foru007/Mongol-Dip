using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using SpeechBuilder;
using System.IO;
using EPocalipse.IFilter;
using System.Drawing;
using System.Text.RegularExpressions;

namespace DocForm
{
    class DocReadThread
    {
        Microsoft.Office.Interop.Word.Application word = null;
        Microsoft.Office.Interop.Word._Document doc = null;
        private static Microsoft.Office.Interop.Word.Application wd = null;

        static int k = 0;
        static int ct;
        private String keyData = null;
        private int occupiedBuffer = 0;
        private static int pre = 0;
        private static int pCoulmn = 0;
        private static int pRow = 0;

        private static string fulltext = "", shiftselecttext = "";

        Object unitChar = Word.WdUnits.wdCharacter;
        Object unitLine = Word.WdUnits.wdLine;
        Object unitWord = Word.WdUnits.wdWord;
        Object unitSentence = Word.WdUnits.wdSentence;
        Object unitParagraph = Word.WdUnits.wdParagraph;
        Object unitN = Word.WdUnits.wdLine;
        Object unitFullPage = Word.WdUnits.wdStory;

        object newTemplate = false;
        object docType = 0;
        object isVisible = true;
        static object p = null;
        private Form2 f2;

        Object count = 1;
        Object extend = Word.WdMovementType.wdMove;

        private SpeechControl speaker;

        public DocReadThread(SpeechControl speaker, Microsoft.Office.Interop.Word.Application word,
            Microsoft.Office.Interop.Word._Document doc, String keyData)
        {
            this.word = word;
            p = word;
            this.doc = doc;
            this.keyData = keyData;
            //speaker = new SpeechControl();
            this.speaker = speaker;
        }
        public void test()
        {
            try
            {
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    word.ActiveDocument.SpellingChecked = false;
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void PgLineNo()   //8
        {
            try
            {
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;

                    speaker.speak(" Page " + word.Selection.get_Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber).ToString());

                    String LN = "Line " + word.Selection.get_Information(Word.WdInformation.wdFirstCharacterLineNumber).ToString();
                    speaker.speak(LN.ToString());

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void PageNumber()   //5
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;

                    Word.Selection select = word.Selection;
                    Word.Selection rangeSelect = word.Selection;

                    select.HomeKey(ref unitN, ref extend);
                    Object start = select.Start;
                    select.EndKey(ref unitN, ref extend);
                    Object end = select.Start;

                    Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                    select.HomeKey(ref unitN, ref extend);


                    String s = rng.Text;
                    String output = " ";
                    //
                    for (int i = 0; i < s.Length; i++)
                    {

                        if (s[i].ToString() == "‘" || s[i].ToString() == "'" || s[i].ToString() == "’")
                            output = output + "Single quotation" + "\n";
                        else if (s[i].ToString() == "”" || s[i].ToString() == "“" || s[i].ToString() == "\"")
                            output = output + "Double quotes" + "\n";
                        else if (s[i].ToString() == "!")
                            output = output + "Exclamation mark" + "\n";
                        //else if (s[i].ToString()=="")
                        //    speaker.speak("Double quotes");
                        //else if (s[i].ToString() == """)
                        //    speaker.speak("Double quotes");
                        else if (s[i].ToString() == "(")
                            output = output + "first bracker Open" + "\n";
                        else if (s[i].ToString() == ")")
                            output = output + "first bracker Close" + "\n";
                        else if (s[i].ToString() == ",")
                            output = output + "Comma" + "\n";
                        else if (s[i].ToString() == "-")
                            output = output + "Hyphen" + "\n";
                        else if (s[i].ToString() == ".")
                            output = output + " full stop" + "\n";
                        else if (s[i].ToString() == ":")
                            output = output + "Colon" + "\n";
                        else if (s[i].ToString() == ";")
                            output = output + "Semicolon" + "\n";
                        else if (s[i].ToString() == "?")
                            output = output + "Question mark" + "\n";
                        else if (s[i].ToString() == "[")
                            output = output + "third bracket Open" + "\n";
                        else if (s[i].ToString() == "]")
                            output = output + "third bracket Close" + "\n";
                        else if (s[i].ToString() == "`")
                            output = output + "Grave accent" + "\n";
                        else if (s[i].ToString() == "{")
                            output = output + "Second bracket Open" + "\n";
                        else if (s[i].ToString() == "}")
                            output = output + "Second bracket Close" + "\n";
                        else if (s[i].ToString() == "~")
                            output = output + "Equivalency sign " + "\n";

                        else
                            output = output + s[i];
                    }
                    //

                    speaker.speak(output + " Page " + word.Selection.get_Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber).ToString());


                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void slectFirstCell()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                    {
                        Word.Selection select = word.Selection;
                        select.Tables[1].Cell(1, 1).Select();

                        String s = select.Text;
                        int count = 0;
                        for (int i = 0; i < s.Length; i++)
                        {
                            if (s[i] == ' ')
                                continue;
                            else
                                count = count + 1;
                        }
                        if (count == 2) speaker.speak("Blank");
                        else speaker.speak(select.Text.ToString());
                    }

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void slectLastCell()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                    {
                        Word.Selection select = word.Selection;
                        int x = select.Tables[1].Rows.Count;
                        int y = select.Tables[1].Columns.Count;
                        select.Tables[1].Cell(x, y).Select();

                        String s = select.Text;
                        int count = 0;
                        for (int i = 0; i < s.Length; i++)
                        {
                            if (s[i] == ' ')
                                continue;
                            else
                                count = count + 1;
                        }
                        if (count == 2) speaker.speak("Blank");
                        else speaker.speak(select.Text.ToString());
                    }

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void columnTitle()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    Word.Selection select = word.Selection;
                    int y = select.Columns.First.Index;
                    speaker.speak(select.Tables[1].Cell(1, y).Range.Text);
                    occupiedBuffer = 0;
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void rowTitle()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    Word.Selection select = word.Selection;
                    int x = select.Rows.First.Index;
                    speaker.speak(select.Tables[1].Cell(x, 1).Range.Text);
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void slectPreviousAndNextRow()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    if (keyData.ToString().Equals("Up"))
                    {
                        if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                        {
                            Word.Selection select = word.Selection;
                            int x = select.Rows.First.Index - 1;
                            int y = select.Columns.First.Index;
                            select.Tables[1].Cell(x, y).Select();
                            String ss = select.Text.ToString();
                            if (ss.Length == 2)
                                speaker.speak("Blank");
                            else
                                speaker.speak(ss);

                        }

                    }
                    else if (keyData.ToString().Equals("Down"))
                    {
                        Word.Selection select = word.Selection;
                        int x = select.Rows.First.Index + 1;
                        int y = select.Columns.First.Index;
                        select.Tables[1].Cell(x, y).Select();
                        String ss = select.Text.ToString();
                        if (ss.Length == 2)
                            speaker.speak("Blank");
                        else
                            speaker.speak(ss);
                    }
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void slectPreviousAndNextCell()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    if (keyData.ToString().Equals("Right"))
                    {
                        if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                        {
                            Word.Selection select = word.Selection;
                            int x = select.Rows.First.Index;
                            int y = select.Columns.First.Index;
                            select.Tables[1].Cell(x, y).Next.Select();
                            String ss = select.Text.ToString();
                            if (ss.Length == 2)
                                speaker.speak("Blank");
                            else
                                speaker.speak(ss);


                        }

                    }
                    else if (keyData.ToString().Equals("Left"))
                    {
                        Word.Selection select = word.Selection;
                        int x = select.Rows.First.Index;
                        int y = select.Columns.First.Index;
                        select.Tables[1].Cell(x, y).Previous.Select();
                        String ss = select.Text.ToString();
                        //MessageBox.Show(ss);
                        if (ss.Length == 2)
                            speaker.speak("Blank");
                        else
                            speaker.speak(ss);
                    }
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void operatePriorSentence()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                object pos = 1;
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    if (keyData.ToString().Equals("Down"))
                    {
                        Word.Selection select = word.Selection;
                        select.MoveStart(ref unitSentence, ref pos);
                        String text = select.Sentences.First.Text;
                        speaker.speak(text);
                    }
                    else if (keyData.ToString().Equals("Up"))
                    {
                        object pos1 = -1;
                        Word.Selection select = word.Selection;
                        select.StartOf(ref unitSentence, ref extend);
                        select.Move(ref unitSentence, ref pos1);
                        String text = select.Sentences.First.Text;
                        speaker.speak(text);
                    }
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void PageSetUp()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;

                    ///////////////////////////////////////////////////////////////////////////////////////////////////////orientation

                    String Page_Orientation = word.ActiveDocument.PageSetup.Orientation.ToString();

                    if (Page_Orientation == "wdOrientLandscape")
                    {
                        speaker.speak("page orientation Landscape ");
                    }
                    else if (Page_Orientation == "wdOrientPortrait")
                    {
                        speaker.speak("page orientation Portrait ");
                    }
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////margin

                    float lm = word.ActiveDocument.PageSetup.LeftMargin, rm = word.ActiveDocument.PageSetup.RightMargin;
                    float tm = word.ActiveDocument.PageSetup.TopMargin, bm = word.ActiveDocument.PageSetup.BottomMargin;
                    lm = lm / 72;
                    rm = rm / 72;
                    tm = tm / 72;
                    bm = bm / 72;

                    speaker.speak(" page margin left " + lm + " inches, right " + rm + " inches, top " + tm + " inches, bottom " + bm + " inches ");

                    /////////////////////////////////////////////////////////////////////////////////////////////////////////
                    int count_table = word.ActiveDocument.Tables.Count;
                    speaker.speak(" number of table in document is " + count_table.ToString());

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void Instraction()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;

                    Word.Selection select = word.Selection;
                    //speaker.speak("Page Number is " + word.Selection.get_Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber).ToString());
                    //speaker.speak("Now Font Size is " + select.Font.Size.ToString());
                    //speaker.speak("Now Font Name is " + select.Font.Name.ToString());
                    String w = select.Font.Color.ToString(), paragraphAllignment = "";   // font color
                    int bold;
                    float linespace;

                    //////////////////////////////////////////////////////////////
                    if ((keyData.ToString().Equals("F")))
                    {
                        bold = select.Font.Bold;
                        if (bold == -1) speaker.speak(" bolded ");

                        speaker.speak(select.Font.Size.ToString() + " point ");

                        //MessageBox.Show(select.Font.ColorIndex.ToString());

                        // font color
                        if (w == "-603914241")
                            speaker.speak(" White");
                        else if (w == "-603917569")
                            speaker.speak(" White Darker 5 percentage");
                        else if (w == "-603923969")
                            speaker.speak(" White Darker 15 percentage");
                        else if (w == "-603930625")
                            speaker.speak(" White Darker 25 percentage");
                        else if (w == "-603937025")
                            speaker.speak(" White Darker 35 percentage");
                        else if (w == "-603946753")
                            speaker.speak(" White Darker 50 percentage");

                        else if (w == "-587137025")
                            speaker.speak(" Black");
                        else if (w == "-587137152")
                            speaker.speak(" Black Lighter 50 percentage");
                        else if (w == "-587137114")
                            speaker.speak(" Black Lighter 35 percentage");
                        else if (w == "-587137089")
                            speaker.speak(" Black Lighter 25 percentage");
                        else if (w == "-587137063")
                            speaker.speak(" Black Lighter 15 percentage");
                        else if (w == "-587137038")
                            speaker.speak(" Black Lighter 5 percentage");

                        else if (w == "-570359809")
                            speaker.speak(" Tan");
                        else if (w == "-570366209")
                            speaker.speak(" Tan Darker 10 percentage");
                        else if (w == "-570376193")
                            speaker.speak(" Tan Darker 25 percentage");
                        else if (w == "-570392321")
                            speaker.speak(" Tan Darker 50 percentage");
                        else if (w == "-570408705")
                            speaker.speak(" Tan Darker 75 percentage");
                        else if (w == "-570418433")
                            speaker.speak(" Tan Darker 90 percentage");

                        else if (w == "-553582593")
                            speaker.speak(" Dark Blue");
                        else if (w == "-553582797")
                            speaker.speak(" Dark Blue Lighter 80 percentage");
                        else if (w == "-553582746")
                            speaker.speak(" Dark Blue Lighter 60 percentage");
                        else if (w == "-553582695")
                            speaker.speak(" Dark Blue Lighter 40 percentage");
                        else if (w == "-553598977")
                            speaker.speak(" Dark Blue Darker 25 percentage");
                        else if (w == "-553615105")
                            speaker.speak(" Dark Blue Darker 50 percentage");

                        else if (w == "-738131969")
                            speaker.speak(" Blue");
                        else if (w == "-738132173")
                            speaker.speak(" Blue Lighter 80 percentage");
                        else if (w == "-738132122")
                            speaker.speak(" Blue Lighter 60 percentage");
                        else if (w == "-738132071")
                            speaker.speak(" Blue Lighter 40 percentage");
                        else if (w == "-738148353")
                            speaker.speak(" Blue Darker 25 percentage");
                        else if (w == "-738164481")
                            speaker.speak(" Blue Darker 50 percentage");


                        else if (w == "-721354753")
                            speaker.speak(" Red");
                        else if (w == "-721354957")
                            speaker.speak(" Red Lighter 80 percentage");
                        else if (w == "-721354906")
                            speaker.speak(" Red Lighter 60 percentage");
                        else if (w == "-721354855")
                            speaker.speak(" Red Lighter 40 percentage");
                        else if (w == "-721371137")
                            speaker.speak(" Red Darker 25 percentage");
                        else if (w == "-721387265")
                            speaker.speak(" Red Darker 50 percentage");


                        else if (w == "-704577537")
                            speaker.speak(" Olive Green");
                        else if (w == "-704577741")
                            speaker.speak(" Olive Green Lighter 80 percentage");
                        else if (w == "-704577690")
                            speaker.speak(" Olive Green Lighter 60 percentage");
                        else if (w == "-704577639")
                            speaker.speak(" Olive Green Lighter 40 percentage");
                        else if (w == "-704593921")
                            speaker.speak(" Olive Green Darker 25 percentage");
                        else if (w == "-704610049")
                            speaker.speak(" Olive Green Darker 50 percentage");


                        else if (w == "-687800321")
                            speaker.speak(" Purple");
                        else if (w == "-687800525")
                            speaker.speak(" Purple Lighter 80 percentage");
                        else if (w == "-687800474")
                            speaker.speak(" Purple Lighter 60 percentage");
                        else if (w == "-687800423")
                            speaker.speak(" Purple Lighter 40 percentage");
                        else if (w == "-687816705")
                            speaker.speak(" Purple Darker 25 percentage");
                        else if (w == "-687832833")
                            speaker.speak(" Purple Darker 50 percentage");

                        else if (w == "-671023105")
                            speaker.speak(" Aqua");
                        else if (w == "-671023309")
                            speaker.speak(" Aqua Lighter 80 percentage");
                        else if (w == "-671023258")
                            speaker.speak(" Aqua Lighter 60 percentage");
                        else if (w == "-671023207")
                            speaker.speak(" Aqua Lighter 40 percentage");
                        else if (w == "-671039489")
                            speaker.speak(" Aqua Darker 25 percentage");
                        else if (w == "-671055617")
                            speaker.speak(" Aqua Darker 50 percentage");

                        else if (w == "-654245889")
                            speaker.speak(" Orange");
                        else if (w == "-654246093")
                            speaker.speak(" Orange Lighter 80 percentage");
                        else if (w == "-654246042")
                            speaker.speak(" Orange Lighter 60 percentage");
                        else if (w == "-654245991")
                            speaker.speak(" Orange Lighter 40 percentage");
                        else if (w == "-654262273")
                            speaker.speak(" Orange Darker 25 percentage");
                        else if (w == "-654278401")
                            speaker.speak(" Orange Darker 50 percentage");


                        else if (w == "192")
                            speaker.speak(" Standard Dark Red");
                        else if (w == "wdColorRed")
                            speaker.speak(" Standard Red");
                        else if (w == "49407")
                            speaker.speak(" Standard Orange");
                        else if (w == "wdColorYellow")
                            speaker.speak(" Standard Yellow");
                        else if (w == "5296274")
                            speaker.speak(" Standard Light Green");
                        else if (w == "5287936")
                            speaker.speak(" Standard Green");
                        else if (w == "15773696")
                            speaker.speak(" Standard Light Blue");
                        else if (w == "12611584")
                            speaker.speak(" Standard Blue");
                        else if (w == "6299648")
                            speaker.speak(" Standard Dark Blue");
                        else if (w == "10498160")
                            speaker.speak(" Standard Purple");
                        else
                            speaker.speak("No Highlighted Color");
                        //////
                        speaker.speak("  " + select.Font.Name.ToString());

                        if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                        {
                            speaker.speak(" table print style ");
                        }
                        else
                        {
                            speaker.speak(" normal style ");
                        }


                        linespace = select.ParagraphFormat.LineSpacing;
                        if (linespace == 12) speaker.speak(" line spaceing colon single ");
                        else if (linespace == 24) speaker.speak(" line spaceing colon double");
                        else if (linespace == 18) speaker.speak(" line spaceing colon 1.5 ");
                        else if (linespace == 36) speaker.speak(" line spaceing colon multiply by 3 ");
                        else speaker.speak(" line spaceing colon exactly " + linespace + " points ");

                        ////////
                        String listFormat = null;
                        if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListSimpleNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                            speaker.speak(" paragraph formating colon simple Numbering");
                            speaker.speak(" outline colon level body text");
                            return;
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListBullet)
                        {
                            listFormat = "Bullet";
                            speaker.speak(" paragraph formating colon bullet list  ");
                            speaker.speak(" outline level colon body text");
                            return;
                        }
                        /////

                        paragraphAllignment = select.ParagraphFormat.Alignment.ToString();
                        if (paragraphAllignment == "wdAlignParagraphLeft")
                            speaker.speak(" paragraph formating colon align left ");
                        else if (paragraphAllignment == "wdAlignParagraphRight")
                            speaker.speak(" paragraph formating colon align Right ");
                        else if (paragraphAllignment == "wdAlignParagraphJustify")
                            speaker.speak(" paragraph formating colon justify ");
                        else if (paragraphAllignment == "wdAlignParagraphCenter")
                            speaker.speak(" paragraph formating colon align center ");
                        //MessageBox.Show(paragraphAllignment);


                        speaker.speak(" outline level colon body text");

                    }

                    else if ((keyData.ToString().Equals("D5")))
                    {
                        // font color
                        if (w == "-603914241")
                            speaker.speak(" White");
                        else if (w == "-603917569")
                            speaker.speak(" White Darker 5 percentage");
                        else if (w == "-603923969")
                            speaker.speak(" White Darker 15 percentage");
                        else if (w == "-603930625")
                            speaker.speak(" White Darker 25 percentage");
                        else if (w == "-603937025")
                            speaker.speak(" White Darker 35 percentage");
                        else if (w == "-603946753")
                            speaker.speak(" White Darker 50 percentage");

                        else if (w == "-587137025")
                            speaker.speak(" Black");
                        else if (w == "-587137152")
                            speaker.speak(" Black Lighter 50 percentage");
                        else if (w == "-587137114")
                            speaker.speak(" Black Lighter 35 percentage");
                        else if (w == "-587137089")
                            speaker.speak(" Black Lighter 25 percentage");
                        else if (w == "-587137063")
                            speaker.speak(" Black Lighter 15 percentage");
                        else if (w == "-587137038")
                            speaker.speak(" Black Lighter 5 percentage");

                        else if (w == "-570359809")
                            speaker.speak(" Tan");
                        else if (w == "-570366209")
                            speaker.speak(" Tan Darker 10 percentage");
                        else if (w == "-570376193")
                            speaker.speak(" Tan Darker 25 percentage");
                        else if (w == "-570392321")
                            speaker.speak(" Tan Darker 50 percentage");
                        else if (w == "-570408705")
                            speaker.speak(" Tan Darker 75 percentage");
                        else if (w == "-570418433")
                            speaker.speak(" Tan Darker 90 percentage");

                        else if (w == "-553582593")
                            speaker.speak(" Dark Blue");
                        else if (w == "-553582797")
                            speaker.speak(" Dark Blue Lighter 80 percentage");
                        else if (w == "-553582746")
                            speaker.speak(" Dark Blue Lighter 60 percentage");
                        else if (w == "-553582695")
                            speaker.speak(" Dark Blue Lighter 40 percentage");
                        else if (w == "-553598977")
                            speaker.speak(" Dark Blue Darker 25 percentage");
                        else if (w == "-553615105")
                            speaker.speak(" Dark Blue Darker 50 percentage");

                        else if (w == "-738131969")
                            speaker.speak(" Blue");
                        else if (w == "-738132173")
                            speaker.speak(" Blue Lighter 80 percentage");
                        else if (w == "-738132122")
                            speaker.speak(" Blue Lighter 60 percentage");
                        else if (w == "-738132071")
                            speaker.speak(" Blue Lighter 40 percentage");
                        else if (w == "-738148353")
                            speaker.speak(" Blue Darker 25 percentage");
                        else if (w == "-738164481")
                            speaker.speak(" Blue Darker 50 percentage");


                        else if (w == "-721354753")
                            speaker.speak(" Red");
                        else if (w == "-721354957")
                            speaker.speak(" Red Lighter 80 percentage");
                        else if (w == "-721354906")
                            speaker.speak(" Red Lighter 60 percentage");
                        else if (w == "-721354855")
                            speaker.speak(" Red Lighter 40 percentage");
                        else if (w == "-721371137")
                            speaker.speak(" Red Darker 25 percentage");
                        else if (w == "-721387265")
                            speaker.speak(" Red Darker 50 percentage");


                        else if (w == "-704577537")
                            speaker.speak(" Olive Green");
                        else if (w == "-704577741")
                            speaker.speak(" Olive Green Lighter 80 percentage");
                        else if (w == "-704577690")
                            speaker.speak(" Olive Green Lighter 60 percentage");
                        else if (w == "-704577639")
                            speaker.speak(" Olive Green Lighter 40 percentage");
                        else if (w == "-704593921")
                            speaker.speak(" Olive Green Darker 25 percentage");
                        else if (w == "-704610049")
                            speaker.speak(" Olive Green Darker 50 percentage");


                        else if (w == "-687800321")
                            speaker.speak(" Purple");
                        else if (w == "-687800525")
                            speaker.speak(" Purple Lighter 80 percentage");
                        else if (w == "-687800474")
                            speaker.speak(" Purple Lighter 60 percentage");
                        else if (w == "-687800423")
                            speaker.speak(" Purple Lighter 40 percentage");
                        else if (w == "-687816705")
                            speaker.speak(" Purple Darker 25 percentage");
                        else if (w == "-687832833")
                            speaker.speak(" Purple Darker 50 percentage");

                        else if (w == "-671023105")
                            speaker.speak(" Aqua");
                        else if (w == "-671023309")
                            speaker.speak(" Aqua Lighter 80 percentage");
                        else if (w == "-671023258")
                            speaker.speak(" Aqua Lighter 60 percentage");
                        else if (w == "-671023207")
                            speaker.speak(" Aqua Lighter 40 percentage");
                        else if (w == "-671039489")
                            speaker.speak(" Aqua Darker 25 percentage");
                        else if (w == "-671055617")
                            speaker.speak(" Aqua Darker 50 percentage");

                        else if (w == "-654245889")
                            speaker.speak(" Orange");
                        else if (w == "-654246093")
                            speaker.speak(" Orange Lighter 80 percentage");
                        else if (w == "-654246042")
                            speaker.speak(" Orange Lighter 60 percentage");
                        else if (w == "-654245991")
                            speaker.speak(" Orange Lighter 40 percentage");
                        else if (w == "-654262273")
                            speaker.speak(" Orange Darker 25 percentage");
                        else if (w == "-654278401")
                            speaker.speak(" Orange Darker 50 percentage");


                        else if (w == "192")
                            speaker.speak(" Standard Dark Red");
                        else if (w == "wdColorRed")
                            speaker.speak(" Standard Red");
                        else if (w == "49407")
                            speaker.speak(" Standard Orange");
                        else if (w == "wdColorYellow")
                            speaker.speak(" Standard Yellow");
                        else if (w == "5296274")
                            speaker.speak(" Standard Light Green");
                        else if (w == "5287936")
                            speaker.speak(" Standard Green");
                        else if (w == "15773696")
                            speaker.speak(" Standard Light Blue");
                        else if (w == "12611584")
                            speaker.speak(" Standard Blue");
                        else if (w == "6299648")
                            speaker.speak(" Standard Dark Blue");
                        else if (w == "10498160")
                            speaker.speak(" Standard Purple");
                        else
                            speaker.speak("No Highlighted Color");

                        string s = select.Document.Background.Fill.ForeColor.RGB.ToString();

                        //background color
                        if (s == "16777215")
                            speaker.speak(" on White");
                        else if (s == "15921906")
                            speaker.speak(" on White Darker 5 percentage");
                        else if (s == "14211288")
                            speaker.speak(" on White Darker 15 percentage");
                        else if (s == "12566463")
                            speaker.speak(" on White Darker 25 percentage");
                        else if (s == "10855845")
                            speaker.speak(" on White Darker 35 percentage");
                        else if (s == "8355711")
                            speaker.speak(" on White Darker 50 percentage");

                        else if (s == "0")
                            speaker.speak(" on Black");
                        else if (s == "8421504")
                            speaker.speak(" on Black Lighter 50 percentage");
                        else if (s == "5921370")
                            speaker.speak(" on Black Lighter 35 percentage");
                        else if (s == "4210752")
                            speaker.speak(" on Black Lighter 25 percentage");
                        else if (s == "2565927")
                            speaker.speak(" on Black Lighter 15 percentage");
                        else if (s == "855309")
                            speaker.speak(" on Black Lighter 5 percentage");

                        else if (s == "14806254")
                            speaker.speak(" on Tan");
                        else if (s == "12769501")
                            speaker.speak(" on Tan Darker 10 percentage");
                        else if (s == "9878724")
                            speaker.speak(" on Tan Darker 25 percentage");
                        else if (s == "5474707")
                            speaker.speak(" on Tan Darker 50 percentage");
                        else if (s == "2704200")
                            speaker.speak(" on Tan Darker 75 percentage");
                        else if (s == "1055260")
                            speaker.speak(" on Tan Darker 90 percentage");

                        else if (s == "8210719")
                            speaker.speak(" on Dark Blue");
                        else if (s == "15849926")
                            speaker.speak(" on Dark Blue Lighter 80 percentage");
                        else if (s == "148557101")
                            speaker.speak(" on Dark Blue Lighter 60 percentage");
                        else if (s == "13929812")
                            speaker.speak(" on Dark Blue Lighter 40 percentage");
                        else if (s == "6108695")
                            speaker.speak(" on Dark Blue Darker 25 percentage");
                        else if (s == "4072463")
                            speaker.speak(" on Dark Blue Darker 50 percentage");

                        else if (s == "12419407")
                            speaker.speak(" on Blue");
                        else if (s == "15853019")
                            speaker.speak(" on Blue Lighter 80 percentage");
                        else if (s == "14994616")
                            speaker.speak(" on Blue Lighter 60 percentage");
                        else if (s == "14136213")
                            speaker.speak(" on Blue Lighter 40 percentage");
                        else if (s == "9527094")
                            speaker.speak(" on Blue Darker 25 percentage");
                        else if (s == "6307620")
                            speaker.speak(" on Blue Darker 50 percentage");


                        else if (s == "5066944")
                            speaker.speak(" on Red");
                        else if (s == "14408690")
                            speaker.speak(" on Red Lighter 80 percentage");
                        else if (s == "12040421")
                            speaker.speak(" on Red Lighter 60 percentage");
                        else if (s == "9737689")
                            speaker.speak(" on Red Lighter 40 percentage");
                        else if (s == "3421844")
                            speaker.speak(" on Red Darker 25 percentage");
                        else if (s == "2303074")
                            speaker.speak(" on Red Darker 50 percentage");


                        else if (s == "5880731")
                            speaker.speak(" on Olive Green");
                        else if (s == "14545386")
                            speaker.speak(" on Olive Green Lighter 80 percentage");
                        else if (s == "12379094")
                            speaker.speak(" on Olive Green Lighter 60 percentage");
                        else if (s == "10213058")
                            speaker.speak(" on Olive Green Lighter 40 percentage");
                        else if (s == "3969654")
                            speaker.speak(" on Olive Green Darker 25 percentage");
                        else if (s == "2646350")
                            speaker.speak(" on Olive Green Darker 50 percentage");


                        else if (s == "10642560")
                            speaker.speak(" on Purple");
                        else if (s == "15523813")
                            speaker.speak(" on Purple Lighter 80 percentage");
                        else if (s == "14270668")
                            speaker.speak(" on Purple Lighter 60 percentage");
                        else if (s == "13083058")
                            speaker.speak(" on Purple Lighter 40 percentage");
                        else if (s == "8014175")
                            speaker.speak(" on Purple Darker 25 percentage");
                        else if (s == "5321023")
                            speaker.speak(" on Purple Darker 50 percentage");

                        else if (s == "13020235")
                            speaker.speak(" on Aqua");
                        else if (s == "15986394")
                            speaker.speak(" on Aqua Lighter 80 percentage");
                        else if (s == "15261110")
                            speaker.speak(" on Aqua Lighter 60 percentage");
                        else if (s == "14470546")
                            speaker.speak(" on Aqua Lighter 40 percentage");
                        else if (s == "10191921")
                            speaker.speak(" on Aqua Darker 25 percentage");
                        else if (s == "6772768")
                            speaker.speak(" on Aqua Darker 50 percentage");

                        else if (s == "4626167")
                            speaker.speak(" on Orange");
                        else if (s == "14281213")
                            speaker.speak(" on Orange Lighter 80 percentage");
                        else if (s == "11851003")
                            speaker.speak(" on Orange Lighter 60 percentage");
                        else if (s == "9420794")
                            speaker.speak(" on Orange Lighter 40 percentage");
                        else if (s == "683235")
                            speaker.speak(" on Orange Darker 25 percentage");
                        else if (s == "411543")
                            speaker.speak(" on Orange Darker 50 percentage");


                        else if (s == "192")
                            speaker.speak(" on Standard Dark Red");
                        else if (s == "255")
                            speaker.speak(" on Standard Red");
                        else if (s == "49407")
                            speaker.speak(" on Standard Orange");
                        else if (s == "65535")
                            speaker.speak(" on Standard Yellow");
                        else if (s == "5296274")
                            speaker.speak(" on Standard Light Green");
                        else if (s == "5287936")
                            speaker.speak(" on Standard Green");
                        else if (s == "15773696")
                            speaker.speak(" on Standard Light Blue");
                        else if (s == "12611584")
                            speaker.speak(" on Standard Blue");
                        else if (s == "6299648")
                            speaker.speak(" on Standard Dark Blue");
                        else if (s == "10498160")
                            speaker.speak(" on Standard Purple");
                        else
                            speaker.speak("on No Background Color");

                    }
                    /////////////////////////////////////////////////////////////

                    //if (w == "-603914241")
                    //    speaker.speak("Font Color is White");
                    //else if (w == "-603917569")
                    //    speaker.speak("Font Color is White Darker 5 percentage");
                    //else if (w == "-603923969")
                    //    speaker.speak("Font Color is White Darker 15 percentage");
                    //else if (w == "-603930625")
                    //    speaker.speak("Font Color is White Darker 25 percentage");
                    //else if (w == "-603937025")
                    //    speaker.speak("Font Color is White Darker 35 percentage");
                    //else if (w == "-603946753")
                    //    speaker.speak("Font Color is White Darker 50 percentage");

                    //else if (w == "-587137025")
                    //    speaker.speak("Font Color is Black");
                    //else if (w == "-587137152")
                    //    speaker.speak("Font Color is Black Lighter 50 percentage");
                    //else if (w == "-587137114")
                    //    speaker.speak("Font Color is Black Lighter 35 percentage");
                    //else if (w == "-587137089")
                    //    speaker.speak("Font Color is Black Lighter 25 percentage");
                    //else if (w == "-587137063")
                    //    speaker.speak("Font Color is Black Lighter 15 percentage");
                    //else if (w == "-587137038")
                    //    speaker.speak("Font Color is Black Lighter 5 percentage");

                    //else if (w == "-570359809")
                    //    speaker.speak("Font Color is Tan");
                    //else if (w == "-570366209")
                    //    speaker.speak("Font Color is Tan Darker 10 percentage");
                    //else if (w == "-570376193")
                    //    speaker.speak("Font Color is Tan Darker 25 percentage");
                    //else if (w == "-570392321")
                    //    speaker.speak("Font Color is Tan Darker 50 percentage");
                    //else if (w == "-570408705")
                    //    speaker.speak("Font Color is Tan Darker 75 percentage");
                    //else if (w == "-570418433")
                    //    speaker.speak("Font Color is Tan Darker 90 percentage");

                    //else if (w == "-553582593")
                    //    speaker.speak("Font Color is Dark Blue");
                    //else if (w == "-553582797")
                    //    speaker.speak("Font Color is Dark Blue Lighter 80 percentage");
                    //else if (w == "-553582746")
                    //    speaker.speak("Font Color is Dark Blue Lighter 60 percentage");
                    //else if (w == "-553582695")
                    //    speaker.speak("Font Color is Dark Blue Lighter 40 percentage");
                    //else if (w == "-553598977")
                    //    speaker.speak("Font Color is Dark Blue Darker 25 percentage");
                    //else if (w == "-553615105")
                    //    speaker.speak("Font Color is Dark Blue Darker 50 percentage");

                    //else if (w == "-738131969")
                    //    speaker.speak("Font Color is Blue");
                    //else if (w == "-738132173")
                    //    speaker.speak("Font Color is Blue Lighter 80 percentage");
                    //else if (w == "-738132122")
                    //    speaker.speak("Font Color is Blue Lighter 60 percentage");
                    //else if (w == "-738132071")
                    //    speaker.speak("Font Color is Blue Lighter 40 percentage");
                    //else if (w == "-738148353")
                    //    speaker.speak("Font Color is Blue Darker 25 percentage");
                    //else if (w == "-738164481")
                    //    speaker.speak("Font Color is Blue Darker 50 percentage");


                    //else if (w == "-721354753")
                    //    speaker.speak("Font Color is Red");
                    //else if (w == "-721354957")
                    //    speaker.speak("Font Color is Red Lighter 80 percentage");
                    //else if (w == "-721354906")
                    //    speaker.speak("Font Color is Red Lighter 60 percentage");
                    //else if (w == "-721354855")
                    //    speaker.speak("Font Color is Red Lighter 40 percentage");
                    //else if (w == "-721371137")
                    //    speaker.speak("Font Color is Red Darker 25 percentage");
                    //else if (w == "-721387265")
                    //    speaker.speak("Font Color is Red Darker 50 percentage");


                    //else if (w == "-704577537")
                    //    speaker.speak("Font Color is Olive Green");
                    //else if (w == "-704577741")
                    //    speaker.speak("Font Color is Olive Green Lighter 80 percentage");
                    //else if (w == "-704577690")
                    //    speaker.speak("Font Color is Olive Green Lighter 60 percentage");
                    //else if (w == "-704577639")
                    //    speaker.speak("Font Color is Olive Green Lighter 40 percentage");
                    //else if (w == "-704593921")
                    //    speaker.speak("Font Color is Olive Green Darker 25 percentage");
                    //else if (w == "-704610049")
                    //    speaker.speak("Font Color is Olive Green Darker 50 percentage");


                    //else if (w == "-687800321")
                    //    speaker.speak("Font Color is Purple");
                    //else if (w == "-687800525")
                    //    speaker.speak("Font Color is Purple Lighter 80 percentage");
                    //else if (w == "-687800474")
                    //    speaker.speak("Font Color is Purple Lighter 60 percentage");
                    //else if (w == "-687800423")
                    //    speaker.speak("Font Color is Purple Lighter 40 percentage");
                    //else if (w == "-687816705")
                    //    speaker.speak("Font Color is Purple Darker 25 percentage");
                    //else if (w == "-687832833")
                    //    speaker.speak("Font Color is Purple Darker 50 percentage");

                    //else if (w == "-671023105")
                    //    speaker.speak("Font Color is Aqua");
                    //else if (w == "-671023309")
                    //    speaker.speak("Font Color is Aqua Lighter 80 percentage");
                    //else if (w == "-671023258")
                    //    speaker.speak("Font Color is Aqua Lighter 60 percentage");
                    //else if (w == "-671023207")
                    //    speaker.speak("Font Color is Aqua Lighter 40 percentage");
                    //else if (w == "-671039489")
                    //    speaker.speak("Font Color is Aqua Darker 25 percentage");
                    //else if (w == "-671055617")
                    //    speaker.speak("Font Color is Aqua Darker 50 percentage");

                    //else if (w == "-654245889")
                    //    speaker.speak("Font Color is Orange");
                    //else if (w == "-654246093")
                    //    speaker.speak("Font Color is Orange Lighter 80 percentage");
                    //else if (w == "-654246042")
                    //    speaker.speak("Font Color is Orange Lighter 60 percentage");
                    //else if (w == "-654245991")
                    //    speaker.speak("Font Color is Orange Lighter 40 percentage");
                    //else if (w == "-654262273")
                    //    speaker.speak("Font Color is Orange Darker 25 percentage");
                    //else if (w == "-654278401")
                    //    speaker.speak("Font Color is Orange Darker 50 percentage");


                    //else if (w == "192")
                    //    speaker.speak("Font Color is Standard Dark Red");
                    //else if (w == "wdColorRed")
                    //    speaker.speak("Font Color is Standard Red");
                    //else if (w == "49407")
                    //    speaker.speak("Font Color is Standard Orange");
                    //else if (w == "wdColorYellow")
                    //    speaker.speak("Font Color is Standard Yellow");
                    //else if (w == "5296274")
                    //    speaker.speak("Font Color is Standard Light Green");
                    //else if (w == "5287936")
                    //    speaker.speak("Font Color is Standard Green");
                    //else if (w == "15773696")
                    //    speaker.speak("Font Color is Standard Light Blue");
                    //else if (w == "12611584")
                    //    speaker.speak("Font Color is Standard Blue");
                    //else if (w == "6299648")
                    //    speaker.speak("Font Color is Standard Dark Blue");
                    //else if (w == "10498160")
                    //    speaker.speak("Font Color is Standard Purple");
                    //else
                    //    speaker.speak("No Highlighted Color");

                    //string s = select.Document.Background.Fill.ForeColor.RGB.ToString();

                    //if (s == "16777215")
                    //    speaker.speak("Background Color is White");
                    //else if (s == "15921906")
                    //    speaker.speak("Background Color is White Darker 5 percentage");
                    //else if (s == "14211288")
                    //    speaker.speak("Background Color is White Darker 15 percentage");
                    //else if (s == "12566463")
                    //    speaker.speak("Background Color is White Darker 25 percentage");
                    //else if (s == "10855845")
                    //    speaker.speak("Background Color is White Darker 35 percentage");
                    //else if (s == "8355711")
                    //    speaker.speak("Background Color is White Darker 50 percentage");

                    //else if (s == "0")
                    //    speaker.speak("Background Color is Black");
                    //else if (s == "8421504")
                    //    speaker.speak("Background Color is Black Lighter 50 percentage");
                    //else if (s == "5921370")
                    //    speaker.speak("Background Color is Black Lighter 35 percentage");
                    //else if (s == "4210752")
                    //    speaker.speak("Background Color is Black Lighter 25 percentage");
                    //else if (s == "2565927")
                    //    speaker.speak("Background Color is Black Lighter 15 percentage");
                    //else if (s == "855309")
                    //    speaker.speak("Background Color is Black Lighter 5 percentage");

                    //else if (s == "14806254")
                    //    speaker.speak("Background Color is Tan");
                    //else if (s == "12769501")
                    //    speaker.speak("Background Color is Tan Darker 10 percentage");
                    //else if (s == "9878724")
                    //    speaker.speak("Background Color is Tan Darker 25 percentage");
                    //else if (s == "5474707")
                    //    speaker.speak("Background Color is Tan Darker 50 percentage");
                    //else if (s == "2704200")
                    //    speaker.speak("Background Color is Tan Darker 75 percentage");
                    //else if (s == "1055260")
                    //    speaker.speak("Background Color is Tan Darker 90 percentage");

                    //else if (s == "8210719")
                    //    speaker.speak("Background Color is Dark Blue");
                    //else if (s == "15849926")
                    //    speaker.speak("Background Color is Dark Blue Lighter 80 percentage");
                    //else if (s == "148557101")
                    //    speaker.speak("Background Color is Dark Blue Lighter 60 percentage");
                    //else if (s == "13929812")
                    //    speaker.speak("Background Color is Dark Blue Lighter 40 percentage");
                    //else if (s == "6108695")
                    //    speaker.speak("Background Color is Dark Blue Darker 25 percentage");
                    //else if (s == "4072463")
                    //    speaker.speak("Background Color is Dark Blue Darker 50 percentage");

                    //else if (s == "12419407")
                    //    speaker.speak("Background Color is Blue");
                    //else if (s == "15853019")
                    //    speaker.speak("Background Color is Blue Lighter 80 percentage");
                    //else if (s == "14994616")
                    //    speaker.speak("Background Color is Blue Lighter 60 percentage");
                    //else if (s == "14136213")
                    //    speaker.speak("Background Color is Blue Lighter 40 percentage");
                    //else if (s == "9527094")
                    //    speaker.speak("Background Color is Blue Darker 25 percentage");
                    //else if (s == "6307620")
                    //    speaker.speak("Background Color is Blue Darker 50 percentage");


                    //else if (s == "5066944")
                    //    speaker.speak("Background Color is Red");
                    //else if (s == "14408690")
                    //    speaker.speak("Background Color is Red Lighter 80 percentage");
                    //else if (s == "12040421")
                    //    speaker.speak("Background Color is Red Lighter 60 percentage");
                    //else if (s == "9737689")
                    //    speaker.speak("Background Color is Red Lighter 40 percentage");
                    //else if (s == "3421844")
                    //    speaker.speak("Background Color is Red Darker 25 percentage");
                    //else if (s == "2303074")
                    //    speaker.speak("Background Color is Red Darker 50 percentage");


                    //else if (s == "5880731")
                    //    speaker.speak("Background Color is Olive Green");
                    //else if (s == "14545386")
                    //    speaker.speak("Background Color is Olive Green Lighter 80 percentage");
                    //else if (s == "12379094")
                    //    speaker.speak("Background Color is Olive Green Lighter 60 percentage");
                    //else if (s == "10213058")
                    //    speaker.speak("Background Color is Olive Green Lighter 40 percentage");
                    //else if (s == "3969654")
                    //    speaker.speak("Background Color is Olive Green Darker 25 percentage");
                    //else if (s == "2646350")
                    //    speaker.speak("Background Color is Olive Green Darker 50 percentage");


                    //else if (s == "10642560")
                    //    speaker.speak("Background Color is Purple");
                    //else if (s == "15523813")
                    //    speaker.speak("Background Color is Purple Lighter 80 percentage");
                    //else if (s == "14270668")
                    //    speaker.speak("Background Color is Purple Lighter 60 percentage");
                    //else if (s == "13083058")
                    //    speaker.speak("Background Color is Purple Lighter 40 percentage");
                    //else if (s == "8014175")
                    //    speaker.speak("Background Color is Purple Darker 25 percentage");
                    //else if (s == "5321023")
                    //    speaker.speak("Background Color is Purple Darker 50 percentage");

                    //else if (s == "13020235")
                    //    speaker.speak("Background Color is Aqua");
                    //else if (s == "15986394")
                    //    speaker.speak("Background Color is Aqua Lighter 80 percentage");
                    //else if (s == "15261110")
                    //    speaker.speak("Background Color is Aqua Lighter 60 percentage");
                    //else if (s == "14470546")
                    //    speaker.speak("Background Color is Aqua Lighter 40 percentage");
                    //else if (s == "10191921")
                    //    speaker.speak("Background Color is Aqua Darker 25 percentage");
                    //else if (s == "6772768")
                    //    speaker.speak("Background Color is Aqua Darker 50 percentage");

                    //else if (s == "4626167")
                    //    speaker.speak("Background Color is Orange");
                    //else if (s == "14281213")
                    //    speaker.speak("Background Color is Orange Lighter 80 percentage");
                    //else if (s == "11851003")
                    //    speaker.speak("Background Color is Orange Lighter 60 percentage");
                    //else if (s == "9420794")
                    //    speaker.speak("Background Color is Orange Lighter 40 percentage");
                    //else if (s == "683235")
                    //    speaker.speak("Background Color is Orange Darker 25 percentage");
                    //else if (s == "411543")
                    //    speaker.speak("Background Color is Orange Darker 50 percentage");


                    //else if (s == "192")
                    //    speaker.speak("Background Color is Standard Dark Red");
                    //else if (s == "255")
                    //    speaker.speak("Background Color is Standard Red");
                    //else if (s == "49407")
                    //    speaker.speak("Background Color is Standard Orange");
                    //else if (s == "65535")
                    //    speaker.speak("Background Color is Standard Yellow");
                    //else if (s == "5296274")
                    //    speaker.speak("Background Color is Standard Light Green");
                    //else if (s == "5287936")
                    //    speaker.speak("Background Color is Standard Green");
                    //else if (s == "15773696")
                    //    speaker.speak("Background Color is Standard Light Blue");
                    //else if (s == "12611584")
                    //    speaker.speak("Background Color is Standard Blue");
                    //else if (s == "6299648")
                    //    speaker.speak("Background Color is Standard Dark Blue");
                    //else if (s == "10498160")
                    //    speaker.speak("Background Color is Standard Purple");
                    //else
                    //    speaker.speak("No Background Color");

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void readTitle()
        {
            try
            {
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;

                    //Word.Selection select = word.Selection;
                    //Object end = select.Start;                            
                    //select.MoveLeft(ref unitChar, ref count, ref extend);
                    //Object start = select.Start;
                    //Word.Range rng=word.ActiveDocument.Range(ref start,ref end);
                    //rng.Delete(ref unitChar, ref count);
                    if (keyData.ToString().Equals("T"))
                        speaker.speak("Title List ");

                    speaker.speak(word.ActiveDocument.Name.ToString());

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void TabPress()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                    {
                        Word.Selection select = word.Selection;
                        int x = word.Selection.Rows.First.Index;
                        int y = word.Selection.Columns.First.Index;                        //MessageBox.Show("Row = " + word.Selection.get_Information(Word.WdInformation.wdStartOfRangeRowNumber).ToString() + " Column " + word.Selection.get_Information(Word.WdInformation.wdStartOfRangeColumnNumber).ToString());
                        String ss = "Row " + x + " Collumn " + y;
                        ss = ss + " " + select.Text;
                        if (select.Text.Length == 2)
                            speaker.speak(ss + " Blank");
                        else
                            speaker.speak(ss);
                    }
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void PressHome()         //3
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;

                    Word.Selection select = word.Selection;
                    Word.Selection rangeSelect = word.Selection;

                    //select.MoveUp(ref unitLine, ref count, ref extend);
                    //Object start = select.Start;                            
                    //select.MoveDown(ref unitLine, ref count, ref extend);
                    //Object end = select.Start;
                    //Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                    select.HomeKey(ref unitN, ref extend);
                    Object start = select.Start;
                    select.EndKey(ref unitN, ref extend);
                    Object end = select.Start;

                    Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                    select.HomeKey(ref unitN, ref extend);


                    String s = rng.Text;

                    if (s[0].ToString() == "‘" || s[0].ToString() == "'" || s[0].ToString() == "’")
                        speaker.speak("Single quotation");
                    else if (s[0].ToString() == "”" || s[0].ToString() == "“" || s[0].ToString() == "\"")
                        speaker.speak("Double quotes");
                    else if (s[0].ToString() == "!")
                        speaker.speak("Exclamation mark");
                    //else if (s[0].ToString()=="")
                    //    speaker.speak("Double quotes");
                    //else if (s[0].ToString() == """)
                    //    speaker.speak("Double quotes");
                    else if (s[0].ToString() == "(")
                        speaker.speak("first bracker Open");
                    else if (s[0].ToString() == ")")
                        speaker.speak("first bracker Close");
                    else if (s[0].ToString() == ",")
                        speaker.speak("Comma");
                    else if (s[0].ToString() == "-")
                        speaker.speak("Hyphen");
                    else if (s[0].ToString() == ".")
                        speaker.speak(" full stop");
                    else if (s[0].ToString() == ":")
                        speaker.speak("Colon");
                    else if (s[0].ToString() == ";")
                        speaker.speak("Semicolon");
                    else if (s[0].ToString() == "?")
                        speaker.speak("Question mark");
                    else if (s[0].ToString() == "[")
                        speaker.speak("third bracket Open");
                    else if (s[0].ToString() == "]")
                        speaker.speak("third bracket Close");
                    else if (s[0].ToString() == "`")
                        speaker.speak("Grave accent");
                    else if (s[0].ToString() == "{")
                        speaker.speak("Second bracket Open");
                    else if (s[0].ToString() == "}")
                        speaker.speak("Second bracket Close");
                    else if (s[0].ToString() == "~")
                        speaker.speak("Equivalency sign ");
                    else if (s[0].ToString() == " ")
                        speaker.speak("Space");
                    else
                        speaker.speak(s[0].ToString());


                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void PressEnd()         //3
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;

                    speaker.speak(" blank");

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void operateSelectedWordPara()                // shift selected text to speech  //11
        {
            try
            {
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);

                }
                else
                {
                    occupiedBuffer = 1;
                    if (keyData.ToString().Equals("Right") || keyData.ToString().Equals("End"))
                    {
                        Word.Selection select = word.Selection;

                        if (select.Text == "‘" || select.Text == "'" || select.Text == "’")
                            speaker.speak("Single quotation");
                        else if (select.Text == "”" || select.Text == "“" || select.Text == "\"")
                            speaker.speak("Double quotes");
                        else if (select.Text == "!")
                            speaker.speak("Exclamation mark");
                        //else if (select.Text=="")
                        //    speaker.speak("Double quotes");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        else if (select.Text == "(")
                            speaker.speak("first bracker Open");
                        else if (select.Text == ")")
                            speaker.speak("first bracker Close");
                        else if (select.Text == ",")
                            speaker.speak("Comma");
                        else if (select.Text == "-")
                            speaker.speak("Hyphen");
                        else if (select.Text == ".")
                            speaker.speak(" full stop");
                        else if (select.Text == ":")
                            speaker.speak("Colon");
                        else if (select.Text == ";")
                            speaker.speak("Semicolon");
                        else if (select.Text == "?")
                            speaker.speak("Question mark");
                        else if (select.Text == "[")
                            speaker.speak("third bracket Open");
                        else if (select.Text == "]")
                            speaker.speak("third bracket Close");
                        else if (select.Text == "`")
                            speaker.speak("Grave accent");
                        else if (select.Text == "{")
                            speaker.speak("Second bracket Open");
                        else if (select.Text == "}")
                            speaker.speak("Second bracket Close");
                        else if (select.Text == "~")
                            speaker.speak("Equivalency sign ");
                        else if (select.Text == " ")
                            speaker.speak("Space");
                        else
                            speaker.speak(select.Text);
                        shiftselecttext = select.Text;
                    }
                    else if (keyData.ToString().Equals("Left") || keyData.ToString().Equals("Home"))
                    {
                        Word.Selection select = word.Selection;
                        //MessageBox.Show(select.Text);
                        if (select.Text == "‘" || select.Text == "'" || select.Text == "’")
                            speaker.speak("Single quotation");
                        else if (select.Text == "”" || select.Text == "“" || select.Text == "\"")
                            speaker.speak("Double quotes");
                        else if (select.Text == "!")
                            speaker.speak("Exclamation mark");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        else if (select.Text == "(")
                            speaker.speak("first bracker Open");
                        else if (select.Text == ")")
                            speaker.speak("first bracker Close");
                        else if (select.Text == ",")
                            speaker.speak("Comma");
                        else if (select.Text == "-")
                            speaker.speak("Hyphen");
                        else if (select.Text == ".")
                            speaker.speak(" full stop");
                        else if (select.Text == ":")
                            speaker.speak("Colon");
                        else if (select.Text == ";")
                            speaker.speak("Semicolon");
                        else if (select.Text == "?")
                            speaker.speak("Question mark");
                        else if (select.Text == "[")
                            speaker.speak("third bracket Open");
                        else if (select.Text == "]")
                            speaker.speak("third bracket Close");
                        else if (select.Text == "`")
                            speaker.speak("Grave accent");
                        else if (select.Text == "{")
                            speaker.speak("Second bracket Open");
                        else if (select.Text == "}")
                            speaker.speak("Second bracket Close");
                        else if (select.Text == "~")
                            speaker.speak("Equivalency sign ");
                        else if (select.Text == " ")
                            speaker.speak("Space");
                        else
                            speaker.speak(select.Text);
                        shiftselecttext = select.Text;
                    }
                    else if (keyData.ToString().Equals("Up"))
                    {
                        Word.Selection select = word.Selection;

                        if (select.Text == "‘" || select.Text == "'" || select.Text == "’")
                            speaker.speak("Single quotation");
                        else if (select.Text == "”" || select.Text == "“" || select.Text == "\"")
                            speaker.speak("Double quotes");
                        else if (select.Text == "!")
                            speaker.speak("Exclamation mark");
                        //else if (select.Text=="")
                        //    speaker.speak("Double quotes");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        else if (select.Text == "(")
                            speaker.speak("first bracker Open");
                        else if (select.Text == ")")
                            speaker.speak("first bracker Close");
                        else if (select.Text == ",")
                            speaker.speak("Comma");
                        else if (select.Text == "-")
                            speaker.speak("Hyphen");
                        else if (select.Text == ".")
                            speaker.speak(" full stop");
                        else if (select.Text == ":")
                            speaker.speak("Colon");
                        else if (select.Text == ";")
                            speaker.speak("Semicolon");
                        else if (select.Text == "?")
                            speaker.speak("Question mark");
                        else if (select.Text == "[")
                            speaker.speak("third bracket Open");
                        else if (select.Text == "]")
                            speaker.speak("third bracket Close");
                        else if (select.Text == "`")
                            speaker.speak("Grave accent");
                        else if (select.Text == "{")
                            speaker.speak("Second bracket Open");
                        else if (select.Text == "}")
                            speaker.speak("Second bracket Close");
                        else if (select.Text == "~")
                            speaker.speak("Equivalency sign ");
                        else if (select.Text == " ")
                            speaker.speak("Space");
                        else
                            speaker.speak(select.Text);
                        shiftselecttext = select.Text;
                    }
                    else if (keyData.ToString().Equals("Down"))
                    {
                        Word.Selection select = word.Selection;
                        //MessageBox.Show(select.Text);
                        if (select.Text == "‘" || select.Text == "'" || select.Text == "’")
                            speaker.speak("Single quotation");
                        else if (select.Text == "”" || select.Text == "“" || select.Text == "\"")
                            speaker.speak("Double quotes");
                        else if (select.Text == "!")
                            speaker.speak("Exclamation mark");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        else if (select.Text == "(")
                            speaker.speak("first bracker Open");
                        else if (select.Text == ")")
                            speaker.speak("first bracker Close");
                        else if (select.Text == ",")
                            speaker.speak("Comma");
                        else if (select.Text == "-")
                            speaker.speak("Hyphen");
                        else if (select.Text == ".")
                            speaker.speak(" full stop");
                        else if (select.Text == ":")
                            speaker.speak("Colon");
                        else if (select.Text == ";")
                            speaker.speak("Semicolon");
                        else if (select.Text == "?")
                            speaker.speak("Question mark");
                        else if (select.Text == "[")
                            speaker.speak("third bracket Open");
                        else if (select.Text == "]")
                            speaker.speak("third bracket Close");
                        else if (select.Text == "`")
                            speaker.speak("Grave accent");
                        else if (select.Text == "{")
                            speaker.speak("Second bracket Open");
                        else if (select.Text == "}")
                            speaker.speak("Second bracket Close");
                        else if (select.Text == "~")
                            speaker.speak("Equivalency sign ");
                        else if (select.Text == " ")
                            speaker.speak("Space");
                        else
                            speaker.speak(select.Text);
                        shiftselecttext = select.Text;
                    }
                    speaker.speak(" Selected");

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            { }
        }

        public void operateSelect()                // shift selected text to speech
        {
            try
            {
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);

                }
                else
                {
                    occupiedBuffer = 1;
                    shiftselecttext = " ";
                    if (keyData.ToString().Equals("Right"))
                    {
                        Word.Selection select = word.Selection;

                        if (select.Text == "‘" || select.Text == "'" || select.Text == "’")
                            speaker.speak("Single quotation");
                        else if (select.Text == "”" || select.Text == "“" || select.Text == "\"")
                            speaker.speak("Double quotes");
                        else if (select.Text == "!")
                            speaker.speak("Exclamation mark");
                        //else if (select.Text=="")
                        //    speaker.speak("Double quotes");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        else if (select.Text == "(")
                            speaker.speak("first bracker Open");
                        else if (select.Text == ")")
                            speaker.speak("first bracker Close");
                        else if (select.Text == ",")
                            speaker.speak("Comma");
                        else if (select.Text == "-")
                            speaker.speak("Hyphen");
                        else if (select.Text == ".")
                            speaker.speak(" full stop");
                        else if (select.Text == ":")
                            speaker.speak("Colon");
                        else if (select.Text == ";")
                            speaker.speak("Semicolon");
                        else if (select.Text == "?")
                            speaker.speak("Question mark");
                        else if (select.Text == "[")
                            speaker.speak("third bracket Open");
                        else if (select.Text == "]")
                            speaker.speak("third bracket Close");
                        else if (select.Text == "`")
                            speaker.speak("Grave accent");
                        else if (select.Text == "{")
                            speaker.speak("Second bracket Open");
                        else if (select.Text == "}")
                            speaker.speak("Second bracket Close");
                        else if (select.Text == "~")
                            speaker.speak("Equivalency sign ");
                        else if (select.Text == " ")
                            speaker.speak("Space");
                        else
                            speaker.speak(select.Text);
                        shiftselecttext = select.Text;
                    }
                    else if (keyData.ToString().Equals("Left"))
                    {
                        Word.Selection select = word.Selection;
                        //MessageBox.Show(select.Text);
                        if (select.Text == "‘" || select.Text == "'" || select.Text == "’")
                            speaker.speak("Single quotation");
                        else if (select.Text == "”" || select.Text == "“" || select.Text == "\"")
                            speaker.speak("Double quotes");
                        else if (select.Text == "!")
                            speaker.speak("Exclamation mark");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        else if (select.Text == "(")
                            speaker.speak("first bracker Open");
                        else if (select.Text == ")")
                            speaker.speak("first bracker Close");
                        else if (select.Text == ",")
                            speaker.speak("Comma");
                        else if (select.Text == "-")
                            speaker.speak("Hyphen");
                        else if (select.Text == ".")
                            speaker.speak(" full stop");
                        else if (select.Text == ":")
                            speaker.speak("Colon");
                        else if (select.Text == ";")
                            speaker.speak("Semicolon");
                        else if (select.Text == "?")
                            speaker.speak("Question mark");
                        else if (select.Text == "[")
                            speaker.speak("third bracket Open");
                        else if (select.Text == "]")
                            speaker.speak("third bracket Close");
                        else if (select.Text == "`")
                            speaker.speak("Grave accent");
                        else if (select.Text == "{")
                            speaker.speak("Second bracket Open");
                        else if (select.Text == "}")
                            speaker.speak("Second bracket Close");
                        else if (select.Text == "~")
                            speaker.speak("Equivalency sign ");
                        else if (select.Text == " ")
                            speaker.speak("Space");
                        else
                            speaker.speak(select.Text);
                        shiftselecttext = select.Text;
                    }
                    else if (keyData.ToString().Equals("Up"))
                    {
                        Word.Selection select = word.Selection;
                        //MessageBox.Show(select.Text);
                        if (select.Text == "‘" || select.Text == "'" || select.Text == "’")
                            speaker.speak("Single quotation");
                        else if (select.Text == "”" || select.Text == "“" || select.Text == "\"")
                            speaker.speak("Double quotes");
                        else if (select.Text == "!")
                            speaker.speak("Exclamation mark");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        else if (select.Text == "(")
                            speaker.speak("first bracker Open");
                        else if (select.Text == ")")
                            speaker.speak("first bracker Close");
                        else if (select.Text == ",")
                            speaker.speak("Comma");
                        else if (select.Text == "-")
                            speaker.speak("Hyphen");
                        else if (select.Text == ".")
                            speaker.speak(" full stop");
                        else if (select.Text == ":")
                            speaker.speak("Colon");
                        else if (select.Text == ";")
                            speaker.speak("Semicolon");
                        else if (select.Text == "?")
                            speaker.speak("Question mark");
                        else if (select.Text == "[")
                            speaker.speak("third bracket Open");
                        else if (select.Text == "]")
                            speaker.speak("third bracket Close");
                        else if (select.Text == "`")
                            speaker.speak("Grave accent");
                        else if (select.Text == "{")
                            speaker.speak("Second bracket Open");
                        else if (select.Text == "}")
                            speaker.speak("Second bracket Close");
                        else if (select.Text == "~")
                            speaker.speak("Equivalency sign ");
                        else if (select.Text == " ")
                            speaker.speak("Space");
                        else
                            speaker.speak(select.Text);
                        shiftselecttext = select.Text;
                    }
                    else if (keyData.ToString().Equals("Down"))
                    {
                        Word.Selection select = word.Selection;
                        //MessageBox.Show(select.Text);
                        if (select.Text == "‘" || select.Text == "'" || select.Text == "’")
                            speaker.speak("Single quotation");
                        else if (select.Text == "”" || select.Text == "“" || select.Text == "\"")
                            speaker.speak("Double quotes");
                        else if (select.Text == "!")
                            speaker.speak("Exclamation mark");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        else if (select.Text == "(")
                            speaker.speak("first bracker Open");
                        else if (select.Text == ")")
                            speaker.speak("first bracker Close");
                        else if (select.Text == ",")
                            speaker.speak("Comma");
                        else if (select.Text == "-")
                            speaker.speak("Hyphen");
                        else if (select.Text == ".")
                            speaker.speak(" full stop");
                        else if (select.Text == ":")
                            speaker.speak("Colon");
                        else if (select.Text == ";")
                            speaker.speak("Semicolon");
                        else if (select.Text == "?")
                            speaker.speak("Question mark");
                        else if (select.Text == "[")
                            speaker.speak("third bracket Open");
                        else if (select.Text == "]")
                            speaker.speak("third bracket Close");
                        else if (select.Text == "`")
                            speaker.speak("Grave accent");
                        else if (select.Text == "{")
                            speaker.speak("Second bracket Open");
                        else if (select.Text == "}")
                            speaker.speak("Second bracket Close");
                        else if (select.Text == "~")
                            speaker.speak("Equivalency sign ");
                        else if (select.Text == " ")
                            speaker.speak("Space");
                        else
                            speaker.speak(select.Text);
                        shiftselecttext = select.Text;
                    }
                    speaker.speak(" Selected");
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            { }
        }

        public void goTopOfPage()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;
                    Word.Selection select = word.Selection;
                    select.HomeKey(ref unitN, ref extend);
                    Object start = select.Start;
                    select.EndKey(ref unitN, ref extend);
                    Object end = select.Start;

                    Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                    select.HomeKey(ref unitN, ref extend);
                    if (rng.Text == null) speaker.speak("Page 1 Top of File  ");
                    else speaker.speak(rng.Text);
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void ShiftSelectText()  // say selected text shift+insert+down arrow 
        {
            try
            {
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;
                    speaker.speak(shiftselecttext);
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void goBottomOfPage()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;
                    Word.Selection select = word.Selection;
                    select.HomeKey(ref unitN, ref extend);
                    Object start = select.Start;
                    select.EndKey(ref unitN, ref extend);
                    Object end = select.Start;

                    Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                    select.HomeKey(ref unitN, ref extend);
                    if (rng.Text == null) speaker.speak("Bottom of the page Blank");
                    else speaker.speak("Bottom of the page " + rng.Text);
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }


        }
        public void ReadWhenPressN()
        {

            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;
                    //MessageBox.Show("fdfd");
                    //MessageBox.Show("ffffffffff")
                    speaker.speak("You press control plus n to create a new word document");

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void backSpace() ///// প্রতিটা charecter operate করার জন্য।
        {

            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    object pos = 1;
                    occupiedBuffer = 1;

                    //////////////////////////////////////////////////////////////////////

                    //Word.Selection select = word.Selection;

                    //select.MoveRight(ref unitWord, ref count, ref extend);
                    //Object end = select.Start;
                    //select.MoveLeft(ref unitWord, ref count, ref extend);
                    //Object start = select.Start;

                    //Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                    //MessageBox.Show(rng.Text);
                    /////////////////////////////////////////////////////////////////////

                    Word.Range s = word.Selection.Previous(ref unitChar, ref pos);
                    //s.Characters.First.Delete(ref unitChar, ref pos);
                    //Word.Selection select = word.Selection;                    
                    //select.MoveLeft(ref unitChar, ref count, ref extend);
                    String deletedText = s.Text;
                    if (deletedText == " ") speaker.speak("Space");
                    speaker.stop();
                    speaker.speak(deletedText);  //  + " Delete");
                    //SendKeys.SendWait("{BACKSPACE}");                    
                    //s.Delete(ref unitChar, ref pos);


                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void LineNumber() ///// প্রতিটা charecter operate করার জন্য।
        {

            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    object pos = 1;
                    occupiedBuffer = 1;

                    String LN = "Line Number " + word.Selection.get_Information(Word.WdInformation.wdFirstCharacterLineNumber).ToString();
                    speaker.speak(LN.ToString());
                    //s.Delete(ref unitChar, ref pos);                    
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception)
            {
            }
        }
        public void delete() ///// প্রতিটা charecter operate করার জন্য।
        {

            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    object pos = 1;
                    occupiedBuffer = 1;
                    Word.Characters selectChar = word.Selection.Characters;
                    String deletedText = selectChar.Last.Text;
                    if (deletedText == " ") deletedText = "Space";
                    //speaker.stop();
                    speaker.speak(deletedText);
                    //s.Delete(ref unitChar, ref pos);                    
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception)
            {
            }
        }
        public void SayLine()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    object pos = 1;
                    occupiedBuffer = 1;
                    String listFormat = null;
                    String ss = null;
                    if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListSimpleNumbering)
                    {
                        listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                    }
                    else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListBullet)
                    {
                        listFormat = "Bullet";
                    }
                    else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListListNumOnly)
                    {
                        listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                    }
                    else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListPictureBullet)
                    {
                        listFormat = "Picture Bullet";
                    }
                    else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListOutlineNumbering)
                    {
                        listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                    }
                    else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListMixedNumbering)
                    {
                        listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                    }

                    if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                    {
                        pre = pre + 1;
                        int x = word.Selection.Tables[1].Rows.Count;
                        int y = word.Selection.Tables[1].Columns.Count;
                        if (pre == 1) ss = "Table With " + x + " Row " + y + " Column ";
                        else ss = null;
                        int row = word.Selection.Rows.First.Index;
                        if (row != pRow) ss = ss + "Row " + row;
                        pRow = row;
                    }
                    else
                    {
                        pre = 0; pRow = 0; pCoulmn = 0;
                    }

                    //String LN = "Line Number is " + word.Selection.get_Information(Word.WdInformation.wdFirstCharacterLineNumber).ToString();
                    Word.Selection select = word.Selection;
                    Word.Selection rangeSelect = word.Selection;

                    //select.MoveUp(ref unitLine, ref count, ref extend);
                    //Object start = select.Start;                            
                    //select.MoveDown(ref unitLine, ref count, ref extend);
                    //Object end = select.Start;
                    //Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                    select.HomeKey(ref unitN, ref extend);
                    Object start = select.Start;
                    select.EndKey(ref unitN, ref extend);
                    Object end = select.Start;

                    Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                    select.HomeKey(ref unitN, ref extend);
                    if (rng.Text == null)
                        speaker.speak(ss + " " + listFormat + " " + "Blank");
                    else
                        speaker.speak(ss + " " + listFormat + " " + rng.Text);
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception)
            {
            }
        }

        public void operateChar() ///// প্রতিটা charecter operate করার জন্য।
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    if (keyData.ToString().Equals("Right"))
                    {
                        if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                        {
                            pre = pre + 1;
                            int y = word.Selection.Columns.First.Index;
                            if (y != pCoulmn) speaker.speak("Column " + y);
                            pCoulmn = y;

                        }
                        else
                        {
                            pre = 0; pCoulmn = 0; pRow = 0;
                        }

                        Word.Characters select = word.Selection.Characters;
                        char[] s=select.Last.Text.ToCharArray();
                        Console.WriteLine((int)s[0]);
                        //select.MoveLeft(ref unitChar, ref count, ref extend);

                        if (select.Last.Text == "‘" || select.Last.Text == "'" || select.Last.Text == "’")
                            speaker.speak("Single quotation");
                        else if (select.Last.Text == "”" || select.Last.Text == "“" || select.Last.Text == "\"")
                            speaker.speak("Double quotes");
                        else if (select.Last.Text == "!")
                            speaker.speak("Exclamation mark");
                        //else if (select.Last.Text == """)
                        //    speaker.speak("Double quotes");
                        //else if (select.Last.Text == """)
                        //    speaker.speak("Double quotes");
                        else if (select.Last.Text == "(")
                            speaker.speak("first bracker Open");
                        else if (select.Last.Text == ")")
                            speaker.speak("first bracker Close");
                        else if (select.Last.Text == ",")
                            speaker.speak("Comma");
                        else if (select.Last.Text == "-")
                            speaker.speak("Minus");
                        else if (select.Last.Text == ".")
                            speaker.speak(" full stop");
                        else if (select.Last.Text == ":")
                            speaker.speak("Colon");
                        else if (select.Last.Text == ";")
                            speaker.speak("Semicolon");
                        else if (select.Last.Text == "?")
                            speaker.speak("Question mark");
                        else if (select.Last.Text == "[")
                            speaker.speak("third bracket Open");
                        else if (select.Last.Text == "]")
                            speaker.speak("third bracket Close");
                        else if (select.Last.Text == "`")
                            speaker.speak("Grave accent");
                        else if (select.Last.Text == "{")
                            speaker.speak("Second bracket Open");
                        else if (select.Last.Text == "}")
                            speaker.speak("Second bracket Close");
                        else if (select.Last.Text == "~")
                            speaker.speak("Equivalency sign ");
                        else if (select.Last.Text == " ")
                            speaker.speak("Space");

                            //
                        else if (select.Last.Text == "A")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "B")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "C")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "D")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "E")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "F")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "G")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "H")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "I")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "J")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "K")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "L")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "M")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "N")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "O")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "P")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "Q")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "R")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "S")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "T")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "U")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "V")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "W")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "X")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "Y")
                            speaker.speak("Capital " + select.Last.Text);
                        else if (select.Last.Text == "Z")
                            speaker.speak("Capital " + select.Last.Text);
                        //
                        else
                            speaker.speak(select.Last.Text);
                        //select.MoveRight(ref unitChar, ref count, ref extend);
                    }
                    else if (keyData.ToString().Equals("Left"))
                    {
                        if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                        {
                            pre = pre + 1;
                            int y = word.Selection.Columns.First.Index;
                            if (y != pCoulmn) speaker.speak("Column " + y);
                            pCoulmn = y;
                        }
                        else
                        {
                            pre = 0; pCoulmn = 0; pRow = 0;
                        }

                        Word.Selection select = word.Selection;

                        if (select.Text == "‘" || select.Text == "'" || select.Text == "’")
                            speaker.speak("Single quotation");
                        else if (select.Text == "”" || select.Text == "“" || select.Text == "\"")
                            speaker.speak("Double quotes");
                        else if (select.Text == "!")
                            speaker.speak("Exclamation mark");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        //else if (select.Text == """)
                        //    speaker.speak("Double quotes");
                        else if (select.Text == "(")
                            speaker.speak("first bracker Open");
                        else if (select.Text == ")")
                            speaker.speak("first bracker Close");
                        else if (select.Text == ",")
                            speaker.speak("Comma");
                        else if (select.Text == "-")
                            speaker.speak("Minus");
                        else if (select.Text == ".")
                            speaker.speak(" full stop");
                        else if (select.Text == ":")
                            speaker.speak("Colon");
                        else if (select.Text == ";")
                            speaker.speak("Semicolon");
                        else if (select.Text == "?")
                            speaker.speak("Question mark");
                        else if (select.Text == "[")
                            speaker.speak("third bracket Open");
                        else if (select.Text == "]")
                            speaker.speak("third bracket Close");
                        else if (select.Text == "`")
                            speaker.speak("Grave accent");
                        else if (select.Text == "{")
                            speaker.speak("Second bracket Open");
                        else if (select.Text == "}")
                            speaker.speak("Second bracket Close");
                        else if (select.Text == "~")
                            speaker.speak("Equivalency sign ");
                        else if (select.Text == " ")
                            speaker.speak("Space");

                            //
                        else if (select.Text == "A")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "B")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "C")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "D")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "E")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "F")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "G")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "H")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "I")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "J")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "K")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "L")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "M")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "N")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "O")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "P")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "Q")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "R")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "S")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "T")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "U")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "V")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "W")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "X")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "Y")
                            speaker.speak("Capital " + select.Text);
                        else if (select.Text == "Z")
                            speaker.speak("Capital " + select.Text);
                        //
                        else
                            speaker.speak(select.Text);
                    }

                    else if (keyData.ToString().Equals("Down") || keyData.ToString().Equals("Tab"))   // say window prompt and text insert+tab
                    // press only down key then read start to end position range text
                    {
                        String listFormat = null;
                        String ss = null;
                        if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListSimpleNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListBullet)
                        {
                            listFormat = "Bullet";
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListListNumOnly)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListPictureBullet)
                        {
                            listFormat = "Picture Bullet";
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListOutlineNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListMixedNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }

                        if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                        {
                            pre = pre + 1;
                            int x = word.Selection.Tables[1].Rows.Count;
                            int y = word.Selection.Tables[1].Columns.Count;
                            if (pre == 1) ss = "Table With " + x + " Row " + y + " Column ";
                            else ss = null;
                            int row = word.Selection.Rows.First.Index;
                            if (row != pRow) ss = ss + "Row " + row;
                            pRow = row;
                        }
                        else
                        {
                            pre = 0; pRow = 0; pCoulmn = 0;
                        }

                        //String LN = "Line Number is " + word.Selection.get_Information(Word.WdInformation.wdFirstCharacterLineNumber).ToString();
                        Word.Selection select = word.Selection;
                        Word.Selection rangeSelect = word.Selection;

                        //select.MoveUp(ref unitLine, ref count, ref extend);
                        //Object start = select.Start;                            
                        //select.MoveDown(ref unitLine, ref count, ref extend);
                        //Object end = select.Start;
                        //Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                        select.HomeKey(ref unitN, ref extend);
                        Object start = select.Start;
                        select.EndKey(ref unitN, ref extend);
                        Object end = select.Start;

                        Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                        select.HomeKey(ref unitN, ref extend);
                        if (rng.Text == null)
                            speaker.speak(ss + " " + listFormat + " " + "Blank");
                        String s = rng.Text;
                        int count = 0;
                        for (int i = 0; i < s.Length; i++)
                        {
                            if (s[i] == ' ')
                                continue;
                            else
                                count = count + 1;
                        }
                        if (count == 0) speaker.speak(ss + " " + listFormat + " " + "Blank");
                        else speaker.speak(ss + " " + listFormat + " " + rng.Text);

                    }
                    else if (keyData.ToString().Equals("Up"))    // press only up key then read start to end possition range text
                    {
                        String listFormat = null;
                        String ss = null;
                        if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListSimpleNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListBullet)
                        {
                            listFormat = "Bullet";
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListListNumOnly)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListPictureBullet)
                        {
                            listFormat = "Picture Bullet";
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListOutlineNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListMixedNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }

                        if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                        {
                            pre = pre + 1;
                            int x = word.Selection.Tables[1].Rows.Count;
                            int y = word.Selection.Tables[1].Columns.Count;
                            if (pre == 1) ss = "Table With " + x + " Row " + y + " Column ";
                            int row = word.Selection.Rows.First.Index;
                            if (row != pRow) ss = ss + "Row " + row;
                            pRow = row;
                        }
                        else
                        {
                            pre = 0; pRow = 0; pCoulmn = 0;
                        }
                        //String LN = "Line Number is " + word.Selection.get_Information(Word.WdInformation.wdFirstCharacterLineNumber).ToString();
                        Word.Selection select = word.Selection;
                        Word.Selection rangeSelect = word.Selection;
                        //select.MoveDown(ref unitLine, ref count, ref extend);
                        //Object end = select.Start;                            
                        //select.MoveUp(ref unitLine, ref count, ref extend);
                        //Object start = select.Start;
                        //Word.Range rng = word.ActiveDocument.Range(ref start, ref end);

                        select.HomeKey(ref unitN, ref extend);
                        Object start = select.Start;
                        select.EndKey(ref unitN, ref extend);
                        Object end = select.Start;

                        Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                        select.HomeKey(ref unitN, ref extend);
                        if (rng.Text == null)
                            speaker.speak(ss + " " + listFormat + " " + "Blank");
                        String s = rng.Text;
                        int count = 0;
                        for (int i = 0; i < s.Length; i++)
                        {
                            if (s[i] == ' ')
                                continue;
                            else
                                count = count + 1;
                        }
                        if (count == 0) speaker.speak(ss + " " + listFormat + " " + "Blank");
                        else speaker.speak(ss + " " + listFormat + " " + rng.Text);
                    }

                    if (word.Selection.Type == Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionInlineShape)
                    {
                        speaker.speak("selection is on image");
                    }

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception)
            {
            }
        }

        public void operateWord()     ///// প্রতিটা word operate করার জন্য।              //7      update
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);

                object pos = 0;

                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;

                    if (keyData.ToString().Equals("Right"))                     //7      update
                    {
                        String output = " ";

                        int sz = word.Selection.Words.Last.Text.Length;
                        String s = word.Selection.Words.Last.Text;

                        for (int i = 0; i < sz; i++)
                        {


                            if (s[i].ToString() == "‘" || s[i].ToString() == "'" || s[i].ToString() == "’")
                                output = output + "Single quotation" + "\n";
                            else if (s[i].ToString() == "”" || s[i].ToString() == "“" || s[i].ToString() == "\"")
                                output = output + "Double quotes" + "\n";
                            else if (s[i].ToString() == "!")
                                output = output + "Exclamation mark" + "\n";
                            //else if (s[i].ToString()=="")
                            //    speaker.speak("Double quotes");
                            //else if (s[i].ToString() == """)
                            //    speaker.speak("Double quotes");
                            else if (s[i].ToString() == "(")
                                output = output + "first bracker Open" + "\n";
                            else if (s[i].ToString() == ")")
                                output = output + "first bracker Close" + "\n";
                            else if (s[i].ToString() == ",")
                                output = output + "Comma" + "\n";
                            else if (s[i].ToString() == "-")
                                output = output + "Hyphen" + "\n";
                            else if (s[i].ToString() == ".")
                                output = output + " full stop" + "\n";
                            else if (s[i].ToString() == ":")
                                output = output + "Colon" + "\n";
                            else if (s[i].ToString() == ";")
                                output = output + "Semicolon" + "\n";
                            else if (s[i].ToString() == "?")
                                output = output + "Question mark" + "\n";
                            else if (s[i].ToString() == "[")
                                output = output + "third bracket Open" + "\n";
                            else if (s[i].ToString() == "]")
                                output = output + "third bracket Close" + "\n";
                            else if (s[i].ToString() == "`")
                                output = output + "Grave accent" + "\n";
                            else if (s[i].ToString() == "{")
                                output = output + "Second bracket Open" + "\n";
                            else if (s[i].ToString() == "}")
                                output = output + "Second bracket Close" + "\n";
                            else if (s[i].ToString() == "~")
                                output = output + "Equivalency sign " + "\n";

                            else
                                output = output + s[i];
                        }

                        speaker.speak(output);

                        //    speaker.speak(word.Selection.Words.Last.Text);
                        //MessageBox.Show(sz.ToString());


                        //MessageBox.Show(keyData.ToString());
                        //Word.Selection select = word.Selection;
                        //Word.Selection rangeSelect = word.Selection;

                        //select.MoveLeft(ref unitWord, ref count, ref extend);

                        //Object start = select.Start;
                        //select.MoveRight(ref unitWord, ref count, ref extend);
                        //Object end = select.Start;
                        //Word.Range rng = word.ActiveDocument.Range(ref start, ref end);

                        //speaker.speak(rng.Text);

                    }
                    if (keyData.ToString().Equals("Left"))                             //7      update
                    {
                        //speaker.speak(word.Selection.Words.Last.Text);
                        Word.Selection select = word.Selection;

                        select.MoveRight(ref unitWord, ref count, ref extend);
                        Object end = select.Start;
                        select.MoveLeft(ref unitWord, ref count, ref extend);
                        Object start = select.Start;

                        Word.Range rng = word.ActiveDocument.Range(ref start, ref end);
                        String output = " ";

                        int sz = word.Selection.Words.Last.Text.Length;
                        String s = word.Selection.Words.Last.Text;

                        for (int i = 0; i < sz; i++)
                        {


                            if (s[i].ToString() == "‘" || s[i].ToString() == "'" || s[i].ToString() == "’")
                                output = output + "Single quotation" + "\n";
                            else if (s[i].ToString() == "”" || s[i].ToString() == "“" || s[i].ToString() == "\"")
                                output = output + "Double quotes" + "\n";
                            else if (s[i].ToString() == "!")
                                output = output + "Exclamation mark" + "\n";
                            //else if (s[i].ToString()=="")
                            //    speaker.speak("Double quotes");
                            //else if (s[i].ToString() == """)
                            //    speaker.speak("Double quotes");
                            else if (s[i].ToString() == "(")
                                output = output + "first bracker Open" + "\n";
                            else if (s[i].ToString() == ")")
                                output = output + "first bracker Close" + "\n";
                            else if (s[i].ToString() == ",")
                                output = output + "Comma" + "\n";
                            else if (s[i].ToString() == "-")
                                output = output + "Hyphen" + "\n";
                            else if (s[i].ToString() == ".")
                                output = output + " full stop" + "\n";
                            else if (s[i].ToString() == ":")
                                output = output + "Colon" + "\n";
                            else if (s[i].ToString() == ";")
                                output = output + "Semicolon" + "\n";
                            else if (s[i].ToString() == "?")
                                output = output + "Question mark" + "\n";
                            else if (s[i].ToString() == "[")
                                output = output + "third bracket Open" + "\n";
                            else if (s[i].ToString() == "]")
                                output = output + "third bracket Close" + "\n";
                            else if (s[i].ToString() == "`")
                                output = output + "Grave accent" + "\n";
                            else if (s[i].ToString() == "{")
                                output = output + "Second bracket Open" + "\n";
                            else if (s[i].ToString() == "}")
                                output = output + "Second bracket Close" + "\n";
                            else if (s[i].ToString() == "~")
                                output = output + "Equivalency sign " + "\n";

                            else
                                output = output + s[i];
                        }

                        speaker.speak(output);

                        //speaker.speak(word.Selection.Words.Last.Text);

                        if (rng.Text.Length > 0)
                        {
                            Word.ProofreadingErrors we = rng.SpellingErrors;
                            int iErrorCount = 0;
                            iErrorCount = we.Count;
                            //Console.WriteLine(iErrorCount);

                            if (iErrorCount != 0)
                            {
                                speaker.speak("this word is wrong");
                                //iErrorCount = 0;   
                            }

                        }

                    }
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception)
            {
            }
        }

        public void Insrt_operateWord()     ///// প্রতিটা word operate করার জন্য।
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);

                object pos = 0;

                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;

                    if (keyData.ToString().Equals("Right"))
                    {

                        //Word.Selection select = word.Selection;
                        //select.Move(ref unitWord, ref pos);

                        object pos1 = 1;
                        Word.Selection select = word.Selection;
                        select.StartOf(ref unitWord, ref extend);
                        select.Move(ref unitWord, ref pos1);

                        speaker.speak(word.Selection.Words.Last.Text);


                        //MessageBox.Show(keyData.ToString());
                        //Word.Selection select = word.Selection;
                        //Word.Selection rangeSelect = word.Selection;

                        //select.MoveLeft(ref unitWord, ref count, ref extend);

                        //Object start = select.Start;
                        //select.MoveRight(ref unitWord, ref count, ref extend);
                        //Object end = select.Start;
                        //Word.Range rng = word.ActiveDocument.Range(ref start, ref end);

                        //speaker.speak(rng.Text);

                    }
                    if (keyData.ToString().Equals("Left"))
                    {
                        object pos1 = -1;
                        Word.Selection select = word.Selection;
                        select.StartOf(ref unitWord, ref extend);
                        select.Move(ref unitWord, ref pos1);

                        speaker.speak(word.Selection.Words.Last.Text);
                        //Word.Selection select = word.Selection;

                        //select.MoveRight(ref unitWord, ref count, ref extend);
                        //Object end = select.Start;
                        //select.MoveLeft(ref unitWord, ref count, ref extend);
                        //Object start = select.Start;

                        //Word.Range rng = word.ActiveDocument.Range(ref start, ref end);


                        //speaker.speak(word.Selection.Words.Last.Text);
                        //if (rng.Text.Length > 0)
                        //{
                        //    Word.ProofreadingErrors we = rng.SpellingErrors;
                        //    int iErrorCount = 0;
                        //    iErrorCount = we.Count;
                        //    //Console.WriteLine(iErrorCount);

                        //    if (iErrorCount != 0)
                        //    {
                        //        speaker.speak("this word is wrong");
                        //        //iErrorCount = 0;   
                        //    }

                        //}

                    }
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception)
            {
            }
        }

        public void TableCurCellInfo()     ///// প্রতিটা word operate করার জন্য।
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);

                object pos = 0;

                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;

                    if ((bool)word.Selection.get_Information(Word.WdInformation.wdWithInTable))
                    {
                        speaker.speak("Row " + word.Selection.get_Information(Word.WdInformation.wdStartOfRangeRowNumber).ToString() + " Column " + word.Selection.get_Information(Word.WdInformation.wdStartOfRangeColumnNumber).ToString());

                    }

                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception)
            {
            }
        }


        public void operateSentence()      ///// প্রতিটা Sentence operate করার জন্য।
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                object pos = 0;

                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;

                    if (keyData.ToString().Equals("p"))
                    {
                        Word.Selection select = word.Selection;
                        Word.Selection rangeSelect = word.Selection;

                        select.MoveLeft(ref unitSentence, ref pos, ref extend);
                        Object start = select.Start;

                        select.MoveRight(ref unitSentence, ref count, ref extend);
                        Object end = select.Start;
                        Word.Range rng = word.ActiveDocument.Range(ref start, ref end);

                        // MessageBox.Show(rng.Text);
                        speaker.speak(rng.Text);



                    }
                    if (keyData.ToString().Equals("o"))
                    {
                        Word.Selection select = word.Selection;

                        select.MoveRight(ref unitSentence, ref pos, ref extend);
                        Object end = select.Start;

                        select.MoveLeft(ref unitSentence, ref count, ref extend);
                        Object start = select.Start;

                        Word.Range rng = word.ActiveDocument.Range(ref start, ref end);

                        MessageBox.Show(rng.Text);
                        speaker.speak(rng.Text);
                    }
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception)
            {
            }
        }

        public void operateParagraph()      ///// প্রতিটা Paragraph operate করার জন্য।
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                object pos = 0;
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;

                    if (keyData.ToString().Equals("Up"))
                    {
                        String listFormat = null;
                        String select = null;
                        String previousSelect = null;
                        while (previousSelect == null)
                        {
                            try
                            {
                                previousSelect = word.Selection.Previous(ref unitParagraph, ref count).Text;
                            }
                            catch (Exception)
                            {
                                while (word.Selection.Paragraphs.Last.Range.Text.Trim() == "")
                                {
                                    word.Selection.MoveDown(ref unitParagraph, ref count, ref extend);
                                }
                                break;
                            }
                            word.Selection.MoveUp(ref unitParagraph, ref count, ref extend);
                            if (previousSelect.Trim() == "") previousSelect = null;
                        }
                        select = word.Selection.Paragraphs.Last.Range.Text;
                        if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListSimpleNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListBullet)
                        {
                            listFormat = "Bullet";
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListListNumOnly)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListPictureBullet)
                        {
                            listFormat = "Picture Bullet";
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListOutlineNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListMixedNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        //Word.Selection select = word.Selection;
                        //Word.Selection rangeSelect = word.Selection;

                        //select.MoveUp(ref unitParagraph, ref pos, ref extend);
                        //Object start = select.Start;

                        //select.MoveDown(ref unitParagraph, ref count, ref extend);
                        //Object end = select.Start;

                        //select.MoveUp(ref unitParagraph, ref count, ref extend);

                        //Word.Range rng = word.ActiveDocument.Range(ref start, ref end);

                        //MessageBox.Show(rng.Text);
                        if (select.Trim() == "") select = "Blank";
                        speaker.speak(listFormat + " " + select);

                    }
                    if (keyData.ToString().Equals("Down"))
                    {
                        String listFormat = null;
                        String select = null;
                        String nextSelect = null;
                        while (nextSelect == null)
                        {
                            try
                            {
                                nextSelect = word.Selection.Next(ref unitParagraph, ref count).Text;
                            }
                            catch (Exception)
                            {
                                while (word.Selection.Paragraphs.Last.Range.Text.Trim() == "")
                                {
                                    word.Selection.MoveUp(ref unitParagraph, ref count, ref extend);
                                }
                                break;
                            }
                            word.Selection.MoveDown(ref unitParagraph, ref count, ref extend);
                            if (nextSelect.Trim() == "") nextSelect = null;
                        }
                        select = word.Selection.Paragraphs.Last.Range.Text;
                        if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListSimpleNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListBullet)
                        {
                            listFormat = "Bullet";
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListListNumOnly)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListPictureBullet)
                        {
                            listFormat = "Picture Bullet";
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListOutlineNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        else if (word.Selection.Range.ListFormat.ListType == Microsoft.Office.Interop.Word.WdListType.wdListMixedNumbering)
                        {
                            listFormat = word.Selection.Range.ListFormat.ListString.ToString();
                        }
                        //speaker.speak(word.Selection.Range.ListFormat.ListString.ToString());
                        //MessageBox.Show(word.Selection.Range.ListFormat.ListTemplate.ListLev);
                        //if (select == null)
                        //    speaker.speak("Blank");
                        //else
                        //    speaker.speak(select);
                        //word.Selection.b

                        //Word.Selection select = word.Selection.Range.ListFormat.ListString.ToString();
                        //select.MoveUp(ref unitParagraph, ref count, ref extend);
                        //Object start = select.Start;

                        //select.MoveDown(ref unitParagraph, ref count, ref extend);
                        //Object end = select.End;

                        //Word.Range rng = word.ActiveDocument.Range(ref start, ref end);                        
                        //MessageBox.Show(rng.Text);
                        if (select.Trim() == "") select = "Blank";
                        speaker.speak(listFormat + " " + select);
                    }
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void RinbbonUnselect()
        {
            try
            {
                shiftselecttext = "";
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {
                    occupiedBuffer = 1;
                    word.CommandBars.ReleaseFocus();
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void stopAll()
        {
            speaker.stop();

        }
    }
}

