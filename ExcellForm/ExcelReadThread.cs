using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using SpeechBuilder;
using System.Windows.Forms;
using System.Threading;
using System.Reflection;

namespace ExcellForm
{
    class ExcelReadThread
    {
        Microsoft.Office.Interop.Excel.Application excel = null;
        Microsoft.Office.Interop.Excel._Workbook book = null;
        private Excel._Worksheet sheet = null;
        //private static Microsoft.Office.Interop.Word.Application wd = null;
        char a;
        static int k = 0;
        static int ct;
        private String keyData = null;
        private int occupiedBuffer = 0;

        private static int track_alt_D1 = 0, track_alt_ctrl_D1 = 0, track_F2 = 0, up_track, down_press = 0, track_first_size = 0;
        //Object unitWord = Excel.;
        //Object unitSentence = Word.WdUnits.wdSentence;
        //Object unitParagraph = Word.WdUnits.wdParagraph;       
        //Object unitFullPage = Word.WdUnits.wdFullPage;
        private static String s = null;
        object newTemplate = false;
        object docType = 0;
        object isVisible = true;
        static object p = null;
        //private Form2 f2;

        static int present_sz;

        Object count = 1;
        //Object extend = Word.WdMovementType.wdMove;
        private SpeechControl speaker;

        public ExcelReadThread(SpeechControl speaker, Microsoft.Office.Interop.Excel.Application excel,
            Microsoft.Office.Interop.Excel._Workbook book, String keyData)
        {
            this.excel = excel;
            //p = excel;
            this.book = book;
            this.keyData = keyData;
            //speaker = new SpeechControl();
            this.speaker = speaker;

        }
        public ExcelReadThread()
        {
        }
        public void Test()
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

                    //Excel.Range rng = (Excel.Range)excel.ActiveCell;
                    //sheet = (Excel._Worksheet)book.ActiveSheet;
                    String s = excel.ActiveWorkbook.Name;
                    speaker.speak(s);

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

        public void operateFull() ///// প্রতিটা charecter operate করার জন্য।
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

                    if ((keyData.ToString().Equals("Right")) || (keyData.ToString().Equals("Left")) || (keyData.ToString().Equals("Up")) || (keyData.ToString().Equals("Down")) || (keyData.ToString().Equals("Tab")))
                    {
                        speaker.stop();
                        Excel.Range rng = excel.ActiveCell;
                        //sheet = (Excel.Worksheet)book.ActiveSheet;
                        //String ss = sheet.Name.ToString(); //return sheet name
                        //String ss = sheet.Index.ToString(); // return sheet index
                        int y = excel.ActiveCell.Cells.Row;
                        int x = excel.ActiveCell.Cells.Column;
                        //MessageBox.Show("Row " + x + " Column " + y);
                        //speaker.speak(rng.Text.ToString());

                        string row = "";
                        int div = 0, per = 0;
                        int q;
                        q = x / (26 * 26);
                        if (q == 1) row = "A";
                        else if (q == 2) row = "B";
                        else if (q == 3) row = "C";
                        else if (q == 4) row = "D";
                        else if (q == 5) row = "E";
                        else if (q == 6) row = "F";
                        else if (q == 7) row = "G";
                        else if (q == 8) row = "H";
                        else if (q == 9) row = "I";
                        else if (q == 10) row = "J";
                        else if (q == 11) row = "K";
                        else if (q == 12) row = "L";
                        else if (q == 13) row = "M";
                        else if (q == 14) row = "N";
                        else if (q == 15) row = "O";
                        else if (q == 16) row = "P";
                        else if (q == 17) row = "Q";
                        else if (q == 18) row = "R";
                        else if (q == 19) row = "S";
                        else if (q == 20) row = "T";
                        else if (q == 21) row = "U";
                        else if (q == 22) row = "V";
                        else if (q == 23) row = "W";
                        else if (q == 24) row = "X";
                        else if (q == 25) row = "Y";
                        else if (q == 26) row = "Z";

                        x = x - (26 * 26 * q);


                        div = x / 26;
                        per = x % 26;
                        Console.WriteLine(div);
                        Console.WriteLine(per);
                        if (div == 1) row = row + "A";
                        else if (div == 2) row = row + "B";
                        else if (div == 3) row = row + "C";
                        else if (div == 4) row = row + "D";
                        else if (div == 5) row = row + "E";
                        else if (div == 6) row = row + "F";
                        else if (div == 7) row = row + "G";
                        else if (div == 8) row = row + "H";
                        else if (div == 9) row = row + "I";
                        else if (div == 10) row = row + "J";
                        else if (div == 11) row = row + "K";
                        else if (div == 12) row = row + "L";
                        else if (div == 13) row = row + "M";
                        else if (div == 14) row = row + "N";
                        else if (div == 15) row = row + "O";
                        else if (div == 16) row = row + "P";
                        else if (div == 17) row = row + "Q";
                        else if (div == 18) row = row + "R";
                        else if (div == 19) row = row + "S";
                        else if (div == 20) row = row + "T";
                        else if (div == 21) row = row + "U";
                        else if (div == 22) row = row + "V";
                        else if (div == 23) row = row + "W";
                        else if (div == 24) row = row + "X";
                        else if (div == 25) row = row + "Y";
                        else if (div == 26) row = row + "Z";


                        if (per == 1) row = row + "A";
                        else if (per == 2) row = row + "B";
                        else if (per == 3) row = row + "C";
                        else if (per == 4) row = row + "D";
                        else if (per == 5) row = row + "E";
                        else if (per == 6) row = row + "F";
                        else if (per == 7) row = row + "G";
                        else if (per == 8) row = row + "H";
                        else if (per == 9) row = row + "I";
                        else if (per == 10) row = row + "J";
                        else if (per == 11) row = row + "K";
                        else if (per == 12) row = row + "L";
                        else if (per == 13) row = row + "M";
                        else if (per == 14) row = row + "N";
                        else if (per == 15) row = row + "O";
                        else if (per == 16) row = row + "P";
                        else if (per == 17) row = row + "Q";
                        else if (per == 18) row = row + "R";
                        else if (per == 19) row = row + "S";
                        else if (per == 20) row = row + "T";
                        else if (per == 21) row = row + "U";
                        else if (per == 22) row = row + "V";
                        else if (per == 23) row = row + "W";
                        else if (per == 24) row = row + "X";
                        else if (per == 25) row = row + "Y";
                        else if (per == 26) row = row + "Z";

                        speaker.speak(row + " " + y);
                        if (rng.Text.ToString() == "") speaker.speak("Blank");
                        speaker.speak(rng.Text.ToString());
                        //MessageBox.Show(row + " " + y);

                        track_alt_D1 = 0;
                        track_alt_ctrl_D1 = 0;
                    }

                    else if (keyData.ToString().Equals("M"))
                    {
                        speaker.stop();
                        Excel.Range rng = excel.ActiveCell;
                        //sheet = (Excel.Worksheet)book.ActiveSheet;
                        //String ss = sheet.Name.ToString(); //return sheet name
                        //String ss = sheet.Index.ToString(); // return sheet index
                        int y = excel.ActiveCell.Cells.Row;
                        int x = excel.ActiveCell.Cells.Column;
                        //speaker.speak("Row " + x + " Column " + y);
                        //speaker.speak(rng.Text.ToString());

                        string row = "";
                        int div = 0, per = 0;
                        int q;
                        q = x / (26 * 26);
                        if (q == 1) row = "A";
                        else if (q == 2) row = "B";
                        else if (q == 3) row = "C";
                        else if (q == 4) row = "D";
                        else if (q == 5) row = "E";
                        else if (q == 6) row = "F";
                        else if (q == 7) row = "G";
                        else if (q == 8) row = "H";
                        else if (q == 9) row = "I";
                        else if (q == 10) row = "J";
                        else if (q == 11) row = "K";
                        else if (q == 12) row = "L";
                        else if (q == 13) row = "M";
                        else if (q == 14) row = "N";
                        else if (q == 15) row = "O";
                        else if (q == 16) row = "P";
                        else if (q == 17) row = "Q";
                        else if (q == 18) row = "R";
                        else if (q == 19) row = "S";
                        else if (q == 20) row = "T";
                        else if (q == 21) row = "U";
                        else if (q == 22) row = "V";
                        else if (q == 23) row = "W";
                        else if (q == 24) row = "X";
                        else if (q == 25) row = "Y";
                        else if (q == 26) row = "Z";

                        x = x - (26 * 26 * q);


                        div = x / 26;
                        per = x % 26;
                        Console.WriteLine(div);
                        Console.WriteLine(per);
                        if (div == 1) row = row + "A";
                        else if (div == 2) row = row + "B";
                        else if (div == 3) row = row + "C";
                        else if (div == 4) row = row + "D";
                        else if (div == 5) row = row + "E";
                        else if (div == 6) row = row + "F";
                        else if (div == 7) row = row + "G";
                        else if (div == 8) row = row + "H";
                        else if (div == 9) row = row + "I";
                        else if (div == 10) row = row + "J";
                        else if (div == 11) row = row + "K";
                        else if (div == 12) row = row + "L";
                        else if (div == 13) row = row + "M";
                        else if (div == 14) row = row + "N";
                        else if (div == 15) row = row + "O";
                        else if (div == 16) row = row + "P";
                        else if (div == 17) row = row + "Q";
                        else if (div == 18) row = row + "R";
                        else if (div == 19) row = row + "S";
                        else if (div == 20) row = row + "T";
                        else if (div == 21) row = row + "U";
                        else if (div == 22) row = row + "V";
                        else if (div == 23) row = row + "W";
                        else if (div == 24) row = row + "X";
                        else if (div == 25) row = row + "Y";
                        else if (div == 26) row = row + "Z";


                        if (per == 1) row = row + "A";
                        else if (per == 2) row = row + "B";
                        else if (per == 3) row = row + "C";
                        else if (per == 4) row = row + "D";
                        else if (per == 5) row = row + "E";
                        else if (per == 6) row = row + "F";
                        else if (per == 7) row = row + "G";
                        else if (per == 8) row = row + "H";
                        else if (per == 9) row = row + "I";
                        else if (per == 10) row = row + "J";
                        else if (per == 11) row = row + "K";
                        else if (per == 12) row = row + "L";
                        else if (per == 13) row = row + "M";
                        else if (per == 14) row = row + "N";
                        else if (per == 15) row = row + "O";
                        else if (per == 16) row = row + "P";
                        else if (per == 17) row = row + "Q";
                        else if (per == 18) row = row + "R";
                        else if (per == 19) row = row + "S";
                        else if (per == 20) row = row + "T";
                        else if (per == 21) row = row + "U";
                        else if (per == 22) row = row + "V";
                        else if (per == 23) row = row + "W";
                        else if (per == 24) row = row + "X";
                        else if (per == 25) row = row + "Y";
                        else if (per == 26) row = row + "Z";

                        speaker.speak(row + " " + y);
                        if (rng.Text.ToString() == "") speaker.speak("Blank");
                        speaker.speak(rng.Text.ToString());
                        //MessageBox.Show(row + " " + y);
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


        public void setF2() ///// প্রতিটা charecter operate করার জন্য।
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

                    Excel.Range rng = (Excel.Range)excel.ActiveCell;
                    sheet = (Excel._Worksheet)book.ActiveSheet;



                    //string s= rng.get_End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight.ToString());

                    // Range lastCell = firstCell.get_End(XlDirection.xlToRight);


                    //track_F2 = track_F2 + 1;

                    //  track_alt_D1 = 0;
                    //  track_alt_ctrl_D1 = 0;



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


        public void set_track_F2()
        {
            track_F2 = 0;
            s = null;
            present_sz = 0;
            up_track = 0;
            down_press = 0;
            track_first_size = 0;

        }

        public void tr_F2()
        {
            track_F2 = 0;
            up_track = track_F2;
            down_press = 0;
        }


        public void setF2_initial()
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

                    Excel.Range rng = (Excel.Range)excel.ActiveCell;
                    //sheet = (Excel._Worksheet)book.ActiveSheet;

                    try
                    {
                        //String ss = sheet.Name.ToString(); //return sheet name
                        //String ss = sheet.Index.ToString(); // return sheet index
                        Excel.Range r = (Excel.Range)excel.Selection;
                        int x = excel.ActiveCell.Cells.Row;
                        int y = excel.ActiveCell.Cells.Column;
                        //   speaker.speak("Row " + x + " Column " + y);
                        //speaker.speak(rng.Text.ToString());
                        //excel.ActiveCell.Formula.ToString();
                        //s = rng.Text.ToString();
                        s = excel.ActiveCell.Formula.ToString();
                        present_sz = s.Length;
                        track_F2 = present_sz;
                        //MessageBox.Show(track_F2.ToString());
                        //MessageBox.Show(rng.get_Characters(1, 1).Text.ToString());
                    }

                    catch (Exception)
                    {

                    }
                    track_alt_D1 = 0;
                    track_alt_ctrl_D1 = 0;

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

        public void operateChar() ///// প্রতিটা charecter operate করার জন্য।
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

                    if ((keyData.ToString().Equals("Right")))
                    {
                        try
                        {

                            present_sz = s.Length;

                            if (present_sz > track_F2)
                            {
                                track_F2 = track_F2 + 1;

                                if (s[track_F2].ToString() == "!")
                                    speaker.speak("Exclamation mark");
                                else if (s[track_F2].ToString() == "(")
                                    speaker.speak("first bracket Open");
                                else if (s[track_F2].ToString() == ")")
                                    speaker.speak("first bracket Close");
                                else if (s[track_F2].ToString() == ",")
                                    speaker.speak("Comma");
                                else if (s[track_F2].ToString() == "-")
                                    speaker.speak("Hyphen");
                                else if (s[track_F2].ToString() == ".")
                                    speaker.speak(" full stop");
                                else if (s[track_F2].ToString() == ":")
                                    speaker.speak("Colon");
                                else if (s[track_F2].ToString() == ";")
                                    speaker.speak("Semicolon");
                                else if (s[track_F2].ToString() == "?")
                                    speaker.speak("Question mark");
                                else if (s[track_F2].ToString() == "[")
                                    speaker.speak("third bracket Open");
                                else if (s[track_F2].ToString() == "]")
                                    speaker.speak("third bracket Close");
                                else if (s[track_F2].ToString() == "`")
                                    speaker.speak("Grave accent");
                                else if (s[track_F2].ToString() == "{")
                                    speaker.speak("Second bracket Open");
                                else if (s[track_F2].ToString() == "}")
                                    speaker.speak("Second bracket Close");
                                else if (s[track_F2].ToString() == "~")
                                    speaker.speak("Equivalency sign ");
                                else if (s[track_F2].ToString() == " ")
                                    speaker.speak("Space");
                                //
                                else if (s[track_F2].ToString() == "‘" || s[track_F2].ToString() == "'" || s[track_F2].ToString() == "’")
                                    speaker.speak("Single quotation");
                                else if (s[track_F2].ToString() == "”" || s[track_F2].ToString() == "“" || s[track_F2].ToString() == "\"")
                                    speaker.speak("Double quotes");

                                       //
                                else if (s[track_F2].ToString() == "A")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "B")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "C")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "D")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "E")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "F")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "G")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "H")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "I")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "J")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "K")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "L")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "M")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "N")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "O")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "P")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "Q")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "R")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "S")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "T")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "U")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "V")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "W")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "X")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "Y")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "Z")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                //

                                else
                                    speaker.speak(s[track_F2].ToString());

                                
                            }
                            //MessageBox.Show(rng.get_Characters(1, 1).Text.ToString());
                        }

                        catch (Exception e)
                        {
                            //MessageBox.Show(e.ToString());
                        }

                        down_press = 0;
                        track_alt_D1 = 0;
                        track_alt_ctrl_D1 = 0;
                    }

                    else if ((keyData.ToString().Equals("Left")))
                    {
                        try
                        {


                            present_sz = s.Length;

                            if (track_F2 > 0)
                            {
                                track_F2 = track_F2 - 1;
                                //
                                if (s[track_F2].ToString() == "!")
                                    speaker.speak("Exclamation mark");
                                else if (s[track_F2].ToString() == "(")
                                    speaker.speak("first bracket Open");
                                else if (s[track_F2].ToString() == ")")
                                    speaker.speak("first bracket Close");
                                else if (s[track_F2].ToString() == ",")
                                    speaker.speak("Comma");
                                else if (s[track_F2].ToString() == "-")
                                    speaker.speak("Hyphen");
                                else if (s[track_F2].ToString() == ".")
                                    speaker.speak(" full stop");
                                else if (s[track_F2].ToString() == ":")
                                    speaker.speak("Colon");
                                else if (s[track_F2].ToString() == ";")
                                    speaker.speak("Semicolon");
                                else if (s[track_F2].ToString() == "?")
                                    speaker.speak("Question mark");
                                else if (s[track_F2].ToString() == "[")
                                    speaker.speak("third bracket Open");
                                else if (s[track_F2].ToString() == "]")
                                    speaker.speak("third bracket Close");
                                else if (s[track_F2].ToString() == "`")
                                    speaker.speak("Grave accent");
                                else if (s[track_F2].ToString() == "{")
                                    speaker.speak("Second bracket Open");
                                else if (s[track_F2].ToString() == "}")
                                    speaker.speak("Second bracket Close");
                                else if (s[track_F2].ToString() == "~")
                                    speaker.speak("Equivalency sign ");
                                else if (s[track_F2].ToString() == " ")
                                    speaker.speak("Space");

                                else if (s[track_F2].ToString() == "‘" || s[track_F2].ToString() == "'" || s[track_F2].ToString() == "’")
                                    speaker.speak("Single quotation");
                                else if (s[track_F2].ToString() == "”" || s[track_F2].ToString() == "“" || s[track_F2].ToString() == "\"")
                                    speaker.speak("Double quotes");

                                     //
                                else if (s[track_F2].ToString() == "A")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "B")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "C")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "D")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "E")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "F")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "G")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "H")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "I")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "J")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "K")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "L")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "M")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "N")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "O")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "P")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "Q")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "R")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "S")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "T")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "U")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "V")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "W")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "X")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "Y")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                                else if (s[track_F2].ToString() == "Z")
                                    speaker.speak("Capital " + s[track_F2].ToString());
                               

                                //
                                else
                                    speaker.speak(s[track_F2].ToString());
                                //MessageBox.Show(s[track_F2].ToString());                                
                            }
                            //MessageBox.Show(rng.get_Characters(1, 1).Text.ToString());
                        }

                        catch (Exception)
                        {

                        }

                        down_press = 0;
                        track_alt_D1 = 0;
                        track_alt_ctrl_D1 = 0;
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

        public void update()
        {
            //MessageBox.Show("u");
            try
            {
                String sort = "";
                //MessageBox.Show("kkkkkkkkkkkk");
                if ((keyData.ToString().Equals("Back")) && track_F2 > 0)
                {
                    char[] aa = s.ToCharArray();
                    //MessageBox.Show(s[2].ToString());
                    //String c = null;

                    int i = 0;

                    for (int j = 0; j < track_F2 - 1; j++)
                        sort = sort + s[j];

                    for (i = track_F2; i < s.Length; i++)
                    {
                        sort = sort + s[i];
                    }

                    s = sort;
                    //MessageBox.Show(s);
                    track_F2 = track_F2 - 1;
                    up_track = track_F2;
                }

                else if ((keyData.ToString().Equals("End")))
                {
                    track_F2 = s.Length;
                    up_track = track_F2;
                    down_press = 0;
                }

                else if ((keyData.ToString().Equals("Down")))
                {
                    down_press = down_press + 1;
                    if (down_press == 1)
                    {
                        up_track = track_F2;
                        track_F2 = s.Length;
                    }

                }

                else if ((keyData.ToString().Equals("Up")))
                {
                    if (down_press != 0)
                        track_F2 = up_track;
                }
                Char code;
                //MessageBox.Show(keyData);
                if (keyData.ToString() == "Space")
                    code = ' ';

                else if (keyData.ToString() == "D1")
                    code = '1';

                else if (keyData.ToString() == "D2")
                    code = '2';

                else if (keyData.ToString() == "D3")
                    code = '3';

                else if (keyData.ToString() == "D4")
                    code = '4';

                else if (keyData.ToString() == "D5")
                    code = '5';

                else if (keyData.ToString() == "D6")
                    code = '6';

                else if (keyData.ToString() == "D7")
                    code = '7';

                else if (keyData.ToString() == "D8")
                    code = '8';

                else if (keyData.ToString() == "D9")
                    code = '9';

                else if (keyData.ToString() == "D0")
                    code = '0';

                else if (keyData.ToString() == "Oemcomma")
                    code = ',';

                else if (keyData.ToString() == "OemPeriod")
                    code = '.';
                else if (keyData.ToString() == "OemQuestion")
                    code = '/';
                else if (keyData.ToString() == "Oem1")
                    code = ';';
                else if (keyData.ToString() == "Oem7")
                    code = '\'';

                else if (keyData.ToString() == "OemMinus")
                    code = '-';
                else if (keyData.ToString() == "Oemplus")
                    code = '=';

                else if (keyData.ToString() == "OemOpenBrackets")
                    code = '[';
                else if (keyData.ToString() == "Oem6")
                    code = ']';
                else if (keyData.ToString() == "Oem5")
                    code = '\\';
                else if (keyData.ToString() == "Oemtilde")
                    code = '`';



                else
                    code = Convert.ToChar(keyData);
                ///////////MessageBox.Show(code.ToString());

                if ((code >= 'a' && code <= 'z') || (code >= 'A' && code <= 'Z') || (code == '\'') || (code == ';') || (code == ' ') || (code == ',') || (code == '.') || (code == '/') || (code >= 48 && code <= 57) || code == '-' || code == '=' || code == '[' || code == ']' || code == '`' || code == '\\')
                {
                    //MessageBox.Show("u");
                    char[] aa = s.ToCharArray();

                    int i = 0;

                    for (int j = 0; j < track_F2; j++)
                        sort = sort + s[j];

                    for (i = track_F2; i < s.Length; i++)
                    {
                        sort = sort + code;
                        a = s[i];
                        aa[i] = code;
                        //MessageBox.Show(aa[i].ToString());
                        code = a;

                    }

                    sort = sort + code;
                    s = sort;
                    //MessageBox.Show(s);
                    track_F2 = track_F2 + 1;
                    up_track = track_F2;
                }

                track_alt_D1 = 0;
                track_alt_ctrl_D1 = 0;
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }
        }

        public void update1()   // shift related work
        {
            //MessageBox.Show("u1");
            try
            {

                String sort = "";

                if ((keyData.ToString().Equals("Back")) && track_F2 > 0)
                {
                    char[] aa = s.ToCharArray();
                    //MessageBox.Show(s[2].ToString());
                    //String c = null;

                    int i = 0;

                    for (int j = 0; j < track_F2 - 1; j++)
                        sort = sort + s[j];

                    for (i = track_F2; i < s.Length; i++)
                    {
                        sort = sort + s[i];
                    }

                    s = sort;
                    // MessageBox.Show(s);
                    track_F2 = track_F2 - 1;
                    up_track = track_F2;
                }

                else if ((keyData.ToString().Equals("End")))
                {
                    track_F2 = s.Length;
                    up_track = track_F2;
                    down_press = 0;
                }

                else if ((keyData.ToString().Equals("Down")))
                {
                    down_press = down_press + 1;
                    if (down_press == 1)
                    {
                        up_track = track_F2;
                        track_F2 = s.Length;
                    }

                }

                else if ((keyData.ToString().Equals("Up")))
                {
                    if (down_press != 0)
                        track_F2 = up_track;
                }
                Char code;
                //MessageBox.Show(keyData);
                if (keyData.ToString() == "Space")
                    code = ' ';

                else if (keyData.ToString() == "D1")
                    code = '!';

                else if (keyData.ToString() == "D2")
                    code = '@';

                else if (keyData.ToString() == "D3")
                    code = '#';

                else if (keyData.ToString() == "D4")
                    code = '$';

                else if (keyData.ToString() == "D5")
                    code = '%';

                else if (keyData.ToString() == "D6")
                    code = '^';

                else if (keyData.ToString() == "D7")
                    code = '&';

                else if (keyData.ToString() == "D8")
                    code = '*';

                else if (keyData.ToString() == "D9")
                    code = '(';

                else if (keyData.ToString() == "D0")
                    code = ')';

                else if (keyData.ToString() == "Oemcomma")
                    code = '<';

                else if (keyData.ToString() == "OemPeriod")
                    code = '>';
                else if (keyData.ToString() == "OemQuestion")
                    code = '?';
                else if (keyData.ToString() == "Oem1")
                    code = ':';
                else if (keyData.ToString() == "Oem7")
                    code = '"';

                else if (keyData.ToString() == "OemMinus")
                    code = '_';
                else if (keyData.ToString() == "Oemplus")
                    code = '+';


                else if (keyData.ToString() == "OemOpenBrackets")
                    code = '{';
                else if (keyData.ToString() == "Oem6")
                    code = '}';
                else if (keyData.ToString() == "Oemtilde")
                    code = '~';
                else if (keyData.ToString() == "Oem5")
                    code = '|';



                else if (keyData.ToString() == "Shift")
                    return;
                else
                    code = Convert.ToChar(keyData);
                /////////////MessageBox.Show(code.ToString());


                if ((code >= 'A' && code <= 'Z') || (code == '"') || (code == '?') || (code == ':') || (code == ' ') || (code == '<') || (code == '>') || code == '!' || code == '@' || code == '#' || code == '$' || code == '%' || code == '^' || code == '&' || code == '*' || code == '(' || code == ')' || code == '_' || code == '+' || code == '(' || code == ')' || code == '_' || code == '+' || code == '(' || code == ')' || code == '_' || code == '+' || code == '{' || code == '}' || code == '|' || code == '~')
                {
                    char[] aa = s.ToCharArray();

                    int i = 0;

                    for (int j = 0; j < track_F2; j++)
                        sort = sort + s[j];

                    for (i = track_F2; i < s.Length; i++)
                    {
                        sort = sort + code;
                        a = s[i];
                        aa[i] = code;
                        //MessageBox.Show(aa[i].ToString());
                        code = a;

                    }

                    sort = sort + code;
                    s = sort;
                    //MessageBox.Show(s);
                    track_F2 = track_F2 + 1;
                    up_track = track_F2;
                }


                track_alt_D1 = 0;
                track_alt_ctrl_D1 = 0;

            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }
        }



        public void HeaderInstruction() ///// প্রতিটা charecter operate করার জন্য।
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

                    if ((keyData.ToString().Equals("E")))
                    {
                        int track;
                        String s;
                        int g = excel.ActiveCell.Cells.Column;

                        for (int i = 1; ; i++)
                        {
                            Excel.Range r = (Excel.Range)excel.Cells.get_Item(i, g);
                            String header = r.Text.ToString();
                            if (header != "")
                            {
                                track = i;
                                s = header;
                                break;
                            }
                        }
                        int y = excel.ActiveCell.Cells.Row;
                        int x = excel.ActiveCell.Cells.Column;

                        string row = "";
                        int div = 0, per = 0;
                        int q;
                        q = x / (26 * 26);
                        if (q == 1) row = "A";
                        else if (q == 2) row = "B";
                        else if (q == 3) row = "C";
                        else if (q == 4) row = "D";
                        else if (q == 5) row = "E";
                        else if (q == 6) row = "F";
                        else if (q == 7) row = "G";
                        else if (q == 8) row = "H";
                        else if (q == 9) row = "I";
                        else if (q == 10) row = "J";
                        else if (q == 11) row = "K";
                        else if (q == 12) row = "L";
                        else if (q == 13) row = "M";
                        else if (q == 14) row = "N";
                        else if (q == 15) row = "O";
                        else if (q == 16) row = "P";
                        else if (q == 17) row = "Q";
                        else if (q == 18) row = "R";
                        else if (q == 19) row = "S";
                        else if (q == 20) row = "T";
                        else if (q == 21) row = "U";
                        else if (q == 22) row = "V";
                        else if (q == 23) row = "W";
                        else if (q == 24) row = "X";
                        else if (q == 25) row = "Y";
                        else if (q == 26) row = "Z";

                        x = x - (26 * 26 * q);
                        div = x / 26;
                        per = x % 26;
                        //Console.WriteLine(div);
                        //Console.WriteLine(per);
                        if (div == 1) row = row + "A";
                        else if (div == 2) row = row + "B";
                        else if (div == 3) row = row + "C";
                        else if (div == 4) row = row + "D";
                        else if (div == 5) row = row + "E";
                        else if (div == 6) row = row + "F";
                        else if (div == 7) row = row + "G";
                        else if (div == 8) row = row + "H";
                        else if (div == 9) row = row + "I";
                        else if (div == 10) row = row + "J";
                        else if (div == 11) row = row + "K";
                        else if (div == 12) row = row + "L";
                        else if (div == 13) row = row + "M";
                        else if (div == 14) row = row + "N";
                        else if (div == 15) row = row + "O";
                        else if (div == 16) row = row + "P";
                        else if (div == 17) row = row + "Q";
                        else if (div == 18) row = row + "R";
                        else if (div == 19) row = row + "S";
                        else if (div == 20) row = row + "T";
                        else if (div == 21) row = row + "U";
                        else if (div == 22) row = row + "V";
                        else if (div == 23) row = row + "W";
                        else if (div == 24) row = row + "X";
                        else if (div == 25) row = row + "Y";
                        else if (div == 26) row = row + "Z";

                        if (per == 1) row = row + "A";
                        else if (per == 2) row = row + "B";
                        else if (per == 3) row = row + "C";
                        else if (per == 4) row = row + "D";
                        else if (per == 5) row = row + "E";
                        else if (per == 6) row = row + "F";
                        else if (per == 7) row = row + "G";
                        else if (per == 8) row = row + "H";
                        else if (per == 9) row = row + "I";
                        else if (per == 10) row = row + "J";
                        else if (per == 11) row = row + "K";
                        else if (per == 12) row = row + "L";
                        else if (per == 13) row = row + "M";
                        else if (per == 14) row = row + "N";
                        else if (per == 15) row = row + "O";
                        else if (per == 16) row = row + "P";
                        else if (per == 17) row = row + "Q";
                        else if (per == 18) row = row + "R";
                        else if (per == 19) row = row + "S";
                        else if (per == 20) row = row + "T";
                        else if (per == 21) row = row + "U";
                        else if (per == 22) row = row + "V";
                        else if (per == 23) row = row + "W";
                        else if (per == 24) row = row + "X";
                        else if (per == 25) row = row + "Y";
                        else if (per == 26) row = row + "Z";

                        speaker.speak("Top Heading " + s);
                        speaker.speak(row + " " + track + "Position");
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


        public void GetColumnText() ///// 
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
                    //
                    Excel.Range rng = excel.ActiveCell;
                    //sheet = (Excel.Worksheet)book.ActiveSheet;
                    //String ss = sheet.Name.ToString(); //return sheet name
                    //String ss = sheet.Index.ToString(); // return sheet index
                    int y = excel.ActiveCell.Cells.Row;
                    int x = excel.ActiveCell.Cells.Column;
                    //speaker.speak("Row " + x + " Column " + y);
                    //speaker.speak(rng.Text.ToString());

                    string row = "";
                    int div = 0, per = 0;
                    int q;
                    q = x / (26 * 26);
                    if (q == 1) row = "A";
                    else if (q == 2) row = "B";
                    else if (q == 3) row = "C";
                    else if (q == 4) row = "D";
                    else if (q == 5) row = "E";
                    else if (q == 6) row = "F";
                    else if (q == 7) row = "G";
                    else if (q == 8) row = "H";
                    else if (q == 9) row = "I";
                    else if (q == 10) row = "J";
                    else if (q == 11) row = "K";
                    else if (q == 12) row = "L";
                    else if (q == 13) row = "M";
                    else if (q == 14) row = "N";
                    else if (q == 15) row = "O";
                    else if (q == 16) row = "P";
                    else if (q == 17) row = "Q";
                    else if (q == 18) row = "R";
                    else if (q == 19) row = "S";
                    else if (q == 20) row = "T";
                    else if (q == 21) row = "U";
                    else if (q == 22) row = "V";
                    else if (q == 23) row = "W";
                    else if (q == 24) row = "X";
                    else if (q == 25) row = "Y";
                    else if (q == 26) row = "Z";

                    x = x - (26 * 26 * q);


                    div = x / 26;
                    per = x % 26;
                    Console.WriteLine(div);
                    Console.WriteLine(per);
                    if (div == 1) row = row + "A";
                    else if (div == 2) row = row + "B";
                    else if (div == 3) row = row + "C";
                    else if (div == 4) row = row + "D";
                    else if (div == 5) row = row + "E";
                    else if (div == 6) row = row + "F";
                    else if (div == 7) row = row + "G";
                    else if (div == 8) row = row + "H";
                    else if (div == 9) row = row + "I";
                    else if (div == 10) row = row + "J";
                    else if (div == 11) row = row + "K";
                    else if (div == 12) row = row + "L";
                    else if (div == 13) row = row + "M";
                    else if (div == 14) row = row + "N";
                    else if (div == 15) row = row + "O";
                    else if (div == 16) row = row + "P";
                    else if (div == 17) row = row + "Q";
                    else if (div == 18) row = row + "R";
                    else if (div == 19) row = row + "S";
                    else if (div == 20) row = row + "T";
                    else if (div == 21) row = row + "U";
                    else if (div == 22) row = row + "V";
                    else if (div == 23) row = row + "W";
                    else if (div == 24) row = row + "X";
                    else if (div == 25) row = row + "Y";
                    else if (div == 26) row = row + "Z";


                    if (per == 1) row = row + "A";
                    else if (per == 2) row = row + "B";
                    else if (per == 3) row = row + "C";
                    else if (per == 4) row = row + "D";
                    else if (per == 5) row = row + "E";
                    else if (per == 6) row = row + "F";
                    else if (per == 7) row = row + "G";
                    else if (per == 8) row = row + "H";
                    else if (per == 9) row = row + "I";
                    else if (per == 10) row = row + "J";
                    else if (per == 11) row = row + "K";
                    else if (per == 12) row = row + "L";
                    else if (per == 13) row = row + "M";
                    else if (per == 14) row = row + "N";
                    else if (per == 15) row = row + "O";
                    else if (per == 16) row = row + "P";
                    else if (per == 17) row = row + "Q";
                    else if (per == 18) row = row + "R";
                    else if (per == 19) row = row + "S";
                    else if (per == 20) row = row + "T";
                    else if (per == 21) row = row + "U";
                    else if (per == 22) row = row + "V";
                    else if (per == 23) row = row + "W";
                    else if (per == 24) row = row + "X";
                    else if (per == 25) row = row + "Y";
                    else if (per == 26) row = row + "Z";


                    //
                    occupiedBuffer = 1;

                    track_alt_D1 = track_alt_D1 + 1;
                    track_alt_D1 = track_alt_D1 % 4;

                    if (track_alt_D1 == 1)
                    {

                        int y1 = excel.ActiveCell.Cells.Column;

                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(1, y1);
                        String header = r.Text.ToString();
                        if (header == "")
                            speaker.speak(row + " 1 " + "Blank");
                        else
                            speaker.speak(row + " 1 " + header);

                    }

                    else if (track_alt_D1 == 2)
                    {

                        int y1 = excel.ActiveCell.Cells.Column;

                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(2, y1);
                        String header = r.Text.ToString();
                        if (header == "")
                            speaker.speak(row + " 2 " + "Blank");
                        else
                            speaker.speak(row + " 2 " + header);
                    }

                    else if (track_alt_D1 == 3)
                    {
                        int y1 = excel.ActiveCell.Cells.Column;

                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(3, y1);
                        String header = r.Text.ToString();
                        if (header == "")
                            speaker.speak(row + " 3 " + "Blank");
                        else
                            speaker.speak(row + " 3 " + header);

                    }

                    else if (track_alt_D1 == 0)
                    {
                        int y1 = excel.ActiveCell.Cells.Column;

                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(4, y1);
                        String header = r.Text.ToString();
                        if (header == "")
                            speaker.speak(row + " 4 " + "Blank");
                        else
                            speaker.speak(row + " 4 " + header);

                    }


                    occupiedBuffer = 0;

                    Monitor.Pulse(this);
                    Monitor.Exit(this);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void GetRowText() ///// 
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

                    track_alt_ctrl_D1 = track_alt_ctrl_D1 + 1;
                    track_alt_ctrl_D1 = track_alt_ctrl_D1 % 4;

                    if (track_alt_ctrl_D1 == 1)
                    {

                        int x = excel.ActiveCell.Cells.Row;

                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(x, 1);
                        String header = r.Text.ToString();
                        if (header == "")
                            speaker.speak(" A " + x + "      " + "Blank");
                        else
                            speaker.speak(" A " + x + "      " + header);

                    }

                    else if (track_alt_ctrl_D1 == 2)
                    {

                        int x = excel.ActiveCell.Cells.Row;

                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(x, 2);
                        String header = r.Text.ToString();
                        if (header == "")
                            speaker.speak("B " + x + "      " + "Blank");
                        else
                            speaker.speak("B " + x + "      " + header);

                    }

                    else if (track_alt_ctrl_D1 == 3)
                    {


                        int x = excel.ActiveCell.Cells.Row;

                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(x, 3);
                        String header = r.Text.ToString();
                        if (header == "")
                            speaker.speak("C " + x + "      " + "Blank");
                        else
                            speaker.speak("C " + x + "      " + header);

                    }

                    else if (track_alt_ctrl_D1 == 0)
                    {

                        int x = excel.ActiveCell.Cells.Row;

                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(x, 4);
                        String header = r.Text.ToString();
                        if (header == "")
                            speaker.speak("D " + x + "      " + "Blank");
                        else
                            speaker.speak("D " + x + "      " + header);

                    }


                    occupiedBuffer = 0;

                    Monitor.Pulse(this);
                    Monitor.Exit(this);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        public void GetLeftText() ///// 
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

                    int a = excel.ActiveCell.Cells.Row; // 7
                    int b = excel.ActiveCell.Cells.Column; // 5
                    String header = "";
                    for (int i = 1; i <= b; i++)
                    {
                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(a, i);
                        header = r.Text.ToString();
                        if (header == "")
                        {
                            continue;
                        }
                        else
                            speaker.speak("    " + header);
                    }
                    if (header == "")
                    {
                        speaker.speak(" blank ");
                    }

                    occupiedBuffer = 0;

                    Monitor.Pulse(this);
                    Monitor.Exit(this);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void GetRightText() ///// 
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

                    int b = excel.ActiveCell.Cells.Column;
                    int tra = 0;
                    String s;
                    int g = excel.ActiveCell.Cells.Row;
                    //MessageBox.Show(g.ToString()); //4
                    for (int i = 256; i >= 0; i--)
                    {
                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(g, i);
                        String header = r.Text.ToString();
                        if (header != "")
                        {
                            tra = i;
                            s = header;
                            break;
                        }
                    }

                    //MessageBox.Show("b=" + b + " tra=" + tra);

                    for (int i = b; i <= tra; i++)
                    {
                        Excel.Range r = (Excel.Range)excel.Cells.get_Item(g, i);
                        String header = r.Text.ToString();
                        if (header == "")
                        {
                            continue;
                        }
                        else
                            speaker.speak("    " + header);
                    }


                    //MessageBox.Show(tra.ToString());

                    //int a = excel.ActiveCell.Cells.Row; // 7
                    //int b = excel.ActiveCell.Cells.Column; // 5

                    //for (int i = 1; i <= b; i++)
                    //{
                    //    Excel.Range r = (Excel.Range)excel.Cells.get_Item(a, i);
                    //    String header = r.Text.ToString();
                    //    if (header == "")
                    //    {
                    //        speaker.speak("Current Cell column " + i + " blank ");
                    //    }
                    //    else
                    //        speaker.speak("Current Cell column " + i + "  " + header);
                    //}


                    occupiedBuffer = 0;

                    Monitor.Pulse(this);
                    Monitor.Exit(this);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void SheetInstruction()
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
                    //MessageBox.Show("keyData");
                    if (keyData.ToString().Equals("Next") || keyData.ToString().Equals("PageUp"))
                    {
                        sheet = (Excel._Worksheet)excel.ActiveSheet;
                        //sheet = (Excel._Worksheet)book.ActiveSheet;
                        String ss = sheet.Name.ToString(); //return sheet name
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

        public void SelectedCurrentRow()
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
                    //MessageBox.Show(keyData);
                    if (keyData.ToString().Equals("Space"))
                    {
                        int x = excel.ActiveCell.Cells.Row;
                        speaker.speak("Current Selected Row is " + x.ToString());
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



        public void SelectedCurrentColumn()
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
                    //MessageBox.Show(keyData);
                    if (keyData.ToString().Equals("Space"))
                    {
                        int x = excel.ActiveCell.Cells.Column;

                        //speaker.speak("Row " + x + " Column " + y);
                        //speaker.speak(rng.Text.ToString());
                        string row = "";
                        int div = 0, per = 0;
                        int q;
                        q = x / (26 * 26);
                        if (q == 1) row = "A";
                        else if (q == 2) row = "B";
                        else if (q == 3) row = "C";
                        else if (q == 4) row = "D";
                        else if (q == 5) row = "E";
                        else if (q == 6) row = "F";
                        else if (q == 7) row = "G";
                        else if (q == 8) row = "H";
                        else if (q == 9) row = "I";
                        else if (q == 10) row = "J";
                        else if (q == 11) row = "K";
                        else if (q == 12) row = "L";
                        else if (q == 13) row = "M";
                        else if (q == 14) row = "N";
                        else if (q == 15) row = "O";
                        else if (q == 16) row = "P";
                        else if (q == 17) row = "Q";
                        else if (q == 18) row = "R";
                        else if (q == 19) row = "S";
                        else if (q == 20) row = "T";
                        else if (q == 21) row = "U";
                        else if (q == 22) row = "V";
                        else if (q == 23) row = "W";
                        else if (q == 24) row = "X";
                        else if (q == 25) row = "Y";
                        else if (q == 26) row = "Z";

                        x = x - (26 * 26 * q);

                        div = x / 26;
                        per = x % 26;
                        Console.WriteLine(div);
                        Console.WriteLine(per);
                        if (div == 1) row = row + "A";
                        else if (div == 2) row = row + "B";
                        else if (div == 3) row = row + "C";
                        else if (div == 4) row = row + "D";
                        else if (div == 5) row = row + "E";
                        else if (div == 6) row = row + "F";
                        else if (div == 7) row = row + "G";
                        else if (div == 8) row = row + "H";
                        else if (div == 9) row = row + "I";
                        else if (div == 10) row = row + "J";
                        else if (div == 11) row = row + "K";
                        else if (div == 12) row = row + "L";
                        else if (div == 13) row = row + "M";
                        else if (div == 14) row = row + "N";
                        else if (div == 15) row = row + "O";
                        else if (div == 16) row = row + "P";
                        else if (div == 17) row = row + "Q";
                        else if (div == 18) row = row + "R";
                        else if (div == 19) row = row + "S";
                        else if (div == 20) row = row + "T";
                        else if (div == 21) row = row + "U";
                        else if (div == 22) row = row + "V";
                        else if (div == 23) row = row + "W";
                        else if (div == 24) row = row + "X";
                        else if (div == 25) row = row + "Y";
                        else if (div == 26) row = row + "Z";


                        if (per == 1) row = row + "A";
                        else if (per == 2) row = row + "B";
                        else if (per == 3) row = row + "C";
                        else if (per == 4) row = row + "D";
                        else if (per == 5) row = row + "E";
                        else if (per == 6) row = row + "F";
                        else if (per == 7) row = row + "G";
                        else if (per == 8) row = row + "H";
                        else if (per == 9) row = row + "I";
                        else if (per == 10) row = row + "J";
                        else if (per == 11) row = row + "K";
                        else if (per == 12) row = row + "L";
                        else if (per == 13) row = row + "M";
                        else if (per == 14) row = row + "N";
                        else if (per == 15) row = row + "O";
                        else if (per == 16) row = row + "P";
                        else if (per == 17) row = row + "Q";
                        else if (per == 18) row = row + "R";
                        else if (per == 19) row = row + "S";
                        else if (per == 20) row = row + "T";
                        else if (per == 21) row = row + "U";
                        else if (per == 22) row = row + "V";
                        else if (per == 23) row = row + "W";
                        else if (per == 24) row = row + "X";
                        else if (per == 25) row = row + "Y";
                        else if (per == 26) row = row + "Z";

                        speaker.speak("Current Selected Column is " + row.ToString());
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
                Monitor.Enter(this);
                if (occupiedBuffer == 1)
                {
                    Monitor.Wait(this);
                }
                else
                {

                    occupiedBuffer = 1;
                    excel.CommandBars.ReleaseFocus();
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
        public void FF2()
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
                    Excel.EditBox r = (Excel.EditBox)excel.Selection;
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
