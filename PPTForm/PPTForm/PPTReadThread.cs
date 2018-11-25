using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PPT = Microsoft.Office.Interop.PowerPoint;
using SpeechBuilder;
using System.Threading;
using System.Windows.Forms;
using System.Text.RegularExpressions;
namespace PPTForm
{
    class PPTReadThread
    {
        Microsoft.Office.Interop.PowerPoint.Application application = null;
        Microsoft.Office.Interop.PowerPoint._Presentation presentation = null;
        Microsoft.Office.Interop.PowerPoint._Slide slide = null;
        Microsoft.Office.Interop.PowerPoint.Slides slides = null;
        //private static Microsoft.Office.Interop.Word.Application wd = null;

        static int k = 0;
        static int ct;
        private String keyData = null;
        private int occupiedBuffer = 0;
        private static int position = 0;

        //Object unitWord = Excel.;
        //Object unitSentence = Word.WdUnits.wdSentence;
        //Object unitParagraph = Word.WdUnits.wdParagraph;       
        //Object unitFullPage = Word.WdUnits.wdFullPage;

        object newTemplate = false;
        object docType = 0;
        object isVisible = true;
        static object p = null;
        //private Form2 f2;

        Object count = 1;
        //Object extend = Word.WdMovementType.wdMove;
        private SpeechControl speaker;

        public PPTReadThread(SpeechControl speaker, Microsoft.Office.Interop.PowerPoint.Application application,
            Microsoft.Office.Interop.PowerPoint._Presentation presentation, String keyData)
        {
            this.application = application;
            //p = excel;
            this.presentation = presentation;
            this.keyData = keyData;
            //speaker = new SpeechControl();
            this.speaker = speaker;

        }
        public void operateCharacter()
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
                        //MessageBox.Show("jjjjjjjjjjjjjjj");
                        PPT.Selection r = application.ActiveWindow.Selection;

                        //MessageBox.Show(r.Type.ToString());

                        if (r.Type.ToString() == "ppSelectionText")
                        {
                            PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                            if (sr.Type != Microsoft.Office.Core.MsoShapeType.msoTable)
                            {
                                //speaker.speak("Powerpoint Selection is on Text");
                                position = r.TextRange.Start;
                                PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                                string s = f.TextRange.Characters(position, 1).Text.ToString();
                                //MessageBox.Show(s);
                                char p = Convert.ToChar(s);
                                int a = Convert.ToInt32(p);

                                if (a >= 65 && a <= 90)
                                {
                                    speaker.speak("Capital " + s);
                                }
                                else if (s == " ")
                                    speaker.speak("space");
                                else if (s == ";")
                                    speaker.speak("semicolon");
                                //
                                else if (s == "!")
                                    speaker.speak("Exclamation mark");
                                else if (s == "(")
                                    speaker.speak("first bracker Open");
                                else if (s == ")")
                                    speaker.speak("first bracker Close");
                                else if (s == ",")
                                    speaker.speak("Comma");
                                else if (s == "-")
                                    speaker.speak("Hyphen");
                                else if (s == ".")
                                    speaker.speak(" full stop");
                                else if (s == ":")
                                    speaker.speak("Colon");
                                else if (s == ";")
                                    speaker.speak("Semicolon");
                                else if (s == "?")
                                    speaker.speak("Question mark");
                                else if (s == "[")
                                    speaker.speak("third bracket Open");
                                else if (s == "]")
                                    speaker.speak("third bracket Close");
                                else if (s == "`")
                                    speaker.speak("Grave accent");
                                else if (s == "{")
                                    speaker.speak("Second bracket Open");
                                else if (s == "}")
                                    speaker.speak("Second bracket Close");
                                else if (s == "~")
                                    speaker.speak("Equivalency sign ");
                                else if (s == "‘" || s == "'" || s == "’")
                                    speaker.speak("Single quotation");
                                else if (s == "”" || s == "“" || s == "\"")
                                    speaker.speak("Double quotes");
                                //
                                else
                                    speaker.speak(f.TextRange.Characters(position, 1).Text.ToString());

                                //r.TextRange.                           
                                //application.ActiveWindow.Selection.)
                                //MessageBox.Show(application.ActiveWindow.Selection.TextRange.Start.ToString());
                            }
                        }
                        else if (r.Type.ToString() == "ppSelectionShapes")
                        {
                            PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                            //speaker.speak("Powerpoint Selection is on Shapes");
                            //PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                            if (sr.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                float left = r.ShapeRange.Left / 3;
                                float width = r.ShapeRange.Width / 3;

                                speaker.speak("Powerpoint Selection is on Text Frame " + "blank space on left " + left.ToString() + " milli-meter and Text Frame Width is " + width.ToString() + " milli-meter");
                            }
                            else if (sr.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                speaker.speak("Powerpoint Selection is on Table");

                                //if (sr.Table.Cell(1, 1).Selected)
                                //{
                                //    MessageBox.Show("ddfd");
                                //}
                            }
                            //MessageBox.Show(sr.);
                            //speaker.speak("Top "+sr.TextFrame.MarginTop.ToString()+" "+"Left "+sr.TextFrame.MarginLeft.ToString()+" "+ "Right "+ sr.TextFrame.MarginRight.ToString());

                        }
                        else if (r.Type.ToString() == "ppSelectionSlides" || r.Type.ToString() == "ppSelectionNone")
                        {
                            PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                            if (rr.Shapes.HasTitle == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                speaker.speak(rr.SlideNumber.ToString());
                                speaker.speak(rr.Shapes.Title.TextFrame.TextRange.Text.ToString() + "  Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                            }
                            else
                            {
                                speaker.speak(rr.SlideNumber.ToString());
                                speaker.speak("Slide " + rr.SlideNumber.ToString());
                                speaker.speak(rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                            }

                            //PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                            //speaker.speak("Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                            ////MessageBox.Show(rr.SlideNumber.ToString());
                            ////MessageBox.Show(rr.SlideID.ToString());
                            //var slide = (PPT.Slide)application.ActiveWindow.View.Slide;

                            //foreach (PPT.Shape shape in slide.Shapes)
                            //{

                            //    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            //    {
                            //        var textFrame = shape.TextFrame.TextRange.Text;
                            //        speaker.speak(textFrame);
                            //        break;
                            //    }
                            //}
                        }
                        //MessageBox.Show(r.Type.ToString());
                    }
                    else if ((keyData.ToString().Equals("Delete")))
                    {
                        //MessageBox.Show("jjjjjjjjjjjjjjj");
                        PPT.Selection r = application.ActiveWindow.Selection;

                        //MessageBox.Show(r.Type.ToString());

                        if (r.Type.ToString() == "ppSelectionText")
                        {
                            PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                            if (sr.Type != Microsoft.Office.Core.MsoShapeType.msoTable)
                            {
                                //speaker.speak("Powerpoint Selection is on Text");
                                position = r.TextRange.Start;
                                PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                                string s = f.TextRange.Characters(position, 1).Text.ToString();
                                //MessageBox.Show(s);
                                char p = Convert.ToChar(s);
                                int a = Convert.ToInt32(p);

                                if (a >= 65 && a <= 90)
                                {
                                    speaker.speak("Capital " + s);
                                }
                                else if (s == " ")
                                    speaker.speak("space");
                                else if (s == ";")
                                    speaker.speak("semicolon");
                                //
                                else if (s == "!")
                                    speaker.speak("Exclamation mark");
                                else if (s == "(")
                                    speaker.speak("first bracker Open");
                                else if (s == ")")
                                    speaker.speak("first bracker Close");
                                else if (s == ",")
                                    speaker.speak("Comma");
                                else if (s == "-")
                                    speaker.speak("Hyphen");
                                else if (s == ".")
                                    speaker.speak(" full stop");
                                else if (s == ":")
                                    speaker.speak("Colon");
                                else if (s == ";")
                                    speaker.speak("Semicolon");
                                else if (s == "?")
                                    speaker.speak("Question mark");
                                else if (s == "[")
                                    speaker.speak("third bracket Open");
                                else if (s == "]")
                                    speaker.speak("third bracket Close");
                                else if (s == "`")
                                    speaker.speak("Grave accent");
                                else if (s == "{")
                                    speaker.speak("Second bracket Open");
                                else if (s == "}")
                                    speaker.speak("Second bracket Close");
                                else if (s == "~")
                                    speaker.speak("Equivalency sign ");
                                else if (s == "‘" || s == "'" || s == "’")
                                    speaker.speak("Single quotation");
                                else if (s == "”" || s == "“" || s == "\"")
                                    speaker.speak("Double quotes");
                                //
                                else
                                    speaker.speak(f.TextRange.Characters(position, 1).Text.ToString());

                                //r.TextRange.                           
                                //application.ActiveWindow.Selection.)
                                //MessageBox.Show(application.ActiveWindow.Selection.TextRange.Start.ToString());
                            }
                        }
                    }
                    else if ((keyData.ToString().Equals("Home")))
                    {
                        PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                        speaker.speak("Top of File Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());

                    }
                    else if ((keyData.ToString().Equals("End")))
                    {
                        PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                        speaker.speak("Bottom of File Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());

                    }

                    else if ((keyData.ToString().Equals("Left")))
                    {
                        PPT.Selection r = application.ActiveWindow.Selection;

                        if (r.Type.ToString() == "ppSelectionText")
                        {
                            PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                            if (sr.Type != Microsoft.Office.Core.MsoShapeType.msoTable)
                            {
                                position = r.TextRange.Start;
                                PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                                string s = f.TextRange.Characters(position, 1).Text.ToString();
                                //MessageBox.Show(s);
                                char p = Convert.ToChar(s);
                                int a = Convert.ToInt32(p);

                                if (a >= 65 && a <= 90)
                                {
                                    speaker.speak("Capital " + s);
                                }
                                else if (s == " ")
                                    speaker.speak("space");
                                else if (s == ";")
                                    speaker.speak("semicolon");
                                //
                                else if (s == "!")
                                    speaker.speak("Exclamation mark");
                                else if (s == "(")
                                    speaker.speak("first bracker Open");
                                else if (s == ")")
                                    speaker.speak("first bracker Close");
                                else if (s == ",")
                                    speaker.speak("Comma");
                                else if (s == "-")
                                    speaker.speak("Hyphen");
                                else if (s == ".")
                                    speaker.speak(" full stop");
                                else if (s == ":")
                                    speaker.speak("Colon");
                                else if (s == ";")
                                    speaker.speak("Semicolon");
                                else if (s == "?")
                                    speaker.speak("Question mark");
                                else if (s == "[")
                                    speaker.speak("third bracket Open");
                                else if (s == "]")
                                    speaker.speak("third bracket Close");
                                else if (s == "`")
                                    speaker.speak("Grave accent");
                                else if (s == "{")
                                    speaker.speak("Second bracket Open");
                                else if (s == "}")
                                    speaker.speak("Second bracket Close");
                                else if (s == "~")
                                    speaker.speak("Equivalency sign ");
                                else if (s == "‘" || s == "'" || s == "’")
                                    speaker.speak("Single quotation");
                                else if (s == "”" || s == "“" || s == "\"")
                                    speaker.speak("Double quotes");
                                //
                                else
                                    speaker.speak(f.TextRange.Characters(r.TextRange.Start, 1).Text.ToString());
                            }
                        }
                        else if (r.Type.ToString() == "ppSelectionShapes")
                        {
                            PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                            speaker.speak("Powerpoint Selection is on Shapes");
                            //PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                            //MessageBox.Show(sr.);
                            //speaker.speak("Top "+sr.TextFrame.MarginTop.ToString()+" "+"Left "+sr.TextFrame.MarginLeft.ToString()+" "+ "Right "+ sr.TextFrame.MarginRight.ToString());

                        }
                        else if (r.Type.ToString() == "ppSelectionSlides" || r.Type.ToString() == "ppSelectionNone")
                        {
                            PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                            if (rr.Shapes.HasTitle == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                speaker.speak(rr.SlideNumber.ToString());
                                speaker.speak(rr.Shapes.Title.TextFrame.TextRange.Text.ToString() + "  Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                            }
                            else
                            {
                                speaker.speak(rr.SlideNumber.ToString());
                                speaker.speak("Slide " + rr.SlideNumber.ToString());
                                speaker.speak(rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                            }
                            //PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                            //speaker.speak("Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                            ////MessageBox.Show(rr.SlideNumber.ToString());
                            ////MessageBox.Show(rr.SlideID.ToString());
                            //var slide = (PPT.Slide)application.ActiveWindow.View.Slide;

                            //foreach (PPT.Shape shape in slide.Shapes)
                            //{

                            //    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            //    {
                            //        var textFrame = shape.TextFrame.TextRange.Text;
                            //        speaker.speak(textFrame);
                            //        break;
                            //    }
                            //}
                        }

                        //PPT.TextRange s = application.ActiveWindow.Selection.TextRange; // return the selected string
                        //PPT.Selection r = application.ActiveWindow.Selection;
                        //String s = PPT.MsoAnimTextUnitEffect.msoAnimTextUnitEffectByCharacter.ToString();
                        //PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange; // rr.SlideNumber return slide number
                        //PPT.TextRa f= application.ActiveWindow.Selection.TextRange


                        //String character = presentation.Application.ActiveWindow.Selection.TextRange.Characters(1, 2).Text.ToString();
                        //s.ShapeRange sr= presentation.Windows.
                        //presentation = application.ActivePresentation;
                        //slides = presentation.Slides;
                        //slide = (PPT.Slide)application.ActiveWindow.View.Slide;
                        //slide = slides.Add(1, PPT.PpSlideLayout.ppLayoutTitleOnly); add new slide
                        //PPT.Slide sli = (PPT.Slide)application.ActiveWindow.View.Slide;
                        //string s =slides[1].NotesPage.Shapes[1].TextFrame.TextRange.Text;
                        // Start:Give nevigate character
                        //PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                        //MessageBox.Show(f.TextRange.Characters(r.TextRange.Start-1,1).Text.ToString());
                        // end: Give nevigate character
                        //PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                        // sr.Name.ToString() return current shape name

                        //MessageBox.Show(r.TextRange.Characters(1,2).Text.ToString()); 


                    }
                    else if ((keyData.ToString().Equals("PageUp")) || keyData.ToString().Equals("Next"))
                    {
                        //speaker.speak("hhhhhhhhhhhhh");
                        PPT.Selection r = application.ActiveWindow.Selection;

                        //MessageBox.Show(r.Type.ToString());
                        //PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                        ////PPT.
                        ////MessageBox.Show();
                        ////MessageBox.Show(f.TextRange.Start.ToString());


                        if (r.Type.ToString() == "ppSelectionSlides" || r.Type.ToString() == "ppSelectionNone")
                        {
                            PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                            speaker.speak("Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                            //MessageBox.Show(rr.SlideNumber.ToString());
                            //MessageBox.Show(rr.Count.ToString());
                        }
                    }

                    else if ((keyData.ToString().Equals("Up")) || (keyData.ToString().Equals("Down")))
                    {
                        PPT.Selection r = application.ActiveWindow.Selection;

                        //MessageBox.Show(r.Type.ToString());
                        //PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                        ////PPT.
                        ////MessageBox.Show();
                        ////MessageBox.Show(f.TextRange.Start.ToString());


                        if (r.Type.ToString() == "ppSelectionSlides" || r.Type.ToString() == "ppSelectionNone")
                        {
                            PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                            if (rr.Shapes.HasTitle == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                speaker.speak(rr.SlideNumber.ToString());
                                speaker.speak(rr.Shapes.Title.TextFrame.TextRange.Text.ToString() + "  Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                            }
                            else
                            {
                                speaker.speak(rr.SlideNumber.ToString());
                                speaker.speak("Slide " + rr.SlideNumber.ToString());
                                speaker.speak(rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                            }

                            //PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                            //speaker.speak("Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                            ////MessageBox.Show(rr.SlideNumber.ToString());
                            ////MessageBox.Show(rr.Count.ToString());

                            ////var powerpoint = Globals.ThisAddIn.Application;
                            ////var presentation = application.ActivePresentation;

                            //var slide = (PPT.Slide)application.ActiveWindow.View.Slide;

                            //foreach (PPT.Shape shape in slide.Shapes)
                            //{

                            //    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            //    {
                            //        var textFrame = shape.TextFrame.TextRange.Text;
                            //        speaker.speak(textFrame);
                            //        break;
                            //    }
                            //}
                        }
                        else if (r.Type.ToString() == "ppSelectionText")
                        {
                            PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                            if (sr.Type != Microsoft.Office.Core.MsoShapeType.msoTable)
                            {
                                position = r.TextRange.Start;
                            }
                        }
                    }
                    else if ((keyData.ToString().Equals("M")))
                    {
                        PPT.Selection r = application.ActiveWindow.Selection;

                        if (r.Type.ToString() == "ppSelectionSlides" || r.Type.ToString() == "ppSelectionNone")
                        {
                            PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;
                            speaker.speak("Slide " + rr.SlideNumber.ToString());
                            //speaker.speak(rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());                            
                        }
                    }
                    //else if ((keyData.ToString().Equals("Down")))
                    //{
                    //    PPT.Selection r = application.ActiveWindow.Selection;

                    //    if (r.Type.ToString() == "ppSelectionSlides" || r.Type.ToString() == "ppSelectionNone")
                    //    {
                    //        PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                    //        if (rr.Shapes.HasTitle == Microsoft.Office.Core.MsoTriState.msoTrue)
                    //        {
                    //            speaker.speak(rr.SlideNumber.ToString());
                    //            speaker.speak(rr.Shapes.Title.TextFrame.TextRange.Text.ToString() + "  Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                    //        }
                    //        else
                    //        {
                    //            speaker.speak(rr.SlideNumber.ToString());
                    //            speaker.speak("Slide " + rr.SlideNumber.ToString());
                    //            speaker.speak(rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                    //        }

                    //        //rr.Shapes.SelectAll();

                    //        //ss.View.Slide.Shapes.Title.TextFrame.TextRange.Text.ToString();

                    //        //MessageBox.Show(rr.Shapes.Title.TextFrame.TextRange.Text.ToString());

                    //        //var slide = (PPT.Slide)application.ActiveWindow.View.Slide;

                    //        //foreach (PPT.Shape shape in slide.Shapes)
                    //        //{

                    //        //    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    //        //    {
                    //        //        var textFrame = shape.TextFrame.TextRange.Text;
                    //        //        speaker.speak(textFrame);
                    //        //        break;
                    //        //    }
                    //        //}
                    //    }
                    //    else if (r.Type.ToString() == "ppSelectionText")
                    //    {
                    //        PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                    //        if (sr.Type != Microsoft.Office.Core.MsoShapeType.msoTable)
                    //        {
                    //            position = r.TextRange.Start;
                    //        }
                    //    }

                    //}
                    else if ((keyData.ToString().Equals("Tab")))
                    {
                        PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                        PPT.Selection r = application.ActiveWindow.Selection;


                        //MessageBox.Show(r.Type.ToString());
                        if (r.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
                        {
                            position = r.TextRange.Start;
                        }
                        else if (r.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
                        {
                            //MessageBox.Show(sr.Type.ToString());
                            if (sr.Type.ToString() == "msoEmbeddedOLEObject" || sr.Type.ToString() == "msoPicture")
                            {
                                speaker.speak("selection is on picture");
                            }
                            else if (sr.Type == Microsoft.Office.Core.MsoShapeType.msoTable)
                            {
                                speaker.speak("Selection is on Table");

                                String TableData = null;
                                for (int i = 1; i <= sr.Table.Rows.Count; i++)
                                {
                                    for (int j = 1; j <= sr.Table.Columns.Count; j++)
                                    {
                                        TableData = TableData + " Cell " + i + " " + j + " " + sr.Table.Cell(i, j).Shape.TextFrame.TextRange.Text.ToString();

                                    }

                                }
                                speaker.speak(TableData);

                                if (r.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
                                {
                                    //MessageBox.Show(sr.Table.Rows.Count.ToString());
                                }
                                ////if (sr.Table.Cell(1, 2).)
                                ////{
                                ////    MessageBox.Show(sr.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text.ToString());
                                ////}
                            }
                            else
                            {
                                PPT.Selection rt = application.ActiveWindow.Selection;

                                if (sr.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoFalse)
                                {
                                    speaker.speak(sr.Name.ToString() + " Blank");
                                }
                                else if (sr.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                                {
                                    speaker.speak(sr.Name.ToString() + "   " + rt.TextRange.Text.ToString());
                                }
                            }

                        }


                        //else if (r.Type.ToString() == "ppSelectionText")
                        //{
                        //    speaker.speak("Tab");
                        //}                            
                        //else if (sr.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                        //{
                        //    if (sr.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoFalse)
                        //    {
                        //        speaker.speak("Selection is on Blank Place Holder");
                        //    }
                        //    else if(sr.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        //    {
                        //        speaker.speak("Selection is on Place Holder with Text");
                        //    }
                        //}
                        //else if (sr.Type == Microsoft.Office.Core.MsoShapeType.msoAutoShape)
                        //{
                        //    if (sr.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoFalse)
                        //    {
                        //        speaker.speak("Selection is on Blank auto Shape");
                        //    }
                        //    else if (sr.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        //    {
                        //        speaker.speak("Selection is on auto Shape with Text");
                        //    }
                        //}
                        //else if (sr.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                        //{
                        //    if (sr.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoFalse)
                        //    {
                        //        speaker.speak("Selection is on Blank Text Box");
                        //    }
                        //    else if (sr.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        //    {
                        //        speaker.speak("Selection is on Text Box with Text");
                        //    }
                        //}
                        //else if (sr.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                        //{
                        //    speaker.speak("Selection is on Picture Holder");
                        //}


                        ////MessageBox.Show(r.Type.ToString());

                        ////PPT.Selection r = application.ActiveWindow.Selection;

                        ////if (r.Type.ToString() == "ppSelectionShapes")
                        ////{
                        ////    speaker.speak("Powerpoint Selection is on Shapes");
                        ////    PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;

                        ////    //MessageBox.Show(sr.);
                        ////    //speaker.speak("Top "+sr.TextFrame.MarginTop.ToString()+" "+"Left "+sr.TextFrame.MarginLeft.ToString()+" "+ "Right "+ sr.TextFrame.MarginRight.ToString());

                        ////}

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

        public void SlideRd()
        {
            PPT.Selection r = application.ActiveWindow.Selection;

            if (r.Type.ToString() == "ppSelectionSlides" || r.Type.ToString() == "ppSelectionNone")
            {
                PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                if (rr.Shapes.HasTitle == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    speaker.speak(rr.SlideNumber.ToString());
                    speaker.speak(rr.Shapes.Title.TextFrame.TextRange.Text.ToString() + "  Slide " + rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                }
                else
                {
                    speaker.speak(rr.SlideNumber.ToString());
                    speaker.speak("Slide " + rr.SlideNumber.ToString());
                    speaker.speak(rr.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
                }               
            }
            else if (r.Type.ToString() == "ppSelectionText")
            {
                PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                if (sr.Type != Microsoft.Office.Core.MsoShapeType.msoTable)
                {
                    position = r.TextRange.Start;
                }
            }
        }

        public void operateSelection()
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
                    PPT.Selection r = application.ActiveWindow.Selection;
                    if (r.Type.ToString() == "ppSelectionText")
                    {
                        position = r.TextRange.Start;
                        speaker.speak(r.TextRange.Text.ToString());
                        speaker.speak("   selected");
                    }
                    else if (r.Type.ToString() == "ppSelectionShapes")
                    {

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
        public void operateInsideTable()
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
                    PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                    //MessageBox.Show(sr.Type.ToString());
                    //if (sr.Type= Microsoft.Office.Core.MsoShapeType.msoTable)
                    //{
                    //    speaker.speak("Selection is on Text");
                    //    MessageBox.Show(sr.Table.ToString());
                    //}

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
        public void operateFull()
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
                    PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                    PPT.Selection r = application.ActiveWindow.Selection;
                    if (sr.Type == Microsoft.Office.Core.MsoShapeType.msoTable)
                    {

                    }
                    else if (sr.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                    {
                        //speaker.speak("Selection is on Text");
                        speaker.speak("Selection is on Text " + r.TextRange.Text.ToString());
                    }
                    else if (sr.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                    {
                        //speaker.speak("Selection is on Text");
                        speaker.speak("Selection is on Text " + r.TextRange.Text.ToString());
                    }
                    else if (sr.Type == Microsoft.Office.Core.MsoShapeType.msoAutoShape)
                    {
                        //speaker.speak("Selection is on Text");
                        speaker.speak("Selection is on Text " + r.TextRange.Text.ToString());
                    }
                    //if (r.Type.ToString() == "ppSelectionText")
                    //{
                    //    speaker.speak("Selection is on Text");
                    //    speaker.speak(r.TextRange.Text.ToString());
                    //}

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
            //MessageBox.Show()

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
                    PPT.Selection r = application.ActiveWindow.Selection;
                    r.Application.CommandBars.ReleaseFocus(); //release Focus
                    //r.Application.CommandBars.ActiveMenuBar.Type
                    //MessageBox.Show(r.Application.CommandBars.ActiveMenuBar.);
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
        public void ShowControl()
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
                    PPT.SlideShowWindow ss = application.ActivePresentation.SlideShowWindow;
                    ss.View.AcceleratorsEnabled = Microsoft.Office.Core.MsoTriState.msoFalse;

                    String PName = ss.Presentation.Name;
                    //String s = application.ActiveWindow.Presentation.Name;
                    //ss.View.AcceleratorsEnabled = Microsoft.Office.Core.MsoTriState.msoFalse;

                    speaker.speak(" powerpoint slide show dash left bracket  ");
                    speaker.speak(PName + " Right Bracket ");


                    //////////////////////////////////////////////

                    if (ss.View.Slide.Shapes.HasTitle == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        //MessageBox.Show(ss.View.Slide.Shapes.Title.TextFrame.TextRange.Text.ToString());
                        String Title = ss.View.Slide.Shapes.Title.TextFrame.TextRange.Text.ToString();
                        if (Title == "") { speaker.speak("slide 1 slide 1 "); }
                        else speaker.speak(Title + " slide slide 1 ");
                    }
                    if (ss.View.Slide.Shapes.HasTitle != Microsoft.Office.Core.MsoTriState.msoTrue)
                    { speaker.speak(" slide 1 slide 1 "); }

                    // full text


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

        public void CurrentPPTName()
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
                    String s = application.ActiveWindow.Presentation.Name;
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

        public void SlideShow()
        {
            //MessageBox.Show("ssss");
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
                    if (keyData.ToString().Equals("Right") || keyData.ToString().Equals("Left") || keyData.ToString().Equals("Up") || keyData.ToString().Equals("Down"))
                    {
                        String SlideText = null;
                        PPT.SlideShowWindow ss = application.ActivePresentation.SlideShowWindow;

                        ss.View.AcceleratorsEnabled = Microsoft.Office.Core.MsoTriState.msoFalse;

                        //ss.Application.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMaximized;
                        //presentation.SlideShowWindow.Activate();
                        int CSP = ss.View.CurrentShowPosition;
                        //MessageBox.Show(CSP.ToString());
                        application.ActivePresentation.Slides.Application.Activate();
                        application.ActivePresentation.Slides.Range(CSP).Select();
                        //presentation.Slides.Range(CSP).Select();

                        //sl.Application.Activate();
                        PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;
                        //rr.Shapes.Range(1).TextFrame.TextRange.Select();
                        //MessageBox.Show(rr.Shapes.Count.ToString());
                        int ShapeCount = rr.Shapes.Count;

                        for (int i = 1; i <= ShapeCount; i++)
                        {
                            String ShapeText = null;
                            if (rr.Shapes.Range(i).HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {

                                String TableData = null;
                                for (int j = 1; j <= rr.Shapes.Range(i).Table.Rows.Count; j++)
                                {
                                    for (int k = 1; k <= rr.Shapes.Range(i).Table.Columns.Count; k++)
                                    {
                                        TableData = TableData + " Cell " + j + " " + k + " " + rr.Shapes.Range(i).Table.Cell(j, k).Shape.TextFrame.TextRange.Text.ToString();

                                    }

                                }
                                ShapeText = " " + TableData;
                            }
                            else if (rr.Shapes.Range(i).HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                ShapeText = " " + rr.Shapes.Range(i).TextFrame.TextRange.Text.ToString();
                            }
                            else
                                continue;
                            SlideText += ShapeText;

                        }
                        ss.Activate();
                        //MessageBox.Show(SlideText);
                        speaker.stop();
                        speaker.stop();
                        speaker.speak(SlideText);

                    }
                    else if (keyData.ToString().Equals("Back"))
                    {
                        PPT.SlideShowWindow ss = application.ActivePresentation.SlideShowWindow;
                        ss.View.AcceleratorsEnabled = Microsoft.Office.Core.MsoTriState.msoFalse;

                        //if(ss.View.LastSlideViewed)
                        //MessageBox.Show(TS.ToString());
                        ss.View.Previous();
                        speaker.speak("Slide " + ss.View.Slide.SlideNumber.ToString());
                        ////////////////////////////////////////////////////////////////////////
                        String SlideText = null;

                        int CSP = ss.View.CurrentShowPosition;
                        application.ActivePresentation.Slides.Application.Activate();
                        application.ActivePresentation.Slides.Range(CSP).Select();

                        PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;

                        int ShapeCount = rr.Shapes.Count;

                        for (int i = 1; i <= ShapeCount; i++)
                        {
                            String ShapeText = null;
                            if (rr.Shapes.Range(i).HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {

                                String TableData = null;
                                for (int j = 1; j <= rr.Shapes.Range(i).Table.Rows.Count; j++)
                                {
                                    for (int k = 1; k <= rr.Shapes.Range(i).Table.Columns.Count; k++)
                                    {
                                        TableData = TableData + " Cell " + j + " " + k + " " + rr.Shapes.Range(i).Table.Cell(j, k).Shape.TextFrame.TextRange.Text.ToString();

                                    }

                                }
                                ShapeText = " " + TableData;
                            }
                            else if (rr.Shapes.Range(i).HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                ShapeText = " " + rr.Shapes.Range(i).TextFrame.TextRange.Text.ToString();
                            }
                            else
                                continue;
                            SlideText += ShapeText;

                        }
                        ss.Activate();
                        //MessageBox.Show(SlideText);
                        speaker.stop();
                        speaker.stop();
                        speaker.speak(SlideText);
                    }
                    else if (keyData.ToString().Equals("Space"))
                    {

                        PPT.SlideShowWindow ss = application.ActivePresentation.SlideShowWindow;
                        ss.View.AcceleratorsEnabled = Microsoft.Office.Core.MsoTriState.msoFalse;
                        int TS = application.ActivePresentation.Slides.Count;
                        int CSP = ss.View.CurrentShowPosition;
                        if (CSP != TS)
                        {
                            ss.View.Next();
                            speaker.speak("Slide " + ss.View.Slide.SlideNumber.ToString());
                        }
                        else speaker.speak("This is Last Presentation Slide");
                        /////////////////////////////////////////////////////////////
                        String SlideText = null;

                        //int CSP = ss.View.CurrentShowPosition;
                        //MessageBox.Show(CSP.ToString());
                        application.ActivePresentation.Slides.Application.Activate();
                        application.ActivePresentation.Slides.Range(CSP + 1).Select();
                        //presentation.Slides.Range(CSP).Select();

                        //sl.Application.Activate();
                        PPT.SlideRange rr = application.ActiveWindow.Selection.SlideRange;
                        //rr.Shapes.Range(1).TextFrame.TextRange.Select();
                        //MessageBox.Show(rr.Shapes.Count.ToString());
                        int ShapeCount = rr.Shapes.Count;

                        for (int i = 1; i <= ShapeCount; i++)
                        {
                            String ShapeText = null;
                            if (rr.Shapes.Range(i).HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {

                                String TableData = null;
                                for (int j = 1; j <= rr.Shapes.Range(i).Table.Rows.Count; j++)
                                {
                                    for (int k = 1; k <= rr.Shapes.Range(i).Table.Columns.Count; k++)
                                    {
                                        TableData = TableData + " Cell " + j + " " + k + " " + rr.Shapes.Range(i).Table.Cell(j, k).Shape.TextFrame.TextRange.Text.ToString();

                                    }

                                }
                                ShapeText = " " + TableData;
                            }
                            else if (rr.Shapes.Range(i).HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                ShapeText = " " + rr.Shapes.Range(i).TextFrame.TextRange.Text.ToString();
                            }
                            else
                                continue;
                            SlideText += ShapeText;

                        }
                        ss.Activate();
                        //MessageBox.Show(SlideText);
                        speaker.stop();
                        speaker.stop();
                        speaker.speak(SlideText);
                    }

                    else if (keyData.ToString().Equals("PageUp") || keyData.ToString().Equals("P"))
                    {
                        PPT.SlideShowWindow ss = application.ActivePresentation.SlideShowWindow;
                        ss.View.AcceleratorsEnabled = Microsoft.Office.Core.MsoTriState.msoFalse;

                        //if(ss.View.LastSlideViewed)
                        //MessageBox.Show(TS.ToString());
                        ss.View.Previous();
                        speaker.speak("Slide " + ss.View.Slide.SlideNumber.ToString());
                        if (ss.View.Slide.Shapes.HasTitle == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            //MessageBox.Show(ss.View.Slide.Shapes.Title.TextFrame.TextRange.Text.ToString());
                            String Title = ss.View.Slide.Shapes.Title.TextFrame.TextRange.Text.ToString();
                            if (Title == "") speaker.speak("title blank");
                            else speaker.speak(Title);
                        }
                        if (ss.View.Slide.Shapes.HasTitle != Microsoft.Office.Core.MsoTriState.msoTrue)
                            speaker.speak("Title Blank");
                    }
                    else if (keyData.ToString().Equals("Next") || keyData.ToString().Equals("N"))
                    {
                        PPT.SlideShowWindow ss = application.ActivePresentation.SlideShowWindow;
                        ss.View.AcceleratorsEnabled = Microsoft.Office.Core.MsoTriState.msoFalse;
                        int TS = application.ActivePresentation.Slides.Count;
                        int CS = ss.View.CurrentShowPosition;
                        if (CS != TS)
                        {
                            ss.View.Next();
                            speaker.speak("Slide " + ss.View.Slide.SlideNumber.ToString());
                        }
                        else speaker.speak("This is Last Presentation Slide");
                        if (ss.View.Slide.Shapes.HasTitle == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            //MessageBox.Show(ss.View.Slide.Shapes.Title.TextFrame.TextRange.Text.ToString());
                            String Title = ss.View.Slide.Shapes.Title.TextFrame.TextRange.Text.ToString();
                            if (Title == "") speaker.speak("title blank");
                            else speaker.speak(Title);
                        }
                        if (ss.View.Slide.Shapes.HasTitle != Microsoft.Office.Core.MsoTriState.msoTrue)
                            speaker.speak("Title Blank");

                    }
                    occupiedBuffer = 0;
                }
                Monitor.Pulse(this);
                Monitor.Exit(this);
            }
            catch (Exception ex)
            {
                try
                {
                    PPT.SlideShowWindow ss = application.ActivePresentation.SlideShowWindow;
                    ss.Activate();
                }
                catch (Exception exx)
                { }
                //MessageBox.Show(ex.ToString());
            }
        }
        public static class WordCounting
        {
            /// <summary>
            /// Count words with Regex.
            /// </summary>
            public static int CountWords1(string s)
            {
                MatchCollection collection = Regex.Matches(s, @"[\S]+");
                return collection.Count;
            }

            /// <summary>
            /// Count word with loop and character tests.
            /// </summary>
            public static int CountWords2(string s)
            {
                int c = 0;
                for (int i = 1; i < s.Length; i++)
                {
                    if (char.IsWhiteSpace(s[i - 1]) == true)
                    {
                        if (char.IsLetterOrDigit(s[i]) == true ||
                            char.IsPunctuation(s[i]))
                        {
                            c++;
                        }
                    }
                }
                if (s.Length > 2)
                {
                    c++;
                }
                return c;
            }
        }

        public void operateWord()
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
                    //MessageBox.Show("fdfd");
                    if ((keyData.ToString().Equals("Right")))
                    {
                        PPT.Selection r = application.ActiveWindow.Selection;
                        if (r.Type.ToString() == "ppSelectionText")
                        {
                            int position2 = r.TextRange.Start;
                            PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                            //MessageBox.Show("position=" + position + ",position2=" + position2);
                            int len = position2 - position;
                            //speaker.speak(f.TextRange.Characters(position - 1, len).Text.ToString());
                            position = position2;
                            //
                            String s = "", t = "";
                            for (int i = position2; ; i++)
                            {
                                t = f.TextRange.Characters(position2, 1).Text.ToString();
                                int k = 0;
                                foreach (char c in t)
                                {
                                    k = (int)c;
                                    break;
                                }
                                //MessageBox.Show(k.ToString());


                                if (t == " " || k == 13 || t == "")
                                {
                                    speaker.speak(s);
                                    break;
                                }
                                else
                                {
                                    s = s + t;
                                    position2 = position2 + 1;
                                }
                            }
                            //speaker.speak(s);
                            //MessageBox.Show(s);
                            //
                        }
                    }
                    else if ((keyData.ToString().Equals("Left")))
                    {
                        PPT.Selection r = application.ActiveWindow.Selection;
                        if (r.Type.ToString() == "ppSelectionText")
                        {
                            int position2 = r.TextRange.Start;
                            PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                            //MessageBox.Show("position=" + position + ",position2=" + position2);
                            int len = position - position2;
                            speaker.speak(f.TextRange.Characters(position2, len).Text.ToString());
                            position = position2;
                        }

                    }


                    else if ((keyData.ToString().Equals("Up")))
                    {
                        PPT.Selection r = application.ActiveWindow.Selection;
                        if (r.Type.ToString() == "ppSelectionText")
                        {
                            int position2 = r.TextRange.Start;
                            PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                            //MessageBox.Show("position=" + position + ",position2=" + position2);
                            int len = position - position2;
                            speaker.speak(f.TextRange.Characters(position2, len).Text.ToString());
                            string t1 = f.TextRange.Characters(position2, len).Text.ToString();
                            int wordCount = WordCounting.CountWords1(t1);
                            if (wordCount == 0) speaker.speak("Blank");

                            position = position2;
                        }
                        //MessageBox.Show(r.Type.ToString());
                    }
                    else if ((keyData.ToString().Equals("Down")))
                    {
                        PPT.Selection r = application.ActiveWindow.Selection;
                        if (r.Type.ToString() == "ppSelectionText")
                        {
                            int position2 = r.TextRange.Start;
                            PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                            //MessageBox.Show("position=" + position + ",position2=" + position2);
                            int len = position2 - position;
                            //speaker.speak(f.TextRange.Characters(position - 1, len).Text.ToString());
                            position = position2;


                            /////////////////

                            String s = "", t = "", data = "";
                            int leng = 0;
                            for (int i = position2; ; i++)
                            {
                                t = f.TextRange.Characters(position2, 1).Text.ToString();
                                int k = 0;
                                foreach (char c in t)
                                {
                                    k = (int)c;
                                    break;
                                }
                                //MessageBox.Show(k.ToString());

                                if (k == 13 || t == "")
                                {
                                    data = data + s + "   ";
                                    //speaker.speak(s);
                                    //MessageBox.Show("k= " + s);
                                    break;
                                }
                                else if (t == " ")
                                {
                                    //MessageBox.Show("s= space then = " + s);
                                    //leng = s.Length;
                                    data = data + s + "   ";
                                    //speaker.speak(s);
                                    position2 = position2 + 1;
                                    s = "";

                                }
                                else
                                {
                                    s = s + t;
                                    position2 = position2 + 1;
                                    //MessageBox.Show("else s=" + s+" position2 = "+position2);
                                }
                            }
                            speaker.speak(data);


                            string t1 = data;
                            int wordCount = WordCounting.CountWords1(t1);
                            if (wordCount == 0) speaker.speak("Blank");


                            //speaker.speak(s);
                            //MessageBox.Show(s);
                            //

                            ////////////////
                        }
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
        public void ReadWhenPressN()
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
                    //MessageBox.Show("fdfd");
                    //MessageBox.Show("ffffffffff")
                    //speaker.speak("You press control plus n to create a new presentation");          
                    speaker.speak("Presentation Slide 1");
                    speaker.speak(" 1 of 1");
                    speaker.speak(" no selection to select and object press tab");
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

        public void SlideShowSlideNo()
        {
            PPT.SlideShowWindow ss = application.ActivePresentation.SlideShowWindow;
            ss.View.AcceleratorsEnabled = Microsoft.Office.Core.MsoTriState.msoFalse;
            speaker.speak("Slide " + ss.View.Slide.SlideNumber.ToString() + " of " + application.ActivePresentation.Slides.Count.ToString());
        }
        public void OldChar()     ///// প্রতিটা word operate করার জন্য।
        {
            try
            {
                PPT.Selection r = application.ActiveWindow.Selection;

                if (r.Type.ToString() == "ppSelectionText")
                {
                    PPT.ShapeRange sr = application.ActiveWindow.Selection.ShapeRange;
                    if (sr.Type != Microsoft.Office.Core.MsoShapeType.msoTable)
                    {
                        //speaker.speak("Powerpoint Selection is on Text");
                        position = r.TextRange.Start;
                        PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                        string s = f.TextRange.Characters(position, -1).Text.ToString();
                        //MessageBox.Show(s);
                        char p = Convert.ToChar(s);
                        int a = Convert.ToInt32(p);

                        if (a >= 65 && a <= 90)
                        {
                            speaker.speak("Capital " + s);
                        }
                        else if (s == " ")
                            speaker.speak("space");
                        else if (s == ";")
                            speaker.speak("semicolon");
                        //
                        else if (s == "!")
                            speaker.speak("Exclamation mark");
                        else if (s == "(")
                            speaker.speak("first bracker Open");
                        else if (s == ")")
                            speaker.speak("first bracker Close");
                        else if (s == ",")
                            speaker.speak("Comma");
                        else if (s == "-")
                            speaker.speak("Hyphen");
                        else if (s == ".")
                            speaker.speak(" full stop");
                        else if (s == ":")
                            speaker.speak("Colon");
                        else if (s == ";")
                            speaker.speak("Semicolon");
                        else if (s == "?")
                            speaker.speak("Question mark");
                        else if (s == "[")
                            speaker.speak("third bracket Open");
                        else if (s == "]")
                            speaker.speak("third bracket Close");
                        else if (s == "`")
                            speaker.speak("Grave accent");
                        else if (s == "{")
                            speaker.speak("Second bracket Open");
                        else if (s == "}")
                            speaker.speak("Second bracket Close");
                        else if (s == "~")
                            speaker.speak("Equivalency sign ");
                        else if (s == "‘" || s == "'" || s == "’")
                            speaker.speak("Single quotation");
                        else if (s == "”" || s == "“" || s == "\"")
                            speaker.speak("Double quotes");
                        //
                        else
                            speaker.speak(f.TextRange.Characters(position, 1).Text.ToString());

                    }
                }
            }
            catch (Exception ex) { }
        }

        public void OldWord()     ///// প্রতিটা word operate করার জন্য।
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
                    //MessageBox.Show("fdfd");

                    PPT.Selection r = application.ActiveWindow.Selection;
                    if (r.Type.ToString() == "ppSelectionText")
                    {
                        int position2 = r.TextRange.Start;
                        PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                        //MessageBox.Show("position=" + position + ",position2=" + position2);
                        int len = position2 - position;
                        //speaker.speak(f.TextRange.Characters(position - 1, len).Text.ToString());
                        position = position2;
                        /////////////////////////
                        int i, j, wordpos1 = 0;
                        String s = "", t = "", Revstr = "";

                        for (i = position2; i >= 0; i--)
                        {
                            t = f.TextRange.Characters(i, 1).Text.ToString();
                            int k = 0;
                            foreach (char c in t)
                            {
                                k = (int)c;
                                break;
                            }
                            //MessageBox.Show(k.ToString());

                            if (t == " " || k == 13 || t == "")
                            {
                                wordpos1 = i;
                                break;
                            }
                        }
                        //MessageBox.Show(wordpos1.ToString());
                        for (j = wordpos1 - 1; j >= 0; j--)
                        {
                            t = f.TextRange.Characters(j, 1).Text.ToString();
                            int k = 0;
                            foreach (char c in t)
                            {
                                k = (int)c;
                                break;
                            }
                            //MessageBox.Show(t);

                            if (t == " " || k == 13 || t == "")
                            {
                                //MessageBox.Show(s);
                                break;
                            }
                            else
                            {
                                s = s + t;
                            }
                        }

                        for (i = s.Length - 1; i >= 0; i--)
                        {

                            Revstr = Revstr + s[i];
                        }
                        speaker.speak(Revstr);

                        //////////////////////////
                        //speaker.speak(s);
                        //MessageBox.Show(s);
                        //
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

        public void OldPara()     ///// প্রতিটা word operate করার জন্য।
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
                    //MessageBox.Show("fdfd");

                    PPT.Selection r = application.ActiveWindow.Selection;
                    if (r.Type.ToString() == "ppSelectionText")
                    {
                        int position2 = r.TextRange.Start;
                        PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                        //MessageBox.Show("position=" + position + ",position2=" + position2);
                        int len = position2 - position;
                        //speaker.speak(f.TextRange.Characters(position - 1, len).Text.ToString());
                        position = position2;
                        /////////////////////////
                        int i, j, wordpos1 = 0;
                        String s = "", t = "", Revstr = "";

                        for (i = position2; i >= 0; i--)
                        {
                            t = f.TextRange.Characters(i, 1).Text.ToString();
                            int k = 0;
                            foreach (char c in t)
                            {
                                k = (int)c;
                                break;
                            }
                            //MessageBox.Show(k.ToString());

                            if (k == 13 || t == "")
                            {
                                wordpos1 = i;
                                break;
                            }
                        }
                        //MessageBox.Show(wordpos1.ToString());
                        for (j = wordpos1 - 1; j >= 0; j--)
                        {
                            t = f.TextRange.Characters(j, 1).Text.ToString();
                            int k = 0;
                            foreach (char c in t)
                            {
                                k = (int)c;
                                break;
                            }
                            //MessageBox.Show(t);

                            if (k == 13)
                            {
                                //MessageBox.Show(s);
                                break;
                            }
                            else
                            {
                                s = s + t;
                            }
                        }

                        for (i = s.Length - 1; i >= 0; i--)
                        {

                            Revstr = Revstr + s[i];
                        }
                        speaker.speak(Revstr);

                        //////////////////////////
                        //speaker.speak(s);
                        //MessageBox.Show(s);
                        //
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

        public void FontInfo()
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
                    //MessageBox.Show("fdfd");

                    PPT.Selection r = application.ActiveWindow.Selection;
                    if (r.Type.ToString() == "ppSelectionText")
                    {
                        String fColor = "";
                        PPT.TextFrame f = application.ActiveWindow.Selection.ShapeRange.TextFrame;
                        String fN = f.TextRange.Font.Name;
                        String fS = f.TextRange.Font.Size.ToString();
                        String bold = f.TextRange.Font.Bold.ToString(); // msoTrue,msoFalse
                        if (bold == "msoTrue") bold = "bolded";
                        else bold = "";

                        String paragraphAllignment = f.TextRange.ParagraphFormat.Alignment.ToString();
                        if (paragraphAllignment == "ppAlignLeft")
                            paragraphAllignment = " Align Left ";
                        else if (paragraphAllignment == "ppAlignRight")
                            paragraphAllignment = " Align Right ";
                        else if (paragraphAllignment == "ppAlignJustify")
                            paragraphAllignment = " Justified ";
                        else if (paragraphAllignment == "ppAlignCenter")
                            paragraphAllignment = " Centered ";
                        //MessageBox.Show(paragraphAllignment);
                        int rgb = f.TextRange.Font.Color.RGB;
                        //MessageBox.Show(rgb.ToString());

                        if (rgb == 11429888)
                            fColor = " Deep Sky Blue 4 ";
                        else if (rgb == 16777215)
                            fColor = " white ";
                        else if (rgb == 15239680)
                            fColor = "  blue ";
                        else if (rgb == 16772300)
                            fColor = " light still blue ";
                        else if (rgb == 10066176)
                            fColor = " cyan4 ";
                        else if (rgb == 14977024)
                            fColor = " deep sky blue 3 ";
                        else if (rgb == 13875626)
                            fColor = " light still blue 3  ";
                        else if (rgb == 14342874)
                            fColor = " gray 85 ";
                        else if (rgb == 13290154)
                            fColor = "  light cyan 3 ";
                        else if (rgb == 13597440)
                            fColor = " blue "; // 10


                        else if (rgb == 16770236)
                            fColor = " slik gray 1 ";
                        else if (rgb == 15921906)
                            fColor = " gray 95 ";
                        else if (rgb == 16771271)
                            fColor = "  slik gray 1 ";
                        else if (rgb == 16767902)
                            fColor = " light sky blue ";
                        else if (rgb == 16777144)
                            fColor = " pale tark 1 ";
                        else if (rgb == 16771271)
                            fColor = " slik gray 1 ";
                        else if (rgb == 16183790)
                            fColor = "  gray 95 ";
                        else if (rgb == 12895428)
                            fColor = " gray 77 ";
                        else if (rgb == 16053486)
                            fColor = "  gray 95 ";
                        else if (rgb == 16770754)
                            fColor = " silk gray 1 "; // 20

                        else if (rgb == 16763257)
                            fColor = " sky blue 1 ";
                        else if (rgb == 14277081)
                            fColor = " gray 85 ";
                        else if (rgb == 16765584)
                            fColor = "  sky blue 1 ";
                        else if (rgb == 16761177)
                            fColor = " still blue ";
                        else if (rgb == 16777072)
                            fColor = " dark slik gray 2 ";
                        else if (rgb == 16765326)
                            fColor = " sky blue 1 ";
                        else if (rgb == 15590365)
                            fColor = " accent 2  ";
                        else if (rgb == 10724259)
                            fColor = " gray 64 ";
                        else if (rgb == 15395549)
                            fColor = "  accent 2 ";
                        else if (rgb == 16764550)
                            fColor = " sky blue 1 "; // 30

                        else if (rgb == 16756277)
                            fColor = " duzzer blue ";
                        else if (rgb == 12566463)
                            fColor = " gray 75 ";
                        else if (rgb == 16759640)
                            fColor = "  still blue ";
                        else if (rgb == 15044608)
                            fColor = " deep sky blue 3  ";
                        else if (rgb == 16777001)
                            fColor = " cyan  ";
                        else if (rgb == 16759638)
                            fColor = " still blue 1 ";
                        else if (rgb == 15062476)
                            fColor = "  gray 84 ";
                        else if (rgb == 7171437)
                            fColor = " gray 43 ";
                        else if (rgb == 14671820)
                            fColor = "  gray 85 ";
                        else if (rgb == 16758089)
                            fColor = " still blue "; // 40


                        else if (rgb == 8539648)
                            fColor = " duzzer blue 4 ";
                        else if (rgb == 10921638)
                            fColor = " gray 65 ";
                        else if (rgb == 11429888)
                            fColor = "  deep sky blue 4 ";
                        else if (rgb == 7555072)
                            fColor = " duzzer blue 4 ";
                        else if (rgb == 7566080)
                            fColor = " deep sky blue 4 ";
                        else if (rgb == 11232768)
                            fColor = " deep sky blue 4 ";
                        else if (rgb == 11765099)
                            fColor = " light slik gray  ";
                        else if (rgb == 3552822)
                            fColor = " gray 21 ";
                        else if (rgb == 10921585)
                            fColor = " blue  ";
                        else if (rgb == 10181632)
                            fColor = " deep sky blue 4 "; //50

                        else if (rgb == 5714944)
                            fColor = " mid night blue ";
                        else if (rgb == 8355711)
                            fColor = " gray 50 ";
                        else if (rgb == 7619840)
                            fColor = " duzzer blue 4  ";
                        else if (rgb == 3022080)
                            fColor = " gray 10 ";
                        else if (rgb == 5065984)
                            fColor = " dark select gray ";
                        else if (rgb == 7488512)
                            fColor = " duzzer blue 4 ";
                        else if (rgb == 8279873)
                            fColor = " still blue 4  ";
                        else if (rgb == 1447446)
                            fColor = " gray 9 ";
                        else if (rgb == 7566151)
                            fColor = " aqua marin 4  ";
                        else if (rgb == 6831616)
                            fColor = " duzzer blue 4 "; // 60

                        else if (rgb == 192)
                            fColor = " red 3 ";
                        else if (rgb == 255)
                            fColor = " red ";
                        else if (rgb == 49407)
                            fColor = " gold  ";
                        else if (rgb == 65535)
                            fColor = " yellow ";
                        else if (rgb == 5296274)
                            fColor = " dark olive green 3 ";
                        else if (rgb == 5287936)
                            fColor = " sprint green 3 ";
                        else if (rgb == 15773696)
                            fColor = " deep sky blue 2  ";
                        else if (rgb == 12611584)
                            fColor = " duzzer blue 3 ";
                        else if (rgb == 6299648)
                            fColor = " mid night blue  ";
                        else if (rgb == 10498160)
                            fColor = " media market 4 "; // 70

                        ///////////

                        speaker.speak("Font is " + bold + paragraphAllignment);
                        speaker.speak(" paragraph level 1 ");
                        speaker.speak(fColor + " on black ");
                        speaker.speak(fN + " " + fS + " point");

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

        public void stopAll()
        {
            speaker.stop();

        }
    }
}
