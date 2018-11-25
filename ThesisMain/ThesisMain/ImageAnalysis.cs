using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.ComponentModel;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Collections;
using System.Data;
using System.Drawing.Drawing2D;




namespace ThesisMain
{
    class ImageAnalysis
    {

        //private Bitmap _screenShot;
        //protected static IntPtr newBMP;

        //private const int SRCCOPY = 13369376;
        //private const int SCREEN_X = 0;
        //private const int SCREEN_Y = 1;

        private static Bitmap bmpScreenshot;
        private static Graphics gfxScreenshot;
        
        //public ImageAnalysis()
        //{
        //    _screenShot = null;
        //}

        public ImageAnalysis()
        {
        }
      
        //public Bitmap GetScreen()
        //{
        //    int xLoc;
        //    int yLoc;
        //    IntPtr dsk;
        //    IntPtr mem;
        //    Bitmap currentView;
        //    Win32API Win32API = new Win32API();
            

        //    //get the handle of the desktop DC
        //    dsk = Win32API.GetDC(Win32API.GetDesktopWindow());

        //    //create memory DC
        //    mem = Win32API.CreateCompatibleDC(dsk);

        //    //get the X coordinates of the screen
        //    xLoc = Win32API.GetSystemMetrics(SCREEN_X);

        //    //get the Y coordinates of screen.
        //    yLoc = Win32API.GetSystemMetrics(SCREEN_Y);

        //    //create a compatible image the size of the desktop
        //    newBMP = Win32API.CreateCompatibleBitmap(dsk, xLoc, yLoc);

        //    //check against IntPtr (cant check IntPtr values against a null value)
        //    if (newBMP != IntPtr.Zero)
        //    {
        //        //select the image in memory
        //        IntPtr oldBmp = (IntPtr)Win32API.SelectObject(mem, newBMP);
        //        //copy the new bitmap into memory
        //        Win32API.BitBlt(mem, 0, 0, xLoc, yLoc, dsk, 0, 0, SRCCOPY);
        //        //select the old bitmap into memory
        //        Win32API.SelectObject(mem, oldBmp);
        //        //delete the memoryDC since we're through with it
        //        Win32API.DeleteDC(mem);
        //        //release dskTopDC to free up the resources
        //        Win32API.ReleaseDC(Win32API.GetDesktopWindow(), dsk);
        //        //create out BitMap
        //        currentView = Image.FromHbitmap(newBMP);
        //        //return the image
        //        return currentView;
        //    }
        //    else  //null value returned
        //    {
        //        MessageBox.Show("Give null");
        //        return null;
        //    }
        //}
        public Bitmap GetScreen()
        {
            bmpScreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format32bppArgb);
            // Create a graphics object from the bitmap

            gfxScreenshot = Graphics.FromImage(bmpScreenshot);

            // Take the screenshot from the upper left corner to the right bottom corner

            gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);

            //bmpScreenshot = ScaleByPercent( bmpScreenshot, 50 );

            //bmpScreenshot = new Bitmap(bmpScreenshot.Width, bmpScreenshot.Height, gfxScreenshot);
            //bmpScreenshot.Save("C:\\alamgir.bmp");
            return bmpScreenshot;
        }


        static Bitmap ScaleByPercent(Bitmap imgPhoto, int Percent)
        {
            float nPercent = ((float)Percent / 100);

            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;
            int sourceX = 0;
            int sourceY = 0;

            int destX = 0;
            int destY = 0;
            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap bmPhoto = new Bitmap(destWidth, destHeight,
                                     PixelFormat.Format24bppRgb);
            bmPhoto.SetResolution(imgPhoto.HorizontalResolution,
                                    imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.InterpolationMode = InterpolationMode.HighQualityBicubic;

            grPhoto.DrawImage(imgPhoto,
                new Rectangle(destX, destY, destWidth, destHeight),
                new Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                GraphicsUnit.Pixel);

            grPhoto.Dispose();
            return bmPhoto;
        }



        //public Bitmap GetScreenShot()
        //{
        //    _screenShot = new Bitmap(GetScreen());
        //    string ingName = "C:" + "\\" + "abul" + ".bmp";
        //    //save the image
        //    _screenShot.Save(ingName);
        //    return _screenShot;
        //}


        public String getIconText( )
        {
            //Bitmap bitmap = new Bitmap("C:\\what.bmp");
            try
            {
                Bitmap bitmap = GetScreen();

                if (bitmap == null) return null;

                int width = bitmap.Width;
                int height = bitmap.Height;
                bool start = false;
                int topX = 0, bottomX = 0, topY = 0, bottomY = 0;
                int i, j;


                //for (i = 0; i < width; i++)
                //{
                //    for (j = 0; j < height; j++)
                //    {
                //        Color pixelColor = bitmap.GetPixel(i, j);
                //        int r = pixelColor.R; // the Red component
                //        int b = pixelColor.B; // the Blue component
                //        int g = pixelColor.G;
                //        //Color newColor;
                //        //bitmap.SetPixel(i, j, newColor);

                //        if (!start && r == 49 && g == 106 && b == 197)
                //        {
                //            start = true;
                //            topX = i;
                //            topY = j;
                //        }
                //        else if (r == 49 && g == 106 && b == 197)
                //        {
                //            bottomX = i;
                //            bottomY = j;
                //        }

                //    }
                //}

                //for (i = 0; i < width; i++)
                //{
                //    for (j = 0; j < height; j++)
                //    {
                //        Color pixelColor = bitmap.GetPixel(i, j);

                //        Color newColor;

                //        int r = pixelColor.R; // the Red component
                //        int b = pixelColor.B; // the Blue component
                //        int g = pixelColor.G; //the green component

                //        if (r == 49 && g == 106 && b == 197)
                //        {
                //            newColor = Color.FromArgb(255, 255, 255);
                //            bitmap.SetPixel(i, j, newColor);
                //            continue;
                //        }
                //        if (r == 255 && g == 255 && b == 255 && i >= topX && i <= bottomX && j >= topY && j <= bottomY)
                //        {
                //            newColor = Color.FromArgb(0, 0, 0);
                //            bitmap.SetPixel(i, j, newColor);
                //            continue;
                //        }

                //        newColor = Color.FromArgb(255, 255, 255);
                //        bitmap.SetPixel(i, j, newColor);

                //    }
                //}


                /*
                 * Testing
                 */

                for (i = 0; i < width; i++)
                {
                    for (j = 0; j < height; j++)
                    {
                        Color pixelColor = bitmap.GetPixel(i, j);
                        int r = pixelColor.R; // the Red component
                        int b = pixelColor.B; // the Blue component
                        int g = pixelColor.G;
                        //Color newColor;
                        //bitmap.SetPixel(i, j, newColor);

                        if (!start && r == 49 && g == 106 && b == 197)
                        {
                            start = true;
                            topX = i;
                            topY = j;
                            break;
                        }
                        //else if (r == 49 && g == 106 && b == 197)
                        //{
                        //    bottomX = i;
                        //    bottomY = j;
                        //}

                    }
                    if (start)
                    {
                        break;
                    }
                }

                for (i = topX; i < width; i++)
                {
                    Color pixelColor = bitmap.GetPixel(i, topY);
                    int r = pixelColor.R; // the Red component
                    int b = pixelColor.B; // the Blue component
                    int g = pixelColor.G;
                    //if (r == 255 && g == 255 && b == 255) continue;
                    if (r == 49 && g == 106 && b == 197)
                    {
                        bottomX = i;
                    }
                    //if (!((r == 255 && g == 255 && b == 255) || (r == 49 && g == 106 && b == 197))) break;
                }

                //j = topY;

                for (j = topY; j < height; j++)
                {
                    Color pixelColor = bitmap.GetPixel(topX, j);
                    int r = pixelColor.R; // the Red component
                    int b = pixelColor.B; // the Blue component
                    int g = pixelColor.G;
                    //if (!((r == 255 && g == 255 && b == 255) || (r == 49 && g == 106 && b == 197))) break;
                    if (r == 49 && g == 106 && b == 197)
                    {
                        bottomY = j;
                    }
                    //if (!(r == 49 && g == 106 && b == 197)) break;
                }

                //bottomY = i;
                //bottomX = j;

                width = bottomX - topX;
                height = bottomY - topY;
                //MessageBox.Show(topX + " " + topY + " " + (bottomX - topX) + " " + (bottomY - topY));


                Bitmap newBitmap = new Bitmap(width + 1, height + 1);
                //return topX + " " + topY + " " + bottomX + " " + bottomY;


                for (i = topX; i <= bottomX; i++)
                {
                    for (j = topY; j <= bottomY; j++)
                    {
                        Color pixelColor = bitmap.GetPixel(i, j);

                        int r = pixelColor.R; // the Red component
                        int b = pixelColor.B; // the Blue component
                        int g = pixelColor.G;//the green component

                        Color newColor;

                        //if (r == 255 && g == 255 && b == 255)
                        //{
                        //    newColor = Color.FromArgb(0, 0, 0);
                        //    newBitmap.SetPixel(i - topX, j - topY, newColor);
                        //}
                        //else
                        //{
                        //    newColor = Color.FromArgb(255, 255, 255);
                        //    newBitmap.SetPixel(i - topX, j - topY, newColor);
                        //}

                        newColor = Color.FromArgb(r, g, b);

                        newBitmap.SetPixel(i - topX, j - topY, newColor);
                    }
                }

                //string imgName = "C:" + "\\" + "alamgir" + ".bmp";
                ////save the image
                //newBitmap.Save(imgName);
                //AspriseOcr aspriseOcr = new AspriseOcr();
                //return aspriseOcr.getTextFromImage(imgName, 0, 0, width, height);

                if (topX == bottomX || topY == bottomY) return null;
                newBitmap = ScaleByPercent(newBitmap, 200);
                
                //newBitmap.Save("C:/alamgir.bmp");

                return getTextFromIcon(newBitmap);
            }
            catch( Exception )
            {
                MessageBox.Show("Exception imageanalysis");
                return null;
            }

            /*
             * 
             */


            //string ingName = "C:" + "\\" + "alamgir" + ".jpg";
            ////save the image
            //bitmap.Save(ingName, ImageFormat.Jpeg);
            //AspriseOcr aspriseOcr = new AspriseOcr();
            //return aspriseOcr.getTextFromImage(ingName, 0, 0, 800, 600);
        }


        private string getTextFromIcon(Bitmap bitmap)
        {
            string str = "";
            try
            {
                using (Bitmap bmp = bitmap)
                {
                    tessnet2.Tesseract tessocr = new tessnet2.Tesseract();
                    String tessDataPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\tessdata";
                    tessocr.Init(tessDataPath, "eng", false);
                    List<tessnet2.Word> result = tessocr.DoOCR(bmp, Rectangle.Empty);

                    foreach (tessnet2.Word word in result)
                    {
                        str += word.Text;
                    }


                }
            }
            catch( Exception )
            {
                
            }
            return str;
        }


    }
}

