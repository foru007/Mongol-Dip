using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace SplashScreen
{
    public class ReadyToStart
    {

        public ReadyToStart()
        {
        }

        public void start()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            SplashScreen.ShowSplashScreen();
            Application.DoEvents();
            SplashScreen.SetStatus("Loading images");
            System.Threading.Thread.Sleep(600);

            SplashScreen.SetStatus("Loading module 1");
            System.Threading.Thread.Sleep(240);
            SplashScreen.SetStatus("Loading module 2");
            System.Threading.Thread.Sleep(900);
            SplashScreen.SetStatus("Loading module 3");
            System.Threading.Thread.Sleep(240);
            SplashScreen.SetStatus("Loading module 4");
            System.Threading.Thread.Sleep(90);
            SplashScreen.SetStatus("Loading module 5");
            System.Threading.Thread.Sleep(400);
            SplashScreen.SetStatus("Loading module 6");
            System.Threading.Thread.Sleep(100);

            SplashScreen.SetStatus("Loading sounds");
            System.Threading.Thread.Sleep(900);
            SplashScreen.SetStatus("Loading module 1");
            System.Threading.Thread.Sleep(500);
            SplashScreen.SetStatus("Loading module 2", false);
            System.Threading.Thread.Sleep(500);
            SplashScreen.SetStatus("Loading module 3", false);
            System.Threading.Thread.Sleep(400);
            SplashScreen.SetStatus("Loading module 4", false);
            System.Threading.Thread.Sleep(1000);
            SplashScreen.SetStatus("Loading module 5", false);
            System.Threading.Thread.Sleep(1000);

            SplashScreen.SetStatus("Loading dll files");
            System.Threading.Thread.Sleep(700);
            SplashScreen.SetStatus("Loading objects");
            System.Threading.Thread.Sleep(800);
            SplashScreen.SetStatus("Loading figures");
            System.Threading.Thread.Sleep(400);
            SplashScreen.SetStatus("Loading modules 6");
            System.Threading.Thread.Sleep(50);



            SplashScreen.SetStatus("Loading module 1");
            System.Threading.Thread.Sleep(20);
            SplashScreen.SetStatus("Loading module 2");
            System.Threading.Thread.Sleep(450);
            SplashScreen.SetStatus("Loading module 3");
            System.Threading.Thread.Sleep(240);
            SplashScreen.SetStatus("Loading module 4");
            System.Threading.Thread.Sleep(90);
            SplashScreen.CloseForm();
        }

    }
}
