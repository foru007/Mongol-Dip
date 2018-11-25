using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace StartLoader
{
    public partial class LoaderForm : Form
    {
        private Rectangle rect;
        private int x, y, width, height;
        private double fraction, timeRemain;

        public LoaderForm()
        {
            InitializeComponent();
            init();
        }

        public void init()
        {
            this.Opacity = 0;
            x = loadingPanel.ClientRectangle.X;
            y = loadingPanel.ClientRectangle.Y;
            height = loadingPanel.ClientRectangle.Height;
            width = 1;
            fraction = (double).08;
            timeRemain = (double)5.0;

            return;
        }


        private void loadingPanel_Paint(object sender, PaintEventArgs e)
        {
            if (e.ClipRectangle.Width > 0)
            {
                rect = new Rectangle(x, y, width, height);
                LinearGradientBrush brBackground = new LinearGradientBrush(rect, Color.Red, Color.Black, LinearGradientMode.Horizontal);
                e.Graphics.FillRectangle(brBackground, rect);
            }
        }

        private void opacityTimer_Tick(object sender, EventArgs e)
        {
            if (this.Opacity < 1 && timeRemain > 3 )
            {
                this.Opacity += .1;
                
            }
            if( timeRemain <= 3 )
            {
                this.Opacity -=.05;
            }
        }

        private void loadingPanelTimer_Tick(object sender, EventArgs e)
        {
            if (width < loadingPanel.ClientRectangle.Width)
            {
                width = (int)(Math.Floor(loadingPanel.ClientRectangle.Width * fraction));

                Random random = new Random();
                double val = random.NextDouble();

                if (val < .15)
                {
                    fraction += val;
                }
                
            }
            loadingPanel.Invalidate();
        }

        private void timeRemainingTimer_Tick(object sender, EventArgs e)
        {
            timeRemain = Math.Floor(10 * (1 - fraction));
            
            if( timeRemain > 0 )
            {
                timeRemaining_label.Text = timeRemain + " seconds Remaining";
            }
            if( timeRemain < 2 )
            {
                this.Dispose();
            }
        }

    }
}
