using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace PaglaPlayer
{
    public partial class PaglaPlayerPane : Form
    {
        Player player;
        bool playState = true;

       
        
        public PaglaPlayerPane( )
        {
            InitializeComponent();
            
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            //song[0] = @"F:\Chumky\se j keno alo na.mp3";
            //song[1] = @"F:\Chumky\ami tomar.e  prem.o  vekhari.mp3";
            //if (song != null)
            //{
            //    //player = new Player(song, playerInterface);
            //}
        }

        public void startPlaying( String[] song )
        {
            if (song != null)
            {
                player = new Player(song, playerInterface);
            }
        }
       
            

        private void axWindowsMediaPlayer1_KeyDownEvent(object sender, AxWMPLib._WMPOCXEvents_KeyDownEvent e)
        {
            /*
             * + sign to increase vol
             * - sign to decrease vol
             * space to pause and start
             * n for next
             * and p for prev song
             */
          
            if(e.nKeyCode.ToString().Equals("107"))
            {
                player.Volume++;
            }
            else if (e.nKeyCode.ToString().Equals("109"))
            {
                player.Volume--;
            }
            else if( e.nKeyCode.ToString().Equals("32"))
            {
                if (playState)
                {
                    playState = false;
                    player.Pause();
                }
                else
                {
                    playState = true;
                    player.Play();
                }
            }
            else if (e.nKeyCode.ToString().Equals("78"))
            {
                player.NextSong();
            }
            else if( e.nKeyCode.ToString().Equals("80"))
            {
                player.PrevSong();
            }
        
        
        }

    }
}
