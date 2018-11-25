using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;




namespace PaglaPlayer
{
	public class Player 
	{
		public bool SongEnded=true;
		
		private System.Windows.Forms.Timer CheckSong;
		private System.ComponentModel.IContainer play_components;
	
		ArrayList SongsInPlaylist = new ArrayList();
		private int Index = 0;
		public AxWMPLib.AxWindowsMediaPlayer MediaPlayer;

     
		public Player(String[] Songs, AxWMPLib.AxWindowsMediaPlayer Player)
		{
			AddSongs(Songs);

			MediaPlayer = Player;
			
			this.play_components = new System.ComponentModel.Container();
			this.CheckSong = new System.Windows.Forms.Timer(this.play_components);
			this.CheckSong.Tick += new System.EventHandler(this.CheckSong_Tick);

			MediaPlayer.PlayStateChange +=new AxWMPLib._WMPOCXEvents_PlayStateChangeEventHandler(MediaPlayer_PlayStateChange);
          

			Play();

		}

       


		public void AddSongs(string[] Songs)
		{
			for( int i=0; i< Songs.Length ; i++)
			{
				AddSong(Songs[i]);
			}
		}
		public void AddSong(string Song)
		{
			SongsInPlaylist.Add(Song);
		}
		public void DeleteSong(string Song)
		{
			if(Song == SongsInPlaylist[Index].ToString())
			{
				MediaPlayer.Ctlcontrols.stop();
				Index--;
			}
			SongsInPlaylist.Remove(Song);
			MediaPlayer.Ctlcontrols.play();
		}
		public void DeletePlaylist()
		{
			MediaPlayer.Ctlcontrols.stop();
			SongsInPlaylist.Clear();
			Index = 0;
		}
		public int Volume
		{
			set { MediaPlayer.settings.volume = value; }
			get { return MediaPlayer.settings.volume; }
		}

		public void Play()
		{
			if(SongsInPlaylist[Index] != null)
			{
				MediaPlayer.URL = SongsInPlaylist[Index].ToString();
			}
		}
		public void Play(int Slot)
		{
			if(SongsInPlaylist[Slot-1] != null)
				MediaPlayer.URL = SongsInPlaylist[Slot-1].ToString();
		}
		public void Play(string name)
		{
			int slot = SongsInPlaylist.BinarySearch(name,null);
			if(slot >= 0 && slot < SongsInPlaylist.Count)
				MediaPlayer.URL = SongsInPlaylist[slot].ToString();
		}

		public void Pause()
		{
			MediaPlayer.Ctlcontrols.pause();
		}
		public void Stop()
		{
			MediaPlayer.Ctlcontrols.stop();
		}
		public void NextSong()
		{
			if(Index != SongsInPlaylist.Count - 1)
			{
				Index++;
				MediaPlayer.Ctlcontrols.stop();
				MediaPlayer.URL = SongsInPlaylist[Index].ToString();
				MediaPlayer.Ctlcontrols.play();
			}
			else
			{
				Index = 0;
				MediaPlayer.Ctlcontrols.stop();
				MediaPlayer.URL = SongsInPlaylist[0].ToString();
				MediaPlayer.Ctlcontrols.play();
			}
		}
		public void PrevSong()
		{
			if(Index != 0)
			{
				Index--;
				MediaPlayer.Ctlcontrols.stop();
				MediaPlayer.URL = SongsInPlaylist[Index].ToString();
				MediaPlayer.Ctlcontrols.play();
			}
			else
			{
				Index = SongsInPlaylist.Count - 1;
				MediaPlayer.Ctlcontrols.stop();
				MediaPlayer.URL = SongsInPlaylist[Index].ToString();
				MediaPlayer.Ctlcontrols.play();
			}
		}
		private void CheckSong_Tick(object sender, System.EventArgs e)
		{
			if(SongEnded)
			{
				NextSong();
				SongEnded = false;
				CheckSong.Stop();
			}
		}

		public void MediaPlayer_PlayStateChange(object sender, AxWMPLib._WMPOCXEvents_PlayStateChangeEvent e)
		{
			switch(MediaPlayer.playState)
			{
				case WMPLib.WMPPlayState.wmppsMediaEnded:
					SongEnded = true;
					CheckSong.Start();
					break;
				default:
					break;
			}
		}
       
        
	}

}

