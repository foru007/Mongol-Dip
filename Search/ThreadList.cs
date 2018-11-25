using System;
using System.Collections;
using System.Threading;
using System.Windows.Forms;
using SpeechBuilder;



namespace Search
{
	/// <summary>
	/// Thread states
	/// </summary>
	public enum SearchThreadState {ready, running, cancelled};

	/// <summary>
	/// 
	/// </summary>
	public class SearchThread
	{
		public string name;
		public Thread thrd;
		public SearchThreadState state;
		public string searchdir;
	}


	/// <summary>
	/// Summary description for ThreadList.
	/// </summary>
	public class ThreadList
	{
		private ArrayList thrdList;
        private SpeechControl speaker=new SpeechControl();
		
		public ThreadList()
		{
			thrdList = new ArrayList();
		}

		public int AddItem(SearchThread thrd ) 
		{
			return thrdList.Add(thrd);
		}

		public SearchThread Item(int index)
		{
			return (SearchThread) thrdList[index];
		}

		public SearchThread Item(string name)
		{
			SearchThread st = null;
			SearchThread t;

			for (int i = 0; i <thrdList.Count ;i++)
			{
				t =(SearchThread)thrdList[i];
				if (t.name == name)
				{
					st = t;
					break;
				}			
			}
            //speaker.speak("Search Completed");
            //MessageBox.Show("pp");
			return st;
            
		}

		public Boolean RemoveItem(int index)
		{
			Thread thrd;

			Boolean bRemoved = false;
			try
			{
				thrd = ((SearchThread)thrdList[index]).thrd;
				// if thread is still alive
				// (it should be already stopped)
				// force an Abort.
				try
				{
					if (thrd.IsAlive)
					{
						thrd.Abort();
					}
				}
				catch
				{
				}
				thrdList.RemoveAt(index);
				bRemoved = true;
			}
			catch
			{
			}
			return bRemoved;
		}

		public int ItemCount()
		{
			return thrdList.Count;
		}

		public int[] ItemState()
		{
			int[] result  = {0,0,0};

			for (int i = 0; i <thrdList.Count ;i++)
			{
				result[(int)(((SearchThread)thrdList[i]).state)]++;
			}
			return result;
		}

	}
}
