using System;
using System.Collections.Specialized;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading;


namespace Search
{
	/// <summary>
	/// Summary description for FileSearch.
	/// </summary>
	public class FileSearch
	{
		private String Dir;
        String SearchPattern, SearchForText;
		bool CaseSensitive;
        Form1 mainForm;
		string thrdName;


        /// <summary>
        /// Delegate for UpdateThreadStatus function
        /// </summary>
        public delegate void UpdateThreadStatusDelegate(String thrdName, SearchThreadState sts);

		/// <summary>
		/// initialize the FileSearch class
		/// </summary>
		/// <param name="Dir"></param>
		/// <param name="SearchPattern"></param>
		/// <param name="SearchForText"></param>
		/// <param name="CaseSensitive"></param>
		/// <param name="mainForm"></param>
		/// <param name="thrdName"></param>
		public FileSearch(String Dir, String SearchPattern, String SearchForText, bool CaseSensitive, Form1 mainForm, string thrdName)
		{
			this.Dir = Dir;
			this.SearchPattern = SearchPattern;
			this.SearchForText = SearchForText;
			this.CaseSensitive = CaseSensitive;
			this.mainForm = mainForm;
			this.thrdName = thrdName;
            
		}

		/// <summary>
		/// Start searching
		/// </summary>
		public void SearchDir()
		{
			// get the UpdateThreadStatus delegate

            Form1.UpdateThreadStatusDelegate UTSDelegate = new Form1.UpdateThreadStatusDelegate(this.mainForm.UpdateThreadStatus);

		//	MainForm.UpdateThreadStatusDelegate UTSDelegate = new MainForm.UpdateThreadStatusDelegate(this.mainForm.UpdateThreadStatus);
			
			// and Invoke it
          

		//	mainForm.Invoke(UTSDelegate, new object[] { this.thrdName, SearchThreadState.running });


			try
			{
				Regex FileExtensionDelim = new Regex(",");
				String []FileExtensions = FileExtensionDelim.Split(SearchPattern);

//				if (FileExtensions.Length == 0)
//				{
//					FileExtensions = new String[1];
//					FileExtensions[0] = "";
//				}

				for (int i = 0; i < FileExtensions.Length; i++)
				{
					FileExtensions[i] = FileExtensions[i].Trim();
				}

				GetFiles(Dir, Dir, FileExtensions, SearchForText, CaseSensitive, mainForm);


				// update thread state
				mainForm.Invoke(UTSDelegate, new object[] { this.thrdName, SearchThreadState.ready });
                //MessageBox.Show("end");

			}
			catch (Exception)
			{
				// update thread state
		//		mainForm.Invoke(UTSDelegate, new object[] { this.thrdName, SearchThreadState.cancelled });
			}
		}

		/// <summary>
		/// Get files
		/// </summary>
		/// <param name="RootDir"></param>
		/// <param name="Dir"></param>
		/// <param name="FileExtensions"></param>
		/// <param name="SearchForText"></param>
		/// <param name="CaseSensitive"></param>
		/// <param name="mainForm"></param>
        private static void GetFiles(string RootDir, string Dir, String[] FileExtensions, String SearchForText, bool CaseSensitive, Form1 mainForm)
		{
			// get the AddListBoxItem delegate
        //    Thesis.Form1

            Form1.AddListBoxItemDelegate ALBIDelegate = new Form1.AddListBoxItemDelegate(mainForm.AddListBoxItem);

		//	MainForm.AddListBoxItemDelegate ALBIDelegate = new MainForm.AddListBoxItemDelegate(mainForm.AddListBoxItem);
         
			try
			{
				foreach (string FileExtension in FileExtensions)
				{
					foreach (string File in Directory.GetFiles(Dir, FileExtension))
					{
						if (FileContainsText(File, SearchForText, CaseSensitive))
						{
							mainForm.Invoke(ALBIDelegate, new Object[] { File });
						}
					}
				}
			}
			catch (Exception)
			{
			}

			// Recursively add all the files in the
			// current directory's subdirectories.
			try
			{
				foreach (string D in Directory.GetDirectories(Dir))
				{
					GetFiles(RootDir, D, FileExtensions, SearchForText, CaseSensitive, mainForm);

				}
			}
			catch (Exception)
			{
			}
		}


		/// <summary>
		/// get the file content
		/// </summary>
		/// <param name="FileName"></param>
		/// <param name="Error"></param>
		/// <returns></returns>
		public static String GetFileContent(String FileName, out bool Error)
		{
			String TextContent = "";

			FileInfo fInfo = new FileInfo(FileName);

			FileStream   fStream = null;
			StreamReader sReader = null;

			Error = false;

			try
			{
				fStream = fInfo.OpenRead();
				sReader = new StreamReader(fStream);

				TextContent = sReader.ReadToEnd();

			
			}
			catch (System.IO.IOException)
			{
				Error = true;
			}
			finally
			{
				if (fStream != null)
				{
					fStream.Close();
				}

				if (sReader != null)
				{
					sReader.Close();
				}

			}

			return TextContent;
		}

		/// <summary>
		/// if SearchForText length i > 0
		/// search it inside the file content
		/// </summary>
		/// <param name="FileName"></param>
		/// <param name="SearchForText"></param>
		/// <param name="CaseSensitive"></param>
		/// <returns></returns>
		public static bool FileContainsText(String FileName, String SearchForText, bool CaseSensitive)
		{
			bool Result = (SearchForText.Length == 0);

			if (!Result)
			{
				bool Error;
				String TextContent = GetFileContent(FileName, out Error);

				if (!Error)
				{
					if (!CaseSensitive)
					{
						TextContent = TextContent.ToUpper();
						SearchForText = SearchForText.ToUpper();
					}

					Result = (TextContent.IndexOf(SearchForText) != -1);
				}			
			}

			return Result;
		}

	}
}
