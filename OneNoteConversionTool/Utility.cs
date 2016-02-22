using System;
using System.IO;
using System.Text;
using System.Web;

namespace OneNoteConversionTool
{
	/// <summary>
	/// Class that contains a bunch of helper methods for ConversionTool
	/// </summary>
	public class Utility
	{
		#region Available Extenssions
		public static string[] WordSupportedExtenssions = { ".doc", ".docx", ".dot", ".dotx", ".docm", ".dotm", ".odt" };
		public static string[] PowerPointSupportedExtenssions = { ".ppt", ".pptx", ".pot", ".potx", ".pptm", ".potm", ".odp" };
		public static string[] PdfSupportedExtenssions = { ".pdf" };
		public static string[] EpubSupportedExtenssions = { ".epub" };
		public static string[] InDesignSupportedExtenssions = { ".indd" };
		#endregion

		#region Files and Folders Utility Methods
		/// <summary>
		/// Helper method to create a directory
		/// </summary>
		/// <param name="pathDir"></param>
		/// <returns></returns>
		public static string CreateDirectory(string pathDir)
		{
			string retPath = pathDir;
			int i = 1;
			// ensure that the name of the directory is unique
			while (Directory.Exists(retPath))
			{
				retPath = String.Format("{0} ({1})", retPath, i);
				i++;
			}
			// create the directory
			Directory.CreateDirectory(retPath);

			return retPath;
		}

		/// <summary>
		/// Helper method to get a path of a new folder in a given directory
		/// Note: this method doesn't create the folder
		/// </summary>
		/// <param name="dir">Directory where the folder will be created</param>
		/// <param name="folderName">Name of the folder</param>
		/// <returns>Full path of the new folder</returns>
		public static string NewFolderPath(string dir, string folderName)
		{
			if (dir == null || folderName == null)
			{
				Console.WriteLine("Error in NewFolderPath: One of the parameters is null");
				return String.Empty;
			}
			if (!Directory.Exists(dir))
			{
				Console.WriteLine("Error in NewFolderPath: {0} doesn't exist", dir);
				return String.Empty;
			}

			// make the folder name "New  Folder" if it was empty
			folderName = folderName != String.Empty ? folderName : "New Folder";

			string retPath = Path.Combine(dir, folderName);
			int i = 1;
			while (Directory.Exists(retPath))
			{
				retPath = String.Format("{0} {1}", Path.Combine(dir, folderName), i);
				i++;
			}
			return retPath;
		}

		/// <summary>
		/// Helper method to get a path of a new file in a given directory
		/// Note: this method doesn't create the file
		/// </summary>
		/// <param name="dir">Directory where the file will be created</param>
		/// <param name="fileName">Name of the file</param>
		/// <param name="extension">Extension of the file</param>
		/// <returns>full path of the new file</returns>
		public static string NewFilePath(string dir, string fileName, string extension)
		{
			if (dir == null || fileName == null || extension == null)
			{
				Console.WriteLine("Error in NewFilePath: One of the parameters is null");
				return String.Empty;
			}
			if (!Directory.Exists(dir))
			{
				Console.WriteLine("Error in NewFilePath: {0} doesn't exist", dir);
				return String.Empty;
			}

			// make the file name "New File" if it was empty
			fileName = fileName != String.Empty ? fileName : "New File";

			string fullFileName = Path.Combine(dir, fileName);
			string retPath = Path.ChangeExtension(fullFileName, extension);
			int i = 1;
			while (File.Exists(retPath))
			{
				retPath = String.Format("{0} {1}{2}", fullFileName, i, extension);
				i++;
			}

			return retPath;
		}

		/// <summary>
		/// Helper method to delete directory and its contents recursively
		/// </summary>
		/// <param name="targetDir"></param>
		/// <returns></returns>
		public static bool DeleteDirectory(string targetDir)
		{
			bool retVal = false;

			try
			{
				Directory.Delete(targetDir, true);
				retVal = true;
			}
			catch
			{
				Console.WriteLine("Failed to delete directory: {0}", targetDir);
			}

			return retVal;
		}
		#endregion

		#region Html/Xml Utility Methods
		/// <summary>
		/// Changes the escape sequence characters with the equivalant characters.
		/// </summary>
		/// <param name="docContent">Document content.</param>
		/// <returns>The converted string is being returned.</returns>
		public static string UnescapeXml(string docContent)
		{
			string unxml = docContent;
			try
			{
				if (!string.IsNullOrEmpty(unxml))
				{
					// replace entities with literal values
					StringWriter sw = new StringWriter();
					HttpUtility.HtmlDecode(unxml, sw);
					unxml = sw.ToString();
					unxml = Encoding.UTF8.GetString(Encoding.Default.GetBytes(unxml));
				}
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in UnescapeXml", e);
			}
			return unxml;
		}
		#endregion
	}
}