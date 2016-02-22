using System;
using System.IO;

namespace OneNoteConversionToolUnitTest
{
	/// <summary>
	/// Class that contains a bunch of helper methods for UnitTesting
	/// </summary>
	public class Utility
	{
		public const string NS = "http://schemas.microsoft.com/office/onenote/2013/onenote";
		public const string TempFolder = @"C:\onetest\temporary";
		public const string RootFolder = @"C:\onetest";
		public const string NonExistentInputFile = @"C:\thisFileDontExist";
		public const string NonExistentOutputPath = @"C:\invalidOutputDir";


		#region Helper methods
		/// <summary>
		/// Helper method to delete entire folder and files
		/// </summary>
		/// <param name="targetDir"></param>
		public static bool DeleteDirectory(string targetDir)
		{
			try
			{
				Directory.Delete(targetDir, true);
			}
			catch
			{
				Console.WriteLine("Failed to delete directory: " + targetDir);
				return false;
			}

			return true;
		}
		#endregion
	}
}
