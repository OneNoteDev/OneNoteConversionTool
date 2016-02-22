using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HtmlAgilityPack;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNoteConversionTool.FormatReaders;

namespace OneNoteConversionToolUnitTest.FormatReaders
{
	[TestClass]
	public class EpubReaderUnitTest
	{
		private const string Title = "EpubSample";
		private static readonly List<string> PageTitles = new List<string>() { "Epub Sample Cover", "Epub Sample TOC", "Epub Sample Page" };
		private static readonly List<int> PageLevels = new List<int>() { 1, 1, 1};

		private const string TestEpubFile = "..\\..\\Resources\\EpubSample.epub";
		private static readonly string TestEpubPath = Path.Combine(Environment.CurrentDirectory, TestEpubFile);

		private static EpubReader _mEpubReader;

		/// <summary>
		/// Initializer method for the Unit Test
		/// </summary>
		/// <param name="testContext"></param>
		[ClassInitialize()]
		public static void MyClassInitialize(TestContext testContext)
		{
			if (!Directory.Exists(Utility.RootFolder))
			{
				Directory.CreateDirectory(Utility.RootFolder);
			}

			_mEpubReader = new EpubReader(TestEpubPath, Utility.RootFolder);
		}

		/// <summary>
		/// Cleanup: Delete temporary folders
		/// </summary>
		[ClassCleanup()]
		public static void MyClassCleanup()
		{
			//delete temporary folders
			if (Directory.Exists(Utility.RootFolder))
			{
				Utility.DeleteDirectory(Utility.RootFolder);
			}
		}

		/// <summary>
		/// Validate Obtaining the title of the file
		/// </summary>
		[TestMethod]
		public void ValidateGetTitle()
		{
			string title = _mEpubReader.GetTitle();
			Console.WriteLine(title);
			Assert.AreEqual(Title, title);
		}

		/// <summary>
		/// Validate obtaining the titles of the pages
		/// </summary>
		[TestMethod]
		public void ValidateGetPageTitles()
		{
			List<string> titles = _mEpubReader.GetPageTitles();
			CollectionAssert.AreEqual(PageTitles, titles);
		}

		/// <summary>
		/// Validate obtaining the pages html as HtmlDocument
		/// </summary>
		[TestMethod]
		public void ValidateGetPagesAsHtmlDocument()
		{
			List<HtmlDocument> pagesHtml = _mEpubReader.GetPagesAsHtmlDocuments();
			Assert.AreEqual(pagesHtml.Count, 3);
		}

		/// <summary>
		/// Validate obtaining the levels of all pages
		/// </summary>
		[TestMethod]
		public void ValidateGetPagesLevel()
		{
			Dictionary<string, int> pageLevels = _mEpubReader.GetPagesLevel();
			List<int> pageLevelsList = (from int level in pageLevels.Values select level).ToList();
			CollectionAssert.AreEqual(PageLevels, pageLevelsList);
		}

		/// <summary>
		/// Validate using non-existing input file
		/// </summary>
		[TestMethod]
		[ExpectedException(typeof(FileNotFoundException))]
		public void ValidateNonExistFile()
		{
			var epubReader = new EpubReader(Utility.NonExistentInputFile, Utility.RootFolder);
		}
	}
}
