using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNoteConversionTool.FormatConversion;
using OneNoteConversionTool.OutputGenerator;
using _WordApplication = Microsoft.Office.Interop.Word._Application;
using _WordDocument = Microsoft.Office.Interop.Word._Document;

namespace OneNoteConversionToolUnitTest.FormatConversion
{
	/// <summary>
	/// Unit test for GenericFormatConverter
	/// </summary>
	[TestClass]
	public class GenericFormatConverterUnitTest
	{
		//Page titles of the generated word document
		private static readonly List<string> DocPageTitles = new List<string>() { "Table of Contents", "this is the first page", "this is the second page", "this is the third page" };
		private static readonly List<string> PptPageTitles = new List<string>() { "Table of Contents", "MainSection", "PresentationTitle", "FirstSection", "Some New Section", "Third Slide", "SecondSection", "PageTitle" };
		private static readonly List<string> EpubPageTitles = new List<string>() { "Epub Sample Cover", "Epub Sample TOC", "Epub Sample Page" };

		private const string TestDocName = "GenericTest.docx";
		private const string TestPptName = "..\\..\\Resources\\SectionSample.pptx";
		private const string TestEpubFile = "..\\..\\Resources\\EpubSample.epub";
		private const string NotebookName = "Generic";
		private static readonly string TestDocPath = Path.Combine(Utility.TempFolder, TestDocName);
		private static readonly string TestPptPath = Path.Combine(Environment.CurrentDirectory, TestPptName);
		private static readonly string TestEpubPath = Path.Combine(Environment.CurrentDirectory, TestEpubFile);
		private static string _mNotebookId = String.Empty;
		private static XNamespace _mXmlNs;
		private static OneNoteGenerator _mOnGenerator;
		
		/// <summary>
		/// Create temporary folders and initialize OneNote Generator
		/// </summary>
		[ClassInitialize()]
		public static void MyClassInitialize(TestContext testContext)
		{
			if (!Directory.Exists(Utility.RootFolder))
			{
				Directory.CreateDirectory(Utility.RootFolder);
			}
			if (!Directory.Exists(Utility.TempFolder))
			{
				Directory.CreateDirectory(Utility.TempFolder);
			}
			_mXmlNs = Utility.NS;
			_mOnGenerator = new OneNoteGenerator(Utility.RootFolder);
			//Get Id of the test notebook so we chould retrieve generated content
			//KindercareFormatConverter will create notebookName as Kindercare
			_mNotebookId = _mOnGenerator.CreateNotebook(NotebookName);

			var word = new Application();
			var doc = word.Application.Documents.Add();

			//add pages to doc
			for (int i = 1; i < DocPageTitles.Count; i++)
			{
				doc.Content.Text += DocPageTitles[i];
				doc.Words.Last.InsertBreak(WdBreakType.wdPageBreak);
			}

			var filePath = TestDocPath as object;
            doc.SaveAs(ref filePath);
            ((_WordDocument)doc).Close();
            ((_WordApplication)word).Quit();
		}
		
		/// <summary>
		/// Delete temporary folders
		/// </summary>
		[ClassCleanup()]
		public static void MyClassCleanup()
		{
			//delete temporary folders
			if (Directory.Exists(Utility.RootFolder))
			{
				Utility.DeleteDirectory(Utility.RootFolder);
			}
			if (Directory.Exists(Utility.NonExistentOutputPath))
			{
				Utility.DeleteDirectory(Utility.NonExistentOutputPath);
			}
		}

		/// <summary>
		/// create a simple docx file
		/// and validate its conversion
		/// </summary>
		[TestMethod]
		public void ValidateWordConversion()
		{
			var converter = new GenericFormatConverter();
			converter.ConvertWordToOneNote(TestDocPath, Utility.RootFolder);

			//retrieve xml from generated notebook
			var xmlDoc = _mOnGenerator.GetPageScopeHierarchy(_mNotebookId);
			Assert.IsNotNull(xmlDoc);

			// retrieve the section for the ppt conversion
			string sectionName = Path.GetFileNameWithoutExtension(TestDocPath);
			XDocument xDoc = XDocument.Parse(xmlDoc);
			XElement xSection = xDoc.Descendants(_mXmlNs + "Section").FirstOrDefault(x => x.Attribute("name").Value.Equals(sectionName));
			Assert.IsNotNull(xSection);

			//get the xml of each pages
			var extractedPageTitles = xSection.Descendants(_mXmlNs + "Page").Select(x => x.Attribute("name").Value).ToList();
			CollectionAssert.AreEqual(DocPageTitles, extractedPageTitles);
		}

		/// <summary>
		/// validate the conversion of the ppt file in the resources folder (SectionSample.pptx)
		/// </summary>
		[TestMethod]
		public void ValidatePowerPointConversion()
		{
			var converter = new GenericFormatConverter();
			converter.ConvertPowerPointToOneNote(TestPptPath, Utility.RootFolder);

			// retrieve xml from generated notebook
			var xmlDoc = _mOnGenerator.GetPageScopeHierarchy(_mNotebookId);
			Assert.IsNotNull(xmlDoc);

			// retrieve the section for the ppt conversion
			string sectionName = Path.GetFileNameWithoutExtension(TestPptPath);
			XDocument xDoc = XDocument.Parse(xmlDoc);
			XElement xSection = xDoc.Descendants(_mXmlNs + "Section").FirstOrDefault(x => x.Attribute("name").Value.Equals(sectionName));
			Assert.IsNotNull(xSection);

			//get the xml of each pages
			var extractedPageTitles = xSection.Descendants(_mXmlNs + "Page").Select(x => x.Attribute("name").Value).ToList();
			CollectionAssert.AreEqual(PptPageTitles, extractedPageTitles);
		}

		/// <summary>
		/// Validate the conversion of the epub file in the resources folder (EpubSample.epub)
		/// </summary>
		[TestMethod]
		public void ValidateEpubConversion()
		{
			var converter = new GenericFormatConverter();
			converter.ConvertEpubToOneNote(TestEpubPath, Utility.RootFolder);

			// retrieve xml from generated notebook
			var xmlDoc = _mOnGenerator.GetPageScopeHierarchy(_mNotebookId);
			Assert.IsNotNull(xmlDoc);

			// retrieve the section for the ppt conversion
			string sectionName = Path.GetFileNameWithoutExtension(TestEpubPath);
			XDocument xDoc = XDocument.Parse(xmlDoc);
			XElement xSection = xDoc.Descendants(_mXmlNs + "Section").FirstOrDefault(x => x.Attribute("name").Value.Equals(sectionName));
			Assert.IsNotNull(xSection);

			//get the xml of each pages
			var extractedPageTitles = xSection.Descendants(_mXmlNs + "Page").Select(x => x.Attribute("name").Value).ToList();
			CollectionAssert.AreEqual(EpubPageTitles, extractedPageTitles);
		}

		/// <summary>
		/// When the input file doesn't exist,
		/// word failed to open it and throw a COMException
		/// </summary>
		[TestMethod]
		[ExpectedException(typeof(COMException))]
		public void ValidateNonExistFileConversion()
		{
			var inputFile = Utility.NonExistentInputFile;
			var converter = new GenericFormatConverter();
			converter.ConvertWordToOneNote(inputFile, Utility.RootFolder);
		}

		/// <summary>
		/// When the output folder doesn't exist, 
		/// it should work without throwing any exceptions
		/// </summary>
		[TestMethod]
		public void ValidateNonExistOutputFolder()
		{
			var outputDir = Utility.NonExistentOutputPath;
			var converter = new GenericFormatConverter();
			converter.ConvertWordToOneNote(TestDocPath, outputDir);
		}
	}
}
