using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNoteConversionTool.FormatConversion;
using OneNoteConversionTool.OutputGenerator;

namespace OneNoteConversionToolUnitTest.FormatConversion
{
	[TestClass]
	public class SmsgrFormatConverterUnitTest
	{
		private static readonly List<string> PptPageTitles = new List<string>() { "Table of Contents", "MainSection", "PresentationTitle", "FirstSection", "Some New Section", "Third Slide", "SecondSection", "PageTitle" };

		private const string TestPptName = "..\\..\\Resources\\SectionSample.pptx";
		private static readonly string TestPptPath = Path.Combine(Environment.CurrentDirectory, TestPptName);
		private const string TrainerNotebookName = "Trainer Notebook";
		private const string StudentNotebookName = "Student Notebook";

		private static XNamespace _mXmlNs;
		private static OneNoteGenerator _mOnGenerator;
		private static string _mTrainerNotebookId = String.Empty;
		private static string _mStudentNotebookId = String.Empty;

		/// <summary>
		/// Create temporary folders and initialize onenote generator
		/// </summary>
		/// <param name="testContext"></param>
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
			// Get the id of the Trainer and student notebooks
			_mTrainerNotebookId = _mOnGenerator.CreateNotebook(TrainerNotebookName);
			_mStudentNotebookId = _mOnGenerator.CreateNotebook(StudentNotebookName);
		}

		/// <summary>
		/// Delete temporary folders at cleanup
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
		/// Validate the conversion of the ppt file in the resources folder (SectionSample.pptx)
		/// </summary>
		[TestMethod]
		public void ValidatePowerPointConversion()
		{
			var converter = new SmsgrFormatConverter();
			converter.ConvertPowerPointToOneNote(TestPptPath, Utility.RootFolder);

			ValidateTrainerNotebook();
			ValidateStudentNotebook();
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
			var converter = new SmsgrFormatConverter();
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
			var converter = new SmsgrFormatConverter();
			converter.ConvertPowerPointToOneNote(TestPptPath, outputDir);
		}

		/// <summary>
		/// Validate the Trainer notebook
		/// </summary>
		private static void ValidateTrainerNotebook()
		{
			// retrieve xml from generated notebook
			var xmlDoc = _mOnGenerator.GetPageScopeHierarchy(_mTrainerNotebookId);
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
		/// Validate the student notebook
		/// </summary>
		private static void ValidateStudentNotebook()
		{
			// retrieve xml from generated notebook
			var xmlDoc = _mOnGenerator.GetPageScopeHierarchy(_mStudentNotebookId);
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
	}
}
