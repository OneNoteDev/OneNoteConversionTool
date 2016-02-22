using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNoteConversionTool.FormatReaders;

namespace OneNoteConversionToolUnitTest.FormatReaders
{
	[TestClass]
	public class PowerPointOpenXmlUnitTest
	{
		private static readonly List<string> PptPageTitles = new List<string>() {"PresentationTitle", "DeletePage", "Third Slide", "PageTitle" };
		private static readonly List<string> PptNotes = new List<string>() { "", "", "", "note\n" };
		private static readonly List<string> PptComments = new List<string>() { "", "", "", "Author: Hichem Zakaria Aichour\nPageComment\n" };
		private static readonly List<string> PptSections = new List<string>() { "MainSection", "FirstSection", "Some New Section", "SecondSection" };
		private const int HiddenSlideNumber = 2;


		private const string TestPptFile = "..\\..\\Resources\\SectionSample.pptx";
		private static readonly string TestPptPath = Path.Combine(Environment.CurrentDirectory, TestPptFile);

		private static PowerPointOpenXml _mPptOpenXml;

		/// <summary>
		/// Initializer method for the Unit Test
		/// </summary>
		/// <param name="testContext"></param>
		[ClassInitialize()]
		public static void MyClassInitialize(TestContext testContext)
		{
			_mPptOpenXml = new PowerPointOpenXml(TestPptPath);
		}

		/// <summary>
		/// Validate obtaining slide titles
		/// </summary>
		[TestMethod]
		public void ValidateAllSlideTitles()
		{
			var titles = _mPptOpenXml.GetAllSlideTitles();
			CollectionAssert.AreEqual(PptPageTitles, titles);
		}

		/// <summary>
		/// Validate obtaining slide notes
		/// </summary>
		[TestMethod]
		public void ValidateAllSlideNotes()
		{
			var notes = _mPptOpenXml.GetAllSlideNotes();
			CollectionAssert.AreEqual(PptNotes, notes);
		}

		/// <summary>
		/// Validate obtaining slide comments
		/// </summary>
		[TestMethod]
		public void ValidateAllSlideComments()
		{
			var comments = _mPptOpenXml.GetAllSlideComments();
			Console.WriteLine(comments[3]);
			Console.WriteLine(PptComments[3]);
			Console.WriteLine(comments[3]);
			CollectionAssert.AreEqual(PptComments, comments);
		}

		/// <summary>
		/// Validate detecting hidden slides
		/// </summary>
		[TestMethod]
		public void ValidateHiddenSlide()
		{
			Assert.IsTrue(_mPptOpenXml.IsHiddenSlide(HiddenSlideNumber));
		}

		/// <summary>
		/// Validate obtaining section names
		/// </summary>
		[TestMethod]
		public void ValidateSectionNames()
		{
			var sections = _mPptOpenXml.GetSectionNames();
			CollectionAssert.AreEqual(PptSections, sections);
		}

		/// <summary>
		/// Validate using non-existing input file
		/// </summary>
		[TestMethod]
		[ExpectedException(typeof(FileNotFoundException))]
		public void ValidateNonExistFile()
		{
			const string inputFile = Utility.NonExistentInputFile;
			var pptOpenXml = new PowerPointOpenXml(inputFile);

			var N = pptOpenXml.NumberOfSlides();
		}
	}
}
