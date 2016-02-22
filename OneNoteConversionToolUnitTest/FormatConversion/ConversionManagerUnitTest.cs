using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OneNoteConversionToolUnitTest.FormatConversion
{
	/// <summary>
	/// UnitTest class for ConversionManager
	/// </summary>
	[TestClass]
	public class ConversionManagerUnitTest
	{
		private static string _mTestPath = string.Empty;
		private const string TestDoc = "test.doc";
		private const string TestDocx = "test.docx";
		private const string TestPdf = "test.pdf";
		private const string TestPpt = "test.ppt";
		private const string TestPptx = "test.pptx";
		private const string TestIndd = "test.indd";
		private const string TestUnsupported = "test.blah";

		/// <summary>
		/// Gets or sets the test context which provides
		/// information about and functionality for the current test run.
		///</summary>
		public TestContext TestContext { get; set; }

		/// <summary>
		/// Use ClassInitialize to run code before running the first test in the class
		/// </summary>
		/// <param name="testContext"></param>
		[ClassInitialize()]
		public static void Initialize(TestContext testContext)
		{
			_mTestPath = Path.GetTempPath() + "ConversionTest\\";
			
			if (!Directory.Exists(_mTestPath))
			{
				Directory.CreateDirectory(_mTestPath);
			}

			File.Create(_mTestPath + TestDoc);
			File.Create(_mTestPath + TestDocx);
			File.Create(_mTestPath + TestPdf);
			File.Create(_mTestPath + TestPpt);
			File.Create(_mTestPath + TestPptx);
			File.Create(_mTestPath + TestIndd);
			File.Create(_mTestPath + TestUnsupported);
		}

		/// <summary>
		/// Gets invoked after all unit tests have been run for this class
		/// </summary>
		[ClassCleanup()]
		public static void CleanUp()
		{
			if (Directory.Exists(_mTestPath))
			{
				Utility.DeleteDirectory(_mTestPath);
			}
		}

		/// <summary>
		/// Gets called for each test method
		/// </summary>
		[TestInitialize()]
		public void TestInitialize()
		{
			MockConversionManager.InitializeWithMockData();	
		}

		/// <summary>
		/// Verifies that the list supported formats are returned correctly
		/// </summary>
		[TestMethod]
		public void ValidateSupportedFormats()
		{
			var expectedFormats = new List<string>();
			expectedFormats.Add(MockFormatConverter.InputFormat);

			var supportedFormats = MockConversionManager.GetSupportedFormats();

			CollectionAssert.AreEqual(expectedFormats, supportedFormats);
		}

		/// <summary>
		/// Verifies that the ConvertWordToOneNote method is called for *.doc input format
		/// </summary>
		[TestMethod]
		public void ValidateConvertInputDoc()
		{
			MockConversionManager.ConvertInput(MockFormatConverter.InputFormat, _mTestPath + TestDoc, _mTestPath);
			Assert.IsTrue(MockConversionManager.GetMockFormatConverter().IsWordToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPdfToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPowerPointToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsInDesignToOneNoteCalled);
		}

		/// <summary>
		/// Verifies that the ConvertWordToOneNote method is called for *.docx input format
		/// </summary>
		[TestMethod]
		public void ValidateConvertInputDocx()
		{

			MockConversionManager.ConvertInput(MockFormatConverter.InputFormat, _mTestPath + TestDocx, _mTestPath);
			Assert.IsTrue(MockConversionManager.GetMockFormatConverter().IsWordToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPdfToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPowerPointToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsInDesignToOneNoteCalled);
		}

		/// <summary>
		/// Verifies that the ConvertPDFToOneNote method is called for *.pdf input format
		/// </summary>
		[TestMethod]
		public void ValidateConvertInputPdf()
		{
			MockConversionManager.ConvertInput(MockFormatConverter.InputFormat, _mTestPath + TestPdf, _mTestPath);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsWordToOneNoteCalled);
			Assert.IsTrue(MockConversionManager.GetMockFormatConverter().IsPdfToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPowerPointToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsInDesignToOneNoteCalled);
		}

		/// <summary>
		/// Verifies that the ConvertPowerPointToOneNote method is called for *.ppt input format
		/// </summary>
		[TestMethod]
		public void ValidateConvertInputPpt()
		{
			MockConversionManager.ConvertInput(MockFormatConverter.InputFormat, _mTestPath + TestPpt, _mTestPath);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsWordToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPdfToOneNoteCalled);
			Assert.IsTrue(MockConversionManager.GetMockFormatConverter().IsPowerPointToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsInDesignToOneNoteCalled);
		}

		/// <summary>
		/// Verifies that the ConvertPowerPointToOneNote method is called for *.pptx input format
		/// </summary>
		[TestMethod]
		public void ValidateConvertInputPptx()
		{
			MockConversionManager.ConvertInput(MockFormatConverter.InputFormat, _mTestPath + TestPptx, _mTestPath);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsWordToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPdfToOneNoteCalled);
			Assert.IsTrue(MockConversionManager.GetMockFormatConverter().IsPowerPointToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsInDesignToOneNoteCalled);
		}

		/// <summary>
		/// Verifies that the ConvertInDesignToOneNote method is called for *.indd input format
		/// </summary>
		[TestMethod]
		public void ValidateConvertInputIndd()
		{
			MockConversionManager.ConvertInput(MockFormatConverter.InputFormat, _mTestPath + TestIndd, _mTestPath);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsWordToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPdfToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPowerPointToOneNoteCalled);
			Assert.IsTrue(MockConversionManager.GetMockFormatConverter().IsInDesignToOneNoteCalled);
		}

		/// <summary>
		/// Verifies that an exception is thrown for a file format that is not expected
		/// </summary>
		[TestMethod]
		public void ValidateConvertInputNotSupportedFile()
		{
			MockConversionManager.ConvertInput(MockFormatConverter.InputFormat, _mTestPath + TestUnsupported, _mTestPath);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsWordToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPdfToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsPowerPointToOneNoteCalled);
			Assert.IsFalse(MockConversionManager.GetMockFormatConverter().IsInDesignToOneNoteCalled);
		}
	}
}
