using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNoteConversionTool.OutputGenerator;

namespace OneNoteConversionToolUnitTest.OutputGenerator
{
	[TestClass]
	public class OneNoteGeneratorUnitTest
	{
		private static XNamespace _mXmlNs;

		private static OneNoteGenerator _mOneNoteGenerator;

		private const string TestNotebookName = "TestNotebook";
		private const string TestSectionName = "TestSection";
		private const string TestPageName = "TestPage";
		private static string _mTestNotebookDir = String.Empty;
		private static string _mTestNotebookId = String.Empty;
		private static string _mTestSectionId = String.Empty;
		private static string _mTestPageId = String.Empty;

		/// <summary>
		/// Initializes Test notebook for UnitTesting
		/// </summary>
		/// <param name="testContext"></param>
		[ClassInitialize]
		public static void InitializeTestNotebook(TestContext testContext)
		{
			_mXmlNs = Utility.NS;

			_mTestNotebookDir = Path.GetTempPath();

			_mOneNoteGenerator = new OneNoteGenerator(_mTestNotebookDir);
			_mTestNotebookId = _mOneNoteGenerator.CreateNotebook(TestNotebookName);
			_mTestSectionId = _mOneNoteGenerator.CreateSection(TestSectionName, _mTestNotebookId);
			//for ValidateCreateSectionNameConflicts()
			_mOneNoteGenerator.CreateSection(TestSectionName, _mTestNotebookId);
			_mOneNoteGenerator.CreateSection(TestSectionName, _mTestNotebookId);

			_mTestPageId = _mOneNoteGenerator.CreatePage(TestPageName, _mTestSectionId);
			
			//This is ugly, but apparently creating notebook/section/page takes time
			Thread.Sleep(4000);
		}

		/// <summary>
		/// Remove Test notebook that was created for UnitTesting
		/// </summary>
		[ClassCleanup]
		public static void CleanupTestNotebook()
		{
			if (Directory.Exists(_mTestNotebookDir + TestNotebookName))
				Utility.DeleteDirectory(_mTestNotebookDir + TestNotebookName);
		}

		/// <summary>
		/// Validates create notebook method
		/// </summary>
		[TestMethod]
		public void ValidateCreateNotebook()
		{
			Assert.IsTrue(Directory.Exists(_mTestNotebookDir + TestNotebookName));

			var notebookHierarchy = _mOneNoteGenerator.GetChildrenScopeHierarchy(_mTestNotebookId);
			Assert.IsNotNull(notebookHierarchy);

			XDocument xdoc = XDocument.Parse(notebookHierarchy);
			var node = xdoc.Elements(_mXmlNs + "Notebook")
				.Single(x => x.Attribute("name").Value.Equals(TestNotebookName));

			Assert.IsNotNull(node);
			Assert.IsTrue(node.Attribute("name").Value.Equals(TestNotebookName));
			Assert.IsTrue(node.Attribute("ID").Value.Equals(_mTestNotebookId));
		}

		/// <summary>
		/// Validates create section method
		/// </summary>
		[TestMethod]
		public void ValidateCreateSection()
		{
			Assert.IsTrue(File.Exists(Path.Combine(_mTestNotebookDir, TestNotebookName, TestSectionName + ".one")));
			
			var sectionHierarchy = _mOneNoteGenerator.GetSectionScopeHierarchy(_mTestNotebookId);
			Assert.IsNotNull(sectionHierarchy);

			var xdoc = XDocument.Parse(sectionHierarchy);
			var node = xdoc.Descendants(_mXmlNs + "Section")
				.Single(x => x.Attribute("name").Value.Equals(TestSectionName));

			Assert.IsNotNull(node);
			Assert.IsTrue(node.Attribute("name").Value.Equals(TestSectionName));
			Assert.IsTrue(node.Attribute("ID").Value.Equals(_mTestSectionId));
		}

		/// <summary>
		/// Validate CreateSection() when section name already exist
		/// </summary>
		[TestMethod]
		public void ValidateCreateSectionNameConflicts()
		{
			var sectionHierarchy = _mOneNoteGenerator.GetSectionScopeHierarchy(_mTestNotebookId);
			Assert.IsNotNull(sectionHierarchy);

			var xdoc = XDocument.Parse(sectionHierarchy);

			const string testSectionName2 = TestSectionName + " (2)";
			const string testSectionName3 = TestSectionName + " (3)";
			var expectedSectionNames = new List<string>() 
			{ 
				TestSectionName, 
				testSectionName2,
				testSectionName3
			};
			var sectionNames = xdoc.Descendants(_mXmlNs + "Section").Select(x => x.Attribute("name").Value).ToList();
			CollectionAssert.AreEqual(expectedSectionNames, sectionNames);
		}

		/// <summary>
		/// Validates create section method
		/// </summary>
		[TestMethod]
		public void ValidateCreatePage()
		{
			var pageHierarchy = _mOneNoteGenerator.GetChildrenScopeHierarchy(_mTestPageId);
			Assert.IsNotNull(pageHierarchy);

			XDocument xdoc = XDocument.Parse(pageHierarchy);
			var node = xdoc.Elements(_mXmlNs + "Page")
				.Single(x => x.Attribute("name").Value.Equals(TestPageName));

			Assert.IsNotNull(node);
			Assert.IsTrue(node.Attribute("name").Value.Equals(TestPageName));
			Assert.IsTrue(node.Attribute("ID").Value.Equals(_mTestPageId));
		}

		/// <summary>
		/// Validates Unescape XML function
		/// </summary>
		[TestMethod]
		public void ValidateUnescapeXml()
		{
			string doc = "&apos;&quot;test&gt;&lt;&amp;";
			const string expected = "'\"test><&";
			string actual = OneNoteConversionTool.Utility.UnescapeXml(doc);

			Assert.AreEqual(expected, actual);
		}

		/// <summary>
		/// Validates Create Table of Content Page function
		/// </summary>
		[TestMethod]
		public void ValidateCreateTableOfContentPage()
		{
			const string tocPageTitle = "Test Table of Content";
			var tocPageId = _mOneNoteGenerator.CreateTableOfContentPage(_mTestSectionId, tocPageTitle);

			var sectionHierarchy = _mOneNoteGenerator.GetChildrenScopeHierarchy(_mTestSectionId);
			Assert.IsNotNull(sectionHierarchy);

			XDocument xdoc = XDocument.Parse(sectionHierarchy);
			var node = xdoc.Elements(_mXmlNs + "Section")
				.Single(x => x.Attribute("name").Value.Equals(TestSectionName))
				.Elements(_mXmlNs + "Page")
				.Single(x => x.Attribute("ID").Value.Equals(tocPageId));

			Assert.IsNotNull(node);
			Assert.IsTrue(node.Attribute("name").Value.Equals(tocPageTitle));
			Assert.IsTrue(node.Attribute("ID").Value.Equals(tocPageId));
		}
	}
}
