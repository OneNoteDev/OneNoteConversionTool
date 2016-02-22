using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using HtmlAgilityPack;
using OneNoteConversionTool.OutputGenerator;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace OneNoteConversionTool.FormatConversion
{
	/// <summary>
	/// Converter that accepts Kindercare format
	/// </summary>
	public class KindercareFormatConverter : GenericFormatConverter
	{
		private const string InputFormat = "Kindercare";

		public override string GetSupportedInputFormat()
		{
			return InputFormat;
		}

		/// <summary>
		/// Converts Html file into OneNote Section of Pages
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="originalPicFolder"></param>
		/// <param name="auxInputFolder"></param>
		/// <param name="onGenerator"></param>
		/// <param name="sectionId"></param>
		/// <returns></returns>
		protected override bool ConvertHtmlToOneNote(string inputFile, string originalPicFolder, string auxInputFolder, OneNoteGenerator onGenerator, string sectionId)
		{
			var retVal = false;

			try
			{
				string htmlContent;
				using (var sr = new StreamReader(inputFile, Encoding.Default))
				{
					htmlContent = sr.ReadToEnd();
				}

				var doc = new HtmlDocument();
				doc.LoadHtml(htmlContent);

				var content = doc.DocumentNode;

				//outer html contains format information
				var htmlBox = doc.DocumentNode;
				//list of onenote page Id 
				var pageIdList = new List<string>();
				//list of onenote page/subpage name
				var pageNameList = new List<string>();
				//separate the whole doc into pages
				var pages = SeparatePages(content) as List<List<HtmlNode>>;
				//get chapter names from table contents
				var chapterNameList = new List<string>();

				//may produce problem by calling first()
				//maybe empty page with pageNameList[0] = null, fix later
				var tableContent = pages.FirstOrDefault();
				if (tableContent != null)
				{
					foreach (var node in tableContent)
					{
						chapterNameList.Add(node.InnerText.Replace("\r\n", " ").Trim());
					}
				}

				//store errors occurred during conversion 
				var errorList = new List<Dictionary<InfoType, string>>();

				//print pages to onenote
				foreach (var page in pages)
				{
					var pageTitle = GetPageTitle(page);
					var pageId = onGenerator.CreatePage(pageTitle, sectionId);

					var errorInfo = new Dictionary<InfoType, string>();
					var pageContent = GeneratePageContent(htmlBox, page, originalPicFolder, auxInputFolder, errorInfo);
					onGenerator.AddPageContentAsHtmlBlock(pageId, pageContent);

					pageIdList.Add(pageId);
					pageNameList.Add(pageTitle);

					if (errorInfo.Count > 0)
					{
						errorInfo.Add(InfoType.Id, pageId);
						errorInfo.Add(InfoType.Title, pageTitle);
						errorList.Add(errorInfo);
					}
				}
				if (errorList.Count > 0)
				{
					CreateErrorPage(onGenerator, sectionId, errorList);
				}
				//set some pages as subpages according to observed rules
				var lastSeenChapterName = String.Empty;
				for (int i = 0; i < pageIdList.Count; i++)
				{
					var pageId = pageIdList[i];
					var name = pageNameList[i];
					//check if the page name is a chapter name (a substring of table content items)
					var isChapterName = chapterNameList.Any(x => x.Contains(name));

					bool isSubpage = !isChapterName && name != lastSeenChapterName;

					//if it is a new chapter name, set it as a page
					//if it is not a chapter name, but it is the first of consecutive same name pages, set it as a page
					if (isChapterName == false 
						&& i > 0 && name != pageNameList[i - 1] 
						&& i < pageNameList.Count && name == pageNameList[i + 1])
					{
						isSubpage = false;
					}

					if (isSubpage)
					{
						onGenerator.SetSubPage(sectionId, pageId);
					}
					if (isChapterName)
					{
						lastSeenChapterName = name;
					}
				}
				retVal = true;
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
				retVal = false;
			}

			return retVal;
		}
	}
}
