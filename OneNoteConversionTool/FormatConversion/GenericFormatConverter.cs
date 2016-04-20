using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;
using Ghostscript.NET.Rasterizer;
using HtmlAgilityPack;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using OneNoteConversionTool.FormatReaders;
using OneNoteConversionTool.OutputGenerator;
using OneNoteConversionTool.Properties;
using Image = System.Drawing.Image;
using WordApplication = Microsoft.Office.Interop.Word.Application;
using _WordApplication = Microsoft.Office.Interop.Word._Application;
using _WordDocument = Microsoft.Office.Interop.Word._Document;
using Path = System.IO.Path;
using System.Diagnostics.CodeAnalysis;

#if INDESIGN_INSTALLED
using InDesign;
using InDesignApplication = InDesign.Application;
using InDesignDocument = InDesign.Document;
#endif

namespace OneNoteConversionTool.FormatConversion
{
	/// <summary>
	/// Handles the generic file format conversion into OneNote
	/// </summary>
	public class GenericFormatConverter : IFormatConverter
	{
		#region Global Variables
		protected const string DefaultPageTitle = "Untitled";
		protected const string HtmlFileExtension = ".htm";
		protected const string XmlFileExtension = ".xml";
		protected const string AuxFileSuffix = "_aux"; 
		protected const string AuxFileLocationSuffix = "_files";
		protected const string ErrorImageHeight = "180";
		protected const string ErrorImageWidth = "180";
		protected const string ErrorPageTitle = "Error Page";
		protected const string ErrorImageName = "error.jpg";
		protected enum InfoType { Message, Title, Id };

		private const string InputFormat = "Generic";
		#endregion

		#region Public Methods
		/// <summary>
		/// Returns the name of the input format that this IFormatConverter supports
		/// </summary>
		/// <returns></returns>
		public virtual string GetSupportedInputFormat()
		{
			return InputFormat;
		}

		/// <summary>
		/// Converts the Word document input file into OneNote
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		public virtual bool ConvertWordToOneNote(string inputFile, string outputDir)
		{
			var retVal = false;

			string inputFileName = Path.GetFileNameWithoutExtension(inputFile);

			var tempPath = outputDir + "\\temp\\";
			if (!Directory.Exists(tempPath))
			{
				Directory.CreateDirectory(tempPath);
			}

			object tempHtmlFilePath = tempPath + inputFileName + HtmlFileExtension;
			object htmlFilteredFormat = WdSaveFormat.wdFormatFilteredHTML;

			//used to get pictures in word document
			object auxiliaryFilePath = tempPath + inputFileName + AuxFileSuffix;
			object htmlFormat = WdSaveFormat.wdFormatHTML;
			var auxiliaryFilePicFolder = auxiliaryFilePath + AuxFileLocationSuffix;
			var originalFilePicFolder = tempPath + inputFileName + AuxFileLocationSuffix;

			try
			{
				//Save the word document as a HTML file temporarily
				var word = new WordApplication();
				var doc = word.Documents.Open(inputFile);

				doc.SaveAs(ref tempHtmlFilePath, ref htmlFilteredFormat);
				doc.SaveAs(ref auxiliaryFilePath, ref htmlFormat);

                ((_WordDocument)doc).Close();
                ((_WordApplication)word).Quit();

				//Create a new OneNote Notebook
				var note = new OneNoteGenerator(outputDir);
				var notebookId = note.CreateNotebook(GetSupportedInputFormat());
				var sectionId = note.CreateSection(inputFileName, notebookId);

				//Now migrate the content in the temporary HTML file into the newly created OneNote Notebook
				if (ConvertHtmlToOneNote(tempHtmlFilePath.ToString(), originalFilePicFolder, auxiliaryFilePicFolder, note, sectionId))
				{
					//Generate table of content
					note.CreateTableOfContentPage(sectionId);
					retVal = true;
				}
			}
			finally
			{
				Utility.DeleteDirectory(tempPath);
			}

			return retVal;
		}

		/// <summary>
		/// Converts PDF file to OneNote by including an image for each page in the document
		/// </summary>
		/// <param name="inputFile">PDF document path</param>
		/// <param name="outputDir">Directory of the output OneNote Notebook</param>
		/// <returns></returns>
		public virtual bool ConvertPdfToOneNote(string inputFile, string outputDir)
		{
			//Get the name of the file
			string inputFileName = Path.GetFileNameWithoutExtension(inputFile);

			//Create a new OneNote Notebook
			var note = new OneNoteGenerator(outputDir);
			string notebookId = note.CreateNotebook(GetSupportedInputFormat());
			string sectionId = note.CreateSection(inputFileName, notebookId);

			using (var rasterizer = new GhostscriptRasterizer())
			{
				rasterizer.Open(inputFile);
				for (var i = 1; i <= rasterizer.PageCount; i++)
				{
					Image img = rasterizer.GetPage(160, 160, i);
					MemoryStream stream = new MemoryStream();
					img.Save(stream, ImageFormat.Png);
					img = Image.FromStream(stream);

					string pageId = note.CreatePage(String.Format("Page{0}", i), sectionId);
					note.AddImageToPage(pageId, img);
				}
			}

			note.CreateTableOfContentPage(sectionId);

			return true;
		}

		/// <summary>
		/// Converts PowerPoint presentation into OneNote section
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		/// <returns></returns>
		public virtual bool ConvertPowerPointToOneNote(string inputFile, string outputDir)
		{
			// Get the name of the file
			string inputFileName = Path.GetFileNameWithoutExtension(inputFile);

			// Convert presentation slides to images
			string imgsPath = ConvertPowerPointToImages(inputFile, outputDir);

			// Create a new OneNote Notebook
			var note = new OneNoteGenerator(outputDir);

			// Convert to OneNote
			var pptOpenXml = new PowerPointOpenXml(inputFile);
			ConvertPowerPointToOneNote(pptOpenXml, imgsPath, note, inputFileName);

			// Delete the temperory imgs directory
			Utility.DeleteDirectory(imgsPath);

			return true;
		}

		/// <summary>
		/// Converts the InDesign document input file into OneNote
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		/// <returns></returns>
		public virtual bool ConvertInDesignToOneNote(string inputFile, string outputDir)
		{
#if INDESIGN_INSTALLED
			// get the file name
			string inputFileName = Path.GetFileNameWithoutExtension(inputFile);

			// gets the InDesign App
			Type inDesignAppType = Type.GetTypeFromProgID(ConfigurationManager.AppSettings.Get("InDesignProgId"));
            if (inDesignAppType == null) throw new InvalidOperationException("Failed to find InDesign application. Please ensure that it is installed and is running on your machine.");

			InDesignApplication app = (InDesignApplication)Activator.CreateInstance(inDesignAppType, true);

			// open the indd file
			try
			{
				app.Open(inputFile);
			}
			catch (Exception e)
			{
				Console.WriteLine(@"Error in ConvertInDesignToOneNote for file {0}: {1}", inputFile, e.Message);
				return false;
			}
			InDesignDocument doc = app.ActiveDocument;

			// temp directory for html
			string htmlDir = Utility.CreateDirectory(Utility.NewFolderPath(outputDir, "html"));

			// Get the file name
			string htmlFile = Utility.NewFilePath(htmlDir, inputFileName, HtmlFileExtension);

			// Save the html file
			SetInDesignHtmlExportOptions(doc);
			doc.Export(idExportFormat.idHTML, htmlFile);

			// Create a new OneNote Notebook
			var note = new OneNoteGenerator(outputDir);
			string notebookId = note.CreateNotebook(GetSupportedInputFormat());
			string sectionId = note.CreateSection(inputFileName, notebookId);


			// get the html
			var htmlDoc = new HtmlDocument();
			htmlDoc.Load(htmlFile);

			// change the links to have the full path
			ModifyInDesignHtmlLinks(htmlDoc, htmlDir);

			// get the title
			string title = GetInDesignTitle(htmlDoc);

			string pageId = note.CreatePage(title, sectionId);
			note.AddPageContentAsHtmlBlock(pageId, htmlDoc.DocumentNode.OuterHtml);

			return true;
#else		
			throw new NotImplementedException();
#endif
		}

		/// <summary>
		/// Converts the ePub document input file into OneNote
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		/// <returns></returns>
		public virtual bool ConvertEpubToOneNote(string inputFile, string outputDir)
		{
			try
			{
				// Initialize epub reader class
				var epub = new EpubReader(inputFile, outputDir);

				// Get the page contents
				List<string> pageTitles = epub.GetPageTitles();
				List<HtmlDocument> pagesHtml = epub.GetPagesAsHtmlDocuments();
				List<string> pagePaths = epub.GetPagePaths();
				Dictionary<string, int> pagesLevel = epub.GetPagesLevel();

				// Create a new OneNote Notebook
				var note = new OneNoteGenerator(outputDir);
				string notebookId = note.CreateNotebook(GetSupportedInputFormat());
				string sectionId = note.CreateSection(epub.GetTitle(), notebookId);

				// Create pages
				var pageIds = pageTitles.Select(pageTitle => note.CreatePage(pageTitle, sectionId)).ToArray();

				// Get links to pages
				var pageLinks = pageIds.Select(pageId => note.GetHyperLinkToObject(pageId)).ToList();

				// Replace links to .html with .one
				ReplaceEpubLinksFromHtmlToOneNote(pagesHtml, pagePaths, pageLinks);

				// Add content to pages
				for (var i = 0; i < pageIds.Length; i++)
				{
					note.AppendPageContentAsHtmlBlock(pageIds[i], pagesHtml[i].DocumentNode.OuterHtml);
					if (pagesLevel.ContainsKey(pagePaths[i]))
					{
						note.SetPageLevel(pageIds[i], pagesLevel[pagePaths[i]]);
					}
				}

				return true;
			}
			catch (Exception e)
			{
				Console.WriteLine(@"Error in ConvertEpubToOneNote for file {0}: {1}", inputFile, e.Message);
				return false;
			}
		}
		#endregion

		#region MS Word Helper Methods
		/// <summary>
		/// Converts Html file into OneNote Section of Pages
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="originalPicFolder"></param>
		/// <param name="auxPicFolder"></param>
		/// <param name="onGenerator"></param>
		/// <param name="sectionId"></param>
		/// <returns></returns>
		protected virtual bool ConvertHtmlToOneNote(string inputFile, string originalPicFolder, string auxPicFolder, OneNoteGenerator onGenerator, string sectionId)
		{
			var retVal = false;
			var htmlContent = string.Empty;

			try
			{
				using (var sr = new StreamReader(inputFile, Encoding.Default))
				{
					htmlContent = sr.ReadToEnd();
				}

				var htmlDoc = new HtmlDocument();
				htmlDoc.LoadHtml(htmlContent);

				//Preserve the HTML surrounding tags
				var content = htmlDoc.DocumentNode;

				var previousPageTitle = string.Empty;

				//store error info
				var errorList = new List<Dictionary<InfoType, string>>();

				foreach (var page in SeparatePages(content))
				{
					if (page != null)
					{
						//Get page title
						var pageTitle = GetPageTitle(page);

						//Create the page
						var pageId = onGenerator.CreatePage(pageTitle, sectionId);

						//error info
						var errorInfo = new Dictionary<InfoType, string>();

						//Add the content of the page
						var pageContent = GeneratePageContent(content, page, originalPicFolder, auxPicFolder, errorInfo);
						onGenerator.AddPageContentAsHtmlBlock(pageId, pageContent);

						//Attempt to do subpage
						if (!pageTitle.Equals(DefaultPageTitle, StringComparison.CurrentCultureIgnoreCase) &&
							pageTitle.Equals(previousPageTitle, StringComparison.CurrentCultureIgnoreCase))
						{
							onGenerator.SetSubPage(sectionId, pageId);
						}

						previousPageTitle = pageTitle;

						if (errorInfo.Count > 0)
						{
							errorInfo.Add(InfoType.Id, pageId);
							errorInfo.Add(InfoType.Title, pageTitle);
							errorList.Add(errorInfo);
						}
					}
				}
				if (errorList.Count > 0)
				{
					CreateErrorPage(onGenerator, sectionId, errorList);
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

		/// <summary>
		/// Create the error page for the section, which contains all the errors occurred during the conversion
		/// </summary>
		/// <param name="onGenerator"></param>
		/// <param name="sectionId"></param>
		/// <param name="errorList"></param>
		protected void CreateErrorPage(OneNoteGenerator onGenerator, string sectionId, IEnumerable<IDictionary<InfoType, string>> errorList)
		{
			var pageId = onGenerator.CreatePage(ErrorPageTitle, sectionId);
			var pageContent = String.Empty;
			foreach (var error in errorList)
			{
				var hyperLink = onGenerator.GetHyperLinkToObject(error[InfoType.Id]);
				pageContent += string.Format("<a href=\"{0}\">{1} : \n{2}</a>", hyperLink, error[InfoType.Title], error[InfoType.Message]) + "\n\n";
			}
			onGenerator.AddPageContent(pageId, pageContent);
		}

		/// <summary>
		/// Splits HtmlNodes into List of HtmlNodes based on page breaks
		/// </summary>
		/// <param name="content"></param>
		/// <returns></returns>
		protected virtual IEnumerable<IEnumerable<HtmlNode>> SeparatePages(HtmlNode content)
		{
			var result = new List<List<HtmlNode>>();
			result.Add(new List<HtmlNode>());

			var bodyNode = content.Descendants("body").FirstOrDefault();

			if (bodyNode != null)
			{
				for (var node = bodyNode.FirstChild; node.NextSibling != null; node = node.NextSibling)
				{
					var nodeList = new List<HtmlNode>();
					if (node.Name.Equals("div"))
					{
						for (var tempNode = node.FirstChild; tempNode.NextSibling != null; tempNode = tempNode.NextSibling)
						{
							nodeList.Add(tempNode);
						}
					}
					else
					{
						nodeList.Add(node);
					}

					foreach (var tempNode in nodeList)
					{
						if (tempNode.SelectNodes(".//br[@style='page-break-before:always']") != null)
						{
							if (result.Last().Any())
							{
								result.Add(new List<HtmlNode>());
							}
						}
						else
						{
							result.Last().Add(tempNode);
						}
					}
				}
			}

			if (result.Last().All(x => String.IsNullOrWhiteSpace(x.InnerText) || x.InnerText.Equals("&nbsp;")))
			{
				result.Remove(result.Last());
			}
			return result;
		}

		/// <summary>
		/// Grabs the page node to be used as title and remove it from the pageNodes
		/// </summary>
		/// <param name="pageNodes"></param>
		/// <returns></returns>
		protected virtual string GetPageTitle(IEnumerable<HtmlNode> pageNodes)
		{
			var retVal = DefaultPageTitle;
			var pageNodesList = pageNodes as List<HtmlNode>;

			if (pageNodesList != null)
			{
				// Search if there is a HtmlNode we can grab the text and use as title
				// TODO: should we check if font is greater than some size?
				var pageTitleNode = pageNodesList.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.InnerText.Trim()) &&
																	  !x.InnerText.Trim()
																		  .Equals("&nbsp;", StringComparison.InvariantCultureIgnoreCase));

				if (pageTitleNode != null)
				{
					// Use the text as the page title, removing newlines if necessary
					// and also remove the node so it doesn't get duplicated in the body
					retVal = pageTitleNode.InnerText.Replace("\r\n", " ").Trim();
					pageNodesList.Remove(pageTitleNode);
				}
			}

			return retVal;
		}

		/// <summary>
		/// Generate Html page content from List of HtmlNode
		/// </summary>
		protected virtual string GeneratePageContent(
			HtmlNode htmlBox,
			IEnumerable<HtmlNode> page,
			string originalPicFolderPath,
			string auxPicFolderPath,
			IDictionary<InfoType, string> errorInfo)
		{
			HtmlNode htmlFrame;

			if (htmlBox.Descendants("body").Any())
			{
				htmlFrame = htmlBox.Descendants("body").First();
			}
			else
			{
				const string template = @"<html><body></body></html>";
				var htmlDoc = new HtmlDocument();
				htmlDoc.LoadHtml(template);
				htmlFrame = htmlDoc.DocumentNode;
			}
			htmlFrame.RemoveAllChildren();

			foreach (var node in page)
			{
				try
				{
					// We need to manipulate img nodes because Word may not have converted them properly
					if (node.Descendants("img").Any())
					{
						foreach (var img in node.Descendants("img"))
						{
							//URL decoding
							var imagePath = HttpUtility.UrlDecode(Path.GetDirectoryName(img.Attributes["src"].Value));
							var imageName = Path.GetFileName(img.Attributes["src"].Value);
							//use auxPicFolderPath to get the image
							img.Attributes["src"].Value = Path.Combine(originalPicFolderPath, imageName);

							// If image doesn't exist in the originalPicFolderPath location,
							if (!File.Exists(img.Attributes["src"].Value))
							{
								// see if it exists in the auxPicFolderPath
								if (File.Exists(Path.Combine(auxPicFolderPath, imageName)))
								{
									img.Attributes["src"].Value = Path.Combine(auxPicFolderPath, imageName);
								}
								else
								{
									//can't find the image, substitute it with ErrorImage, then throw exception
									img.Attributes["src"].Value = GetErrorImagePath(auxPicFolderPath);
									img.Attributes["height"].Value = ErrorImageHeight;
									img.Attributes["width"].Value = ErrorImageWidth;
									htmlFrame.AppendChild(node);
									throw new Exception("Failed to import image: Image Not Found");
								}
							}
							//scale image size
							if (File.Exists(img.Attributes["src"].Value))
							{
								using (var image = Image.FromFile(img.Attributes["src"].Value))
								{
									int height = image.Height * int.Parse(img.Attributes["width"].Value) / image.Width;
									img.Attributes["height"].Value = height.ToString();
								}
							}
						}
					}
					htmlFrame.AppendChild(node);
				}
				catch (Exception e)
				{
					if (errorInfo.ContainsKey(InfoType.Message))
					{
						errorInfo[InfoType.Message] += "\n" + e.Message;
					}
					else
					{
						errorInfo.Add(InfoType.Message, e.Message);
					}
				}
			}
			return htmlBox.OuterHtml;
		}

		/// <summary>
		/// a big error image to indicate user where had failed to import an image
		/// </summary>
		/// <param name="picFolderPath"></param>
		/// <returns></returns>
		private string GetErrorImagePath(string picFolderPath)
		{
			var tempPath = Path.Combine(picFolderPath, ErrorImageName);
			if (!File.Exists(tempPath))
			{
				var image = Resources.redX;
				image.Save(tempPath);
			}
			return tempPath;
		}
		#endregion

		#region MS PowerPoint Helper Methods
		/// <summary>
		/// Converts PowerPoint presentation to images where is slide is given by one image
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		/// <returns>the directory of the images</returns>
		protected string ConvertPowerPointToImages(string inputFile, string outputDir)
		{
			// temp directory for images
			string imgsPath = Utility.CreateDirectory(Utility.NewFolderPath(outputDir, "imgs"));

			//Get the presentation
			var powerPoint = new Microsoft.Office.Interop.PowerPoint.Application();
			var presentation = powerPoint.Presentations.Open(inputFile);

			// Save each slide as picture
			try
			{
				presentation.SaveAs(imgsPath, PpSaveAsFileType.ppSaveAsPNG);
			}
			catch (COMException e)
			{
                Console.WriteLine("Presenation doesn't have any slide.");
                Console.WriteLine(e.Message);
				return String.Empty;
			}
			finally
			{
				// Close power point application
				presentation.Close();
				powerPoint.Quit();
			}
			return imgsPath;
		}

		/// <summary>
		/// Converts PowerPoint presentan to OneNote while converting the sections in power point to main pages, and slides to sub pages
		/// </summary>
		/// <param name="pptOpenXml"></param>
		/// <param name="imgsPath"></param>
		/// <param name="note"></param>
		/// <param name="sectionName"></param>
		protected virtual void ConvertPowerPointToOneNote(PowerPointOpenXml pptOpenXml, string imgsPath, OneNoteGenerator note,
			string sectionName)
		{
			string notebookId = note.CreateNotebook(GetSupportedInputFormat());
			string sectionId = note.CreateSection(sectionName, notebookId);

			if (pptOpenXml.HasSections())
			{
				List<string> sectionNames = pptOpenXml.GetSectionNames();
				List<List<int>> slidesInSections = pptOpenXml.GetSlidesInSections();
				var pptSectionsPageIds = new List<string>();
				for (int i = 0; i < sectionNames.Count; i++)
				{
					string pptSectionPageId = note.CreatePage(sectionNames[i], sectionId);
					foreach (var slideNumber in slidesInSections[i])
					{
						string pageId = InsertPowerPointSlideInOneNote(slideNumber, pptOpenXml, imgsPath, note, sectionId);
						if (!String.IsNullOrEmpty(pageId))
						{
							note.SetSubPage(sectionId, pageId);
						}
					}
					pptSectionsPageIds.Add(pptSectionPageId);
				}

				note.CreateTableOfContentPage(sectionId);

				foreach (var pptSectionPageId in pptSectionsPageIds)
				{
					note.SetCollapsePage(pptSectionPageId);
				}
			}
			else
			{
				for (var i = 1; i <= pptOpenXml.NumberOfSlides(); i++)
				{
					InsertPowerPointSlideInOneNote(i, pptOpenXml, imgsPath, note, sectionId);
				}
			}
		}

		/// <summary>
		/// Inserts a power point slide into a given section in OneNote as a page
		/// </summary>
		/// <param name="slideNumber"></param>
		/// <param name="pptOpenXml"></param>
		/// <param name="imgsPath"></param>
		/// <param name="note"></param>
		/// <param name="sectionId"></param>
		/// <param name="showComments"></param>
		/// <param name="commentsStr"></param>
		/// <param name="showNotes"></param>
		/// <param name="notesStr"></param>
		/// <param name="hiddenSlideNotIncluded"></param>
		/// <returns>the page ID</returns>
		protected string InsertPowerPointSlideInOneNote(int slideNumber, PowerPointOpenXml pptOpenXml, string imgsPath,
			OneNoteGenerator note, string sectionId, bool showComments = true, string commentsStr = "Comments",
			bool showNotes = true, string notesStr = "Notes", bool hiddenSlideNotIncluded = true)
		{
			// skip hidden slides
			if (hiddenSlideNotIncluded && pptOpenXml.IsHiddenSlide(slideNumber))
			{
				return String.Empty;
			}

			// get the image representing the current slide as HTML
			string imgPath = String.Format("{0}\\Slide{1}.png", imgsPath, slideNumber);
			Image img;
			try
			{
				img = Image.FromFile(imgPath);
			}
			catch (FileNotFoundException e)
			{
				Console.WriteLine("Slide {0} was not converted", slideNumber);
                Console.WriteLine(e.Message);
				img = null;
			}

			// insert the image
			string pageTitle = pptOpenXml.GetSlideTitle(slideNumber);
			pageTitle = String.IsNullOrEmpty(pageTitle) ? String.Format("Slide{0}", slideNumber) : pageTitle;
			string pageId = note.CreatePage(pageTitle, sectionId);
			if (img != null)
			{
				note.AddImageToPage(pageId, img);
				img.Dispose();
			}

			// Add comments
			string slideComments = pptOpenXml.GetSlideComments(slideNumber, false);
			if (showComments && !String.IsNullOrEmpty(slideComments))
			{
				note.AppendPageContent(pageId, commentsStr + ": \n\n" + slideComments, (int)note.GetPageWidth(pageId));
			}

			// Add notes
			string slideNotes = pptOpenXml.GetSlideNotes(slideNumber);
			if (showNotes && !String.IsNullOrEmpty(slideNotes))
			{
				note.AppendPageContent(pageId, notesStr + ": \n\n" + slideNotes, (int)note.GetPageWidth(pageId));
			}

			// remove the author
			note.RemoveAuthor(pageId);

			return pageId;
		}
		#endregion

#if INDESIGN_INSTALLED
		#region InDesign Helper Methods
		/// <summary>
		/// Sets the export options of the indesign document to HTML
		/// </summary>
		/// <param name="doc"></param>
		protected void SetInDesignHtmlExportOptions(InDesignDocument doc)
		{
			// General export options
			doc.HTMLExportPreferences.ExportSelection = false;
			doc.HTMLExportPreferences.ExportOrder = idExportOrder.idLayoutOrder;
			doc.HTMLExportPreferences.BulletExportOption = idBulletListExportOption.idUnorderedList;
			doc.HTMLExportPreferences.NumberedListExportOption = idNumberedListExportOption.idOrderedList;
			doc.HTMLExportPreferences.ViewDocumentAfterExport = true;

			// Image export options
			doc.HTMLExportPreferences.ImageExportOption = idImageExportOption.idOptimizedImage;
			doc.HTMLExportPreferences.PreserveLayoutAppearence = true;
			doc.HTMLExportPreferences.ImageExportResolution = idImageResolution.idPpi150;
			doc.HTMLExportPreferences.CustomImageSizeOption = idImageSizeOption.idSizeFixed;
			doc.HTMLExportPreferences.ImageAlignment = idImageAlignmentType.idAlignCenter;
			doc.HTMLExportPreferences.ImageSpaceBefore = 0;
			doc.HTMLExportPreferences.ImageSpaceAfter = 0;
			//doc.HTMLExportPreferences.ApplyImageAlignmentToAnchoredObjectSettings = true;
			doc.HTMLExportPreferences.ImageConversion = idImageConversion.idJPEG;
			doc.HTMLExportPreferences.JPEGOptionsQuality = idJPEGOptionsQuality.idMedium;
			doc.HTMLExportPreferences.JPEGOptionsFormat = idJPEGOptionsFormat.idProgressiveEncoding;
			doc.HTMLExportPreferences.IgnoreObjectConversionSettings = false;

			// Advanced export options
			//doc.HTMLExportPreferences.CSSExportOption = idStyleSheetExportOption.idEmbeddedCSS;
			//doc.HTMLExportPreferences.IncludeCSSDefinition = true;
			doc.HTMLExportPreferences.PreserveLocalOverride = false;
		}

		/// <summary>
		/// change the links to have the full path
		/// </summary>
		/// <param name="htmlDoc"></param>
		/// <param name="htmlDir"></param>
		protected void ModifyInDesignHtmlLinks(HtmlDocument htmlDoc, string htmlDir)
		{
			HtmlNodeCollection links = htmlDoc.DocumentNode.SelectNodes("//*[@background or @lowsrc or @src or @href]");
			foreach (HtmlNode link in links)
			{
				if (link.Attributes["background"] != null)
				{
					link.Attributes["background"].Value = htmlDir + "\\" + link.Attributes["background"].Value;
				}
				if (link.Attributes["lowsrc"] != null)
				{
					link.Attributes["lowsrc"].Value = htmlDir + "\\" + link.Attributes["lowsrc"].Value;
				}
				if (link.Attributes["src"] != null)
				{
					link.Attributes["src"].Value = htmlDir + "\\" + link.Attributes["src"].Value;
				}
				if (link.Attributes["href"] != null)
				{
					link.Attributes["href"].Value = htmlDir + "\\" + link.Attributes["href"].Value;
				}
			}
		}

		/// <summary>
		/// Get the title and remove its elements from the document
		/// </summary>
		/// <param name="htmlDoc"></param>
		/// <returns></returns>
		protected string GetInDesignTitle(HtmlDocument htmlDoc)
		{
			HtmlNode titleInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//*[contains(@class,'title-info')]");
			HtmlNodeCollection titleNodes = htmlDoc.DocumentNode.SelectNodes("//*[contains(@class,'main-title')]");
			string title = String.Empty;
			if (titleInfoNode != null)
			{
				title += titleInfoNode.InnerText;
				titleInfoNode.ParentNode.Remove();
			}
			if (titleNodes != null)
			{
				foreach (var titleNode in titleNodes)
				{
					title += titleNode.InnerText + " ";
					titleNode.ParentNode.Remove();
				}
			}
			return title.Trim();
		}
		#endregion
#endif

		#region ePub Helper Methods
		/// <summary>
		/// Replace the url links in the epub Html pages to links for the OneNote Pages
		/// </summary>
		/// <param name="pagesHtml">Html Documents for the epub pages</param>
		/// <param name="pagePaths">Paths of the epub html pages</param>
		/// <param name="pageLinks">Links to OneNote pages the coorspond to the Epub pages</param>
		private void ReplaceEpubLinksFromHtmlToOneNote(List<HtmlDocument> pagesHtml, List<string> pagePaths,
			List<string> pageLinks)
		{
			foreach (var pageHtml in pagesHtml)
			{
				HtmlNodeCollection links = pageHtml.DocumentNode.SelectNodes("//*[@href]");
				if (links == null)
				{
					continue;
				}

				foreach (var link in links)
				{

					if (link.Attributes["href"] == null || !pagePaths.Contains(link.Attributes["href"].Value))
					{
						continue;
					}

					int index = pagePaths.IndexOf(link.Attributes["href"].Value);
					link.Attributes["href"].Value = pageLinks[index];
				}
			}
		}
		#endregion
	}
}
