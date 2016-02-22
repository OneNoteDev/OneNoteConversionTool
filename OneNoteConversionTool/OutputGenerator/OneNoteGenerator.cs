using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteConversionTool.OutputGenerator
{
	/// <summary>
	/// Handles the interaction with the OneNote document output, including creating notebook, section, page, adding content, and other things
	/// </summary>
	public class OneNoteGenerator
	{
		#region Global Variables
		private const string NS = "http://schemas.microsoft.com/office/onenote/2013/onenote";
		private const string WordOutline = @"<one:Position x=""36.0"" y=""{0}"" z=""2"" /><one:Size width=""{2}"" height=""14.00""/><one:OEChildren><one:OE style=""font-family:Segoe;font-size:12.0pt"" alignment=""left""><one:T><![CDATA[{1}]]></one:T></one:OE></one:OEChildren>";
		private const string HtmlOutline = @"<one:Position x=""36.0"" y=""{0}"" z=""2"" /><one:Size width=""{1}"" height=""14.00""/><one:OEChildren><one:HTMLBlock><one:Data><![CDATA[{2}]]></one:Data></one:HTMLBlock></one:OEChildren>";
		private const string XmlImageContent = "<one:Position x=\"36.0\" y=\"{3}\" z=\"2\" /><one:Size width=\"{1}\" height=\"{2}\" isSetByUser=\"true\" /><one:Data>{0}</one:Data>";

		private const string IsCollapsedAttrKey = "isCollapsed";
		private const string PageLevelAttrKey = "pageLevel";
		private const string ShowDateAttrKey = "showDate";
		private const string ShowTimeAttrKey = "showTime";

		private const int ContentBlockMargin = 20;
		private const int MaxPageWidth = 960;

		private readonly Application _mApp;
		private readonly string _mOutputPath;
		#endregion

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="mOutputDir"></param>
		public OneNoteGenerator(string mOutputDir)
		{
			_mApp = new Application();
			_mOutputPath = mOutputDir;
		}

		#region Scope Hierarchy
		/// <summary>
		/// Returns the children scope hierarchy of the OneNote object in XML format
		/// Gets the immediate child nodes of the start node, and no descendants in higher or lower subsection groups.
		/// </summary>
		/// <param name="objectId"></param>
		/// <returns></returns>
		public string GetChildrenScopeHierarchy(string objectId)
		{
			string xml;
			_mApp.GetHierarchy(objectId, HierarchyScope.hsChildren, out xml);
			return xml;
		}

		/// <summary>
		/// Returns the page scope hierarchy of the onenote object(page\section\notebook) in XML format
		/// Gets all pages below the start node, including all pages in section groups and subsection groups.
		/// </summary>
		/// <param name="objectId"></param>
		/// <returns></returns>
		public string GetPageScopeHierarchy(string objectId)
		{
			string xml;
			_mApp.GetHierarchy(objectId, HierarchyScope.hsPages, out xml);
			return xml;
		}

		/// <summary>
		/// Returns the section scope hierarchy of the onenote object(section\notebook) in XML format
		/// Gets all sections below the start node, including sections in section groups and subsection groups.
		/// </summary>
		/// <param name="objectId"></param>
		/// <returns></returns>
		public string GetSectionScopeHierarchy(string objectId)
		{
			string xml;
			_mApp.GetHierarchy(objectId, HierarchyScope.hsSections, out xml);
			return xml;
		}
		#endregion

		#region Creation
		/// <summary>
		/// Creates a new OneNote notebook with the given name
		/// </summary>
		/// <param name="notebookName"></param>
		/// <returns></returns>
		public string CreateNotebook(string notebookName)
		{
			string notebookId;

			try
			{
				_mApp.OpenHierarchy(Path.Combine(_mOutputPath, notebookName), String.Empty, out notebookId, CreateFileType.cftNotebook);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in CreateNotebook: " + e.Message, e);
			}

			return notebookId;
		}

		/// <summary>
		/// Creates one note section under the given notebook ID
		/// </summary>
		/// <param name="sectionName"></param>
		/// <param name="notebookId"></param>
		/// <returns></returns>
		public string CreateSection(string sectionName, string notebookId)
		{
			string sectionId;

			try
			{
				//read notebook xml to get all the existed section names
				string docXml;
				_mApp.GetHierarchy(notebookId, HierarchyScope.hsSections, out docXml);
				var xDoc = XDocument.Parse(docXml);
				XNamespace xNs = xDoc.Root.Name.Namespace;
				var sectionNames = xDoc.Root.Elements(xNs + "Section").Select(section => section.Attribute("name").Value);

				//if sectionName already exist, set sectionName as sectionName (2), and etc..
				int ordinal = 2;
				var adjustedSectionName = sectionName;
				while (sectionNames.Contains(adjustedSectionName))
				{
					adjustedSectionName = sectionName + " (" + ordinal + ")";
					++ordinal;
				}

				_mApp.OpenHierarchy(adjustedSectionName + ".one", notebookId, out sectionId, CreateFileType.cftSection);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in CreateSection: " + e.Message, e);
			}
			return sectionId;
		}

		/// <summary>
		/// Creates a one note page under the given section ID
		/// </summary>
		/// <param name="pageName"></param>
		/// <param name="sectionId"></param>
		/// <returns></returns>
		public string CreatePage(string pageName, string sectionId)
		{
			string pageId;
			try
			{
				_mApp.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);
				
				// Get the title and set it to our page name
				string xml;
				_mApp.GetPageContent(pageId, out xml, PageInfo.piAll);
				var doc = XDocument.Parse(xml);
				var ns = doc.Root.Name.Namespace;
				var title = doc.Descendants(ns + "T").First();
				title.Value = pageName;

				// Update the page
				_mApp.UpdatePageContent(doc.ToString());

			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in CreatePage: " + e.Message, e);
			}
			return pageId;
		}

		/// <summary>
		/// Generates a Table of Content Page of all the pages underneath the given section
		/// Also sets the Table of Content Page to be the first page in the given section
		/// </summary>
		/// <param name="sectionId"></param>
		/// <param name="tocPageTitle"></param>
		/// <returns></returns>
		public string CreateTableOfContentPage(string sectionId, string tocPageTitle = "Table of Contents")
		{
			string retVal = CreatePage(tocPageTitle, sectionId);

			var sectionHierarchy = GetChildrenScopeHierarchy(sectionId);

			XDocument xdoc = XDocument.Parse(sectionHierarchy);
			XNamespace xNs = xdoc.Root.Name.Namespace;
			var sectionElement = xdoc.Elements(xNs + "Section")
				.Single(x => x.Attribute("ID").Value.Equals(sectionId));

			var tocContent = string.Empty;

			var pageElements = sectionElement.Elements(xNs + "Page");

			//Iterate through the sections and get hyperlinks for them to be put into the new page.
			foreach (var pageElement in pageElements)
			{
				var pageId = pageElement.Attribute("ID").Value;

				//Don't add a link to the ToC page itself
				if (pageId != retVal)
				{
					var pageTitle = pageElement.Attribute("name").Value;
					pageTitle = Encoding.Default.GetString(Encoding.UTF8.GetBytes(pageTitle));
					string hyperLink;

					_mApp.GetHyperlinkToObject(pageId, string.Empty, out hyperLink);

					var pageLevel = int.Parse(pageElement.Attribute("pageLevel").Value);

					while (pageLevel > 1)
					{
						tocContent += "\t";
						--pageLevel;
					}

					tocContent += string.Format("<a href=\"{0}\">{1}</a>", hyperLink, pageTitle) + "\n\n";
				}
			}

			AddPageContent(retVal, tocContent);

			//Move the ToC page to the first page
			SetAsFirstPage(retVal, sectionId);

			return retVal;
		}
		#endregion

		#region Modify Page
		/// <summary>
		/// Sets the page attribute to the given value
		/// </summary>
		/// <param name="pageId">ID of the page</param>
		/// <param name="attr">The attribute of interest</param>
		/// <param name="value">The new value of the attribute</param>
		public void SetPageAttribute(string pageId, string attr, string value)
		{
			try
			{
				// get the page
				XmlDocument doc = GetPageContent(pageId);
				XmlNode page = doc.DocumentElement;
				if (page == null || page.Attributes == null) return;
				if (page.Attributes[attr] == null)
				{
					XmlAttribute xAttr = doc.CreateAttribute(attr);
					xAttr.Value = value;
					page.Attributes.Append(xAttr);
				}
				else
				{
					page.Attributes[attr].Value = value;
				}

				// update the page
				_mApp.UpdatePageContent(doc.OuterXml);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in SetPageAttribute: " + e.Message, e);
			}
		}

		/// <summary>
		/// Sets the page title attribute to the given value
		/// </summary>
		/// <param name="pageId">ID of the page</param>
		/// <param name="attr">The attribute of interest for one:Title node</param>
		/// <param name="value">The new value of the attribute</param>
		public void SetPageTitleAttribute(string pageId, string attr, string value)
		{
			try
			{
				// get the page
				XmlDocument doc = GetPageContent(pageId);
				XmlNode page = doc.DocumentElement;
				if (page == null) return;
				XmlNode title = page.SelectSingleNode("//one:Title", GetNSManager(doc.NameTable));
				if (title == null || title.Attributes == null) return;

				if (title.Attributes[attr] == null)
				{
					XmlAttribute xAttr = doc.CreateAttribute(attr);
					xAttr.Value = value;
					title.Attributes.Append(xAttr);
				}
				else
				{
					title.Attributes[attr].Value = value;
				}

				// update the page
				_mApp.UpdatePageContent(doc.OuterXml);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in SetPageTitleAttribute", e);
			}
		}

		/// <summary>
		/// Make the page as subpage (if isSet true) or promote it (if isSet if false)
		/// </summary>
		/// <param name="sectionId"></param>
		/// <param name="pageId"></param>
		/// <param name="isSet">defaults to true, if true, increment pageLevel, else decrement pageLevel</param>
		public void SetSubPage(string sectionId, string pageId, bool isSet = true)
		{
			try
			{
				string hierarchy = string.Empty;
				_mApp.GetHierarchy(sectionId, HierarchyScope.hsPages, out hierarchy);

				XmlDocument doc = new XmlDocument();
				doc.LoadXml(hierarchy);

				string xpath = string.Format("//one:Page[@ID='{0}']", pageId);

				XmlNode page = doc.SelectSingleNode(xpath, GetNSManager(doc.NameTable));
				var pageLevel = int.Parse(page.Attributes["pageLevel"].Value);


				if (isSet)
				{
					++pageLevel;
				}
				else
				{
					--pageLevel;
					pageLevel = pageLevel > 0 ? pageLevel : 1;
				}

				page.Attributes["pageLevel"].Value = pageLevel.ToString();
				_mApp.UpdateHierarchy(doc.OuterXml);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in SetAsSubPage: " + e.Message, e);
			}
		}

		/// <summary>
		/// Sets the page level of a given page (note, page level must be {1, 2, or 3})
		/// </summary>
		/// <param name="pageId"></param>
		/// <param name="pageLevel"></param>
		public void SetPageLevel(string pageId, int pageLevel)
		{
			// ensure that page level is within acceptable range = [1, 3]
			pageLevel = pageLevel < 1 ? 1 : pageLevel;
			pageLevel = pageLevel > 3 ? 3 : pageLevel;

			SetPageAttribute(pageId, PageLevelAttrKey, pageLevel.ToString());
		}

		/// <summary>
		/// Sets the collapse attribute of the page to hide or unhide the subpages under the given page
		/// </summary>
		/// <param name="pageId">ID of the page</param>
		/// <param name="isCollapsed">isCollapsed attribute value</param>
		public void SetCollapsePage(string pageId, bool isCollapsed = true)
		{
			SetPageAttribute(pageId, IsCollapsedAttrKey, isCollapsed.ToString().ToLower());
		}

		/// <summary>
		/// Sets whether the date should be shown or not in the page
		/// </summary>
		/// <param name="pageId"></param>
		/// <param name="isShown"></param>
		public void SetShowDate(string pageId, bool isShown = true)
		{
			SetPageTitleAttribute(pageId, ShowDateAttrKey, isShown.ToString().ToLower());
		}

		/// <summary>
		/// Sets whether the time should be shown or not in the page
		/// </summary>
		/// <param name="pageId"></param>
		/// <param name="isShown"></param>
		public void SetShowTime(string pageId, bool isShown = true)
		{
			SetPageTitleAttribute(pageId, ShowTimeAttrKey, isShown.ToString().ToLower());
		}

		/// <summary>
		/// Adds the content to the corresponding page.
		/// </summary>
		/// <param name="pageId">Page Identifier in which the content to be displayed.</param>
		/// <param name="content">The content which needs to be added to the page.</param>
		/// <param name="yPos">Starting vertical position of the block in the page in Pixels from the top of the page</param>
		/// <param name="width">Maximum width of the outline block where the content is added</param>
		public void AddPageContent(string pageId, string content, int yPos = 80, int width = 520)
		{
			var doc = GetPageContent(pageId);
			var page = doc.SelectSingleNode("//one:Page", GetNSManager(doc.NameTable));
			if (page == null)
				return;

			var childOutline = doc.CreateElement("one:Outline", NS);
			
			try
			{
				childOutline.InnerText = string.Format(WordOutline, yPos, content, width);
				

				page.AppendChild(childOutline);
				string childContent = Utility.UnescapeXml(childOutline.InnerXml);
				string newPageContent = doc.InnerXml.Replace(childOutline.InnerXml, childContent);
				_mApp.UpdatePageContent(newPageContent, DateTime.MinValue);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in AddPageContent: " + e.Message, e);
			}
		}

		/// <summary>
		/// Adds the content as a HTML block.
		/// </summary>
		/// <param name="pageId">Id of the page</param>
		/// <param name="content">content to be added</param>
		/// <param name="yPos">Starting position of the block</param>
		/// <param name="width">Width of the block</param>
		public void AddPageContentAsHtmlBlock(string pageId, string content, int yPos = 80, int width = 520)
		{
			var doc = GetPageContent(pageId);
			var page = doc.SelectSingleNode("//one:Page", GetNSManager(doc.NameTable));
			if (page == null)
				return;

			try
			{
				//If outline doesn't exist, create one
				if (doc.SelectSingleNode("//one:Page/one:Outline", GetNSManager(doc.NameTable)) == null)
				{
					XmlNode childOutline = doc.CreateElement("one:Outline", NS);

					childOutline.InnerText = string.Format(HtmlOutline, yPos, width, content);

					page.AppendChild(childOutline);
					string childContent = Utility.UnescapeXml(childOutline.InnerXml);
					string newPageContent = doc.InnerXml.Replace(childOutline.InnerXml, childContent);

					//Override the content
					_mApp.UpdatePageContent(newPageContent, DateTime.MinValue);
				}
				else
				{
					UpdatePageContent(pageId, content);
				}
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in AddPageContent: " + e.Message, e);
			}
		}

		/// <summary>
		/// Adds an image to the page
		/// The width of the page is at most MaxPageWidth (960 pixels)
		/// If the width is bigger, the image is proportionally minimized to MaxPageWidth
		/// </summary>
		/// <param name="pageId"></param>
		/// <param name="img"></param>
		/// <param name="yPos"></param>
		public void AddImageToPage(string pageId, Image img, int yPos = 80)
		{
			if (img == null)
				return;

			Size size = new Size(img.Width, img.Height);
			if (img.Width > MaxPageWidth)
			{
				size.Height = (size.Height*MaxPageWidth)/(size.Width);
				size.Width = MaxPageWidth;
			}

			// convert the image
			Bitmap bitmap = new Bitmap(img, size);
			MemoryStream stream = new MemoryStream();
			bitmap.Save(stream, img.RawFormat);
			string imgString = Convert.ToBase64String(stream.ToArray());

			// get the image xml
			string imgXmlStr = String.Format(XmlImageContent, imgString, bitmap.Width, bitmap.Height, yPos);

			// get the page
			XmlDocument doc = GetPageContent(pageId);
			XmlNode page = doc.DocumentElement;
			if (page == null) return;
			XmlNode imgNode = doc.CreateNode(XmlNodeType.Element, "one:Image", NS);
			page.AppendChild(imgNode);
			imgNode.InnerXml = imgXmlStr;

			// update the page adding the image
			_mApp.UpdatePageContent(doc.OuterXml);
		}

		/// <summary>
		/// Adds the content at the end of the corresponding page.
		/// </summary>
		/// <param name="pageId">Id of the page</param>
		/// <param name="content">content to be added</param>
		/// <param name="width">Width of the block</param>
		public void AppendPageContent(string pageId, string content, int width = 520)
		{
			AddPageContent(pageId, content, (int) GetPageHeight(pageId) + ContentBlockMargin, width);
		}

		/// <summary>
		/// Adds the content as HTML block at the end of the corresponding page
		/// </summary>
		/// <param name="pageId">ID of the page</param>
		/// <param name="content">HTML block to be added</param>
		/// <param name="width">Width of the block</param>
		public void AppendPageContentAsHtmlBlock(string pageId, string content, int width = 520)
		{
			AddPageContentAsHtmlBlock(pageId, content, (int) GetPageHeight(pageId) + ContentBlockMargin, width);
		}

		/// <summary>
		/// Adds the image at the end of the corresponding page
		/// </summary>
		/// <param name="pageId">ID of the page</param>
		/// <param name="img">The Image to be added</param>
		public void AppendImageToPage(string pageId, Image img)
		{
			AddImageToPage(pageId, img, (int)GetPageHeight(pageId) + ContentBlockMargin);
		}

		/// <summary>
		/// UpdatePageContent
		/// this method is invoked when a previous page word section
		/// needs to be updated
		/// </summary>
		/// <param name="pageId"></param>
		/// <param name="content"></param>
		public void UpdatePageContent(string pageId, string content)
		{
			var doc = GetPageContent(pageId);
			var children = doc.SelectSingleNode("//one:Page/one:Outline[position() = last()]/one:OEChildren", GetNSManager(doc.NameTable));
			var htmBlock = doc.CreateElement("one:HTMLBlock", NS);

			try
			{
				htmBlock.InnerText = string.Format(@"<one:Data><![CDATA[{0}]]></one:Data>", content);
				children.AppendChild(htmBlock);

				string newPageContent = doc.InnerXml.Replace(htmBlock.InnerXml, Utility.UnescapeXml(htmBlock.InnerXml));
				_mApp.UpdatePageContent(newPageContent, DateTime.MinValue);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in UpdatePageContent: " + e.Message, e);
			}
		}

		/// <summary>
		/// set a page as the first page in the section
		/// </summary>
		/// <param name="pageId"></param>
		/// <param name="sectionId"></param>
		public void SetAsFirstPage(string pageId, string sectionId)
		{
			var sectionXml = GetChildrenScopeHierarchy(sectionId);
			var xDoc = XDocument.Parse(sectionXml);
			XNamespace xNs = xDoc.Root.Name.Namespace;
			var sectionElement = xDoc.Elements(xNs + "Section")
				.Single(x => x.Attribute("ID").Value.Equals(sectionId));

			var pageElements = sectionElement.Elements(xNs + "Page");
			if (pageElements.Count() > 1)
			{
				var page = pageElements.Single(x => x.Attribute("ID").Value.Equals(pageId));
				page.Remove();

				sectionElement.Elements(xNs + "Page").FirstOrDefault().AddBeforeSelf(page);
				_mApp.UpdateHierarchy(xDoc.ToString());
			}
		}

		/// <summary>
		/// Removes the author from any node in the page
		/// </summary>
		/// <param name="pageId"></param>
		public void RemoveAuthor(string pageId)
		{
			var doc = GetPageContent(pageId);
			var nodes = doc.SelectNodes("//*[@author or @authorInitials or @lastModifiedBy or @lastModifiedByInitials]", GetNSManager(doc.NameTable));
			
			if (nodes == null)
				return;

			foreach (XmlNode node in nodes)
			{
				if (node.Attributes == null)
					continue;

				if (node.Attributes["author"] != null)
					node.Attributes["author"].Value = String.Empty;
				if (node.Attributes["authorInitials"] != null)
					node.Attributes["authorInitials"].Value = String.Empty;
				if (node.Attributes["lastModifiedBy"] != null)
					node.Attributes["lastModifiedBy"].Value = String.Empty;
				if (node.Attributes["lastModifiedByInitials"] != null)
					node.Attributes["lastModifiedByInitials"].Value = String.Empty;
			}

			// update the page
			_mApp.UpdatePageContent(doc.OuterXml);
		}
		#endregion

		#region Get
		/// <summary>
		/// Gets the ID of an existing notebook in the current directory
		/// </summary>
		/// <param name="notebookName">name of the notebook of interest</param>
		/// <returns>ID of the notebook of interest</returns>
		public string GetNotebook(string notebookName)
		{
			string notebookId;

			try
			{
				_mApp.OpenHierarchy(Path.Combine(_mOutputPath, notebookName), String.Empty, out notebookId);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in GetNotebook: " + e.Message, e);
			}

			return notebookId;
		}

		/// <summary>
		/// Gets the ID of an existing section in the given notebook
		/// If more than one exists, it returns the first one
		/// </summary>
		/// <param name="sectionName">section name</param>
		/// <param name="notebookId">ID of the notebook where the section exists</param>
		/// <returns></returns>
		public string GetSection(string sectionName, string notebookId)
		{
			string sectionId;

			try
			{
				_mApp.OpenHierarchy(sectionName + ".one", notebookId, out sectionId);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in GetSection: " + e.Message, e);
			}

			return sectionId;
		}

		/// <summary>
		/// Gets the ID o fan existing page in the given section
		/// If more than one exists, it returns the first one
		/// </summary>
		/// <param name="pageName">page name</param>
		/// <param name="sectionId">ID of the section where the page exists</param>
		/// <returns></returns>
		public string GetPage(string pageName, string sectionId)
		{
			var pageId = String.Empty;

			try
			{
				XmlDocument xDoc = new XmlDocument();
				xDoc.LoadXml(GetPageScopeHierarchy(sectionId));
				XmlNode node = xDoc.SelectSingleNode(String.Format("//one:Page[@name='{0}']", pageName), GetNSManager(xDoc.NameTable));
				if (node != null && node.Attributes != null) 
					pageId = node.Attributes["ID"].Value;
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in GetPage: " + e.Message, e);
			}

			return pageId;
		}

		/// <summary>
		/// GetPageContent
		/// return DOM representation of onenote's page
		/// </summary>
		/// <param name="pageId"></param>
		/// <returns></returns>
		public XmlDocument GetPageContent(string pageId)
		{
			var doc = new XmlDocument();
			try
			{
				string pageInfo;
				_mApp.GetPageContent(pageId, out pageInfo, PageInfo.piAll);
				doc.LoadXml(pageInfo);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in GetPageContent: " + e.Message + pageId.ToString(), e);
			}
			return doc;
		}

		/// <summary>
		/// get the hyper link to the object
		/// </summary>
		/// <param name="objectId"></param>
		/// <returns></returns>
		public string GetHyperLinkToObject(string objectId)
		{
			string hyperLink;
			_mApp.GetHyperlinkToObject(objectId, String.Empty, out hyperLink);
			return hyperLink;
		}

		/// <summary>
		/// Get the height of the page
		/// </summary>
		/// <param name="pageId">id of the page</param>
		/// <returns>height of the page (default is 80.0 if there are no elements in the page)</returns>
		public double GetPageHeight(string pageId)
		{
			double retVal = 80.0;

			XmlDocument xmlPageDoc = GetPageContent(pageId);
			XmlNodeList nodeList = xmlPageDoc.SelectNodes("//*[one:Position]", GetNSManager(xmlPageDoc.NameTable));
			if (nodeList != null)
			{
				foreach (XmlNode node in nodeList)
				{
					XmlElement xmlPosition = node["one:Position"];
					XmlElement xmlSize = node["one:Size"];
					if (xmlPosition == null || xmlSize == null) continue;

					XmlAttribute xmlY = xmlPosition.Attributes["y"];
					XmlAttribute xmlHeight = xmlSize.Attributes["height"];

					double y = xmlY == null ? 0 : Double.Parse(xmlY.Value);
					double height = xmlHeight == null ? 0 : Double.Parse(xmlHeight.Value);

					retVal = Math.Max(retVal, y + height);
				}
			}

			return retVal;
		}

		/// <summary>
		/// Get the height of the page
		/// </summary>
		/// <param name="pageId">id of the page</param>
		/// <returns>height of the page (default is 520.0 if there are no elements in the page)</returns>
		public double GetPageWidth(string pageId)
		{
			double retVal = 520.0;

			XmlDocument xmlPageDoc = GetPageContent(pageId);
			XmlNodeList nodeList = xmlPageDoc.SelectNodes("//*[one:Position]", GetNSManager(xmlPageDoc.NameTable));
			if (nodeList != null)
			{
				foreach (XmlNode node in nodeList)
				{
					XmlElement xmlPosition = node["one:Position"];
					XmlElement xmlSize = node["one:Size"];
					if (xmlPosition == null || xmlSize == null) continue;

					XmlAttribute xmlWidth = xmlSize.Attributes["width"];

					double width = xmlWidth == null ? 0 : Double.Parse(xmlWidth.Value);

					retVal = Math.Max(retVal, width);
				}
			}

			return retVal;
		}

		/// <summary>
		/// Gets the value of an attribute of the page
		/// </summary>
		/// <param name="pageId">ID of the page</param>
		/// <param name="attr">the attribute of interest</param>
		/// <returns>the value of the attribute</returns>
		public string GetPageAttribute(string pageId, string attr)
		{
			try
			{
				string hierarchy;
				string sectionId;
				_mApp.GetHierarchyParent(pageId, out sectionId);
				_mApp.GetHierarchy(sectionId, HierarchyScope.hsPages, out hierarchy);

				var doc = new XmlDocument();
				doc.LoadXml(hierarchy);

				string xpath = String.Format("//one:Page[@ID='{0}']", pageId);
				XmlNode page = doc.SelectSingleNode(xpath, GetNSManager(doc.NameTable));

				if (page == null || page.Attributes == null || page.Attributes[attr] == null) return String.Empty;

				return page.Attributes[attr].Value;
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in GetPageAttribute: " + e.Message, e);
			}
		}
		#endregion

		#region Private Methods
		/// <summary>
		/// Returns the namespace manager from the passed xml name table.
		/// </summary>
		/// <param name="nameTable">Name table of the xml document.</param>
		/// <returns>Returns the namespace manager.</returns>
		private XmlNamespaceManager GetNSManager(XmlNameTable nameTable)
		{
			var nsManager = new XmlNamespaceManager(nameTable);
			try
			{
				nsManager.AddNamespace("one", NS);
			}
			catch (Exception e)
			{
				throw new ApplicationException("Error in GetNSManager: " + e.Message, e);
			}
			return nsManager;
		}
		#endregion
	}
}
