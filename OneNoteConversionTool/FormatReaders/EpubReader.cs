using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using eBdb.EpubReader;
using HtmlAgilityPack;
using Ionic.Zip;

namespace OneNoteConversionTool.FormatReaders
{
	/// <summary>
	/// Obtains data from epub file
	/// </summary>
	public class EpubReader
	{
		#region Global Variables
		private const string MathMlNameSpace = "http://www.w3.org/1998/Math/MathML";
		private const string MathMlOutline = "<!--[if mathML]>{0}<![endif]-->";

		private const double MaxWidth = 960;

		private readonly Epub _mEpub;
		private readonly string _mFilePath;
		private readonly string _mEpubDir;
		private readonly string _mContentDir;
		#endregion

		public EpubReader(string filePath, string outputDir)
		{
			try
			{
				_mEpub = new Epub(filePath);
			}
			catch (Exception e)
			{
				_mEpub = null;
                Console.WriteLine("Error in EpubReader: _mEpub couldn't be initialized.");
                Console.WriteLine(e.Message);
			}
			_mFilePath = filePath;

			string epubExtractDir = Directory.Exists(Path.Combine(outputDir, "Epub Extracted Files"))
				? Path.Combine(outputDir, "Epub Extracted Files")
				: Utility.CreateDirectory(Path.Combine(outputDir, "Epub Extracted Files"));
			_mEpubDir = Utility.CreateDirectory(Utility.NewFolderPath(epubExtractDir, Path.GetFileNameWithoutExtension(filePath)));
			ZipFile.Read(filePath).ExtractAll(_mEpubDir);

			_mContentDir = Path.GetDirectoryName(GetOpfFilePath());
		}

		#region Public Methods
		/// <summary>
		/// Gets the title of the epub file
		/// </summary>
		/// <returns></returns>
		public string GetTitle()
		{
			return _mEpub != null && _mEpub.Title != null && _mEpub.Title.Count > 0 
				? _mEpub.Title[0].Trim()
				: Path.GetFileNameWithoutExtension(_mFilePath);
		}

		/// <summary>
		/// Gets the title of each page
		/// </summary>
		/// <returns></returns>
		public List<string> GetPageTitles()
		{
			var pageTitles = new List<string>();

			// Get the title of each page
			foreach (var pagePath in GetPagePaths())
			{
				var doc = new HtmlDocument();
				doc.Load(pagePath);

				HtmlNode titleNode = doc.DocumentNode.SelectSingleNode("//h1") ?? doc.DocumentNode.SelectSingleNode("//title");

				pageTitles.Add(titleNode != null ? Utility.UnescapeXml(titleNode.InnerText.Trim()) : String.Empty);
			}

			return pageTitles;
		}

		/// <summary>
		/// Get the file paths of all pages
		/// </summary>
		/// <returns></returns>
		public List<string> GetPagePaths()
		{
			// Get the paths of each page using _mEpub if it is not null
			if (_mEpub != null)
			{
				return (from string htmlPath in _mEpub.Content.Keys
					select Path.GetFullPath(Path.Combine(_mContentDir, htmlPath))).ToList();
			}

			var pagePaths = new List<string>();

			// Get the opf path and ensure that it exists
			string opfPath = GetOpfFilePath();
			if (opfPath.Equals(String.Empty))
			{
				return pagePaths;
			}

			// Load the opf file
			var xOpfDoc = new XmlDocument();
			xOpfDoc.Load(opfPath);

			// Get the spine data
			XmlNodeList xNodes = xOpfDoc.SelectNodes("//*");
			if (xNodes == null)
			{
				return pagePaths;
			}

			List<XmlNode> xItemRefNodes = (from XmlNode xNode in xNodes 
										   where xNode.Name.Contains("itemref") 
										   select xNode).ToList();
			foreach (XmlNode xItemRefNode in xItemRefNodes)
			{
				if (xItemRefNode.Attributes == null || xItemRefNode.Attributes["idref"] == null)
				{
					pagePaths.Add(String.Empty);
					continue;
				}
				string idref = xItemRefNode.Attributes["idref"].Value;
				string xPath = String.Format("//*[@id=\"{0}\"]", idref);
				XmlNode node = xOpfDoc.SelectSingleNode(xPath);
				
				pagePaths.Add(node == null || node.Attributes == null || node.Attributes["href"] == null
					? String.Empty 
					: Path.GetFullPath(Path.Combine(_mContentDir, node.Attributes["href"].Value)));
			}

			return pagePaths;
		}

		/// <summary>
		/// Gets the html code for each page
		/// </summary>
		/// <returns></returns>
		public List<HtmlDocument> GetPagesAsHtmlDocuments()
		{
			var pagesAsHtml = new List<HtmlDocument>();

			// Get the pages as html
			foreach (string pagePath in GetPagePaths())
			{
				var htmlDoc = new HtmlDocument();
				htmlDoc.Load(pagePath);

				// Remove Title
				HtmlNode titleNode = htmlDoc.DocumentNode.SelectSingleNode("//h1");
				if (titleNode != null)
				{
					titleNode.Remove();
				}

				// Remove <![CDATA[]]> from Script and style nodes
				HtmlNodeCollection scriptAndStyleNodes = htmlDoc.DocumentNode.SelectNodes("//script | //style");
				if (scriptAndStyleNodes != null)
				{
					foreach (var node in scriptAndStyleNodes)
					{
						node.InnerHtml = node.InnerHtml.Replace("<![CDATA[", "").Replace("]]>", "");
					}
				}

				// Convert MathML
				ConvertMathMl(htmlDoc);

				// Replace relative paths in links with full paths
				UseFullPathForLinks(htmlDoc, pagePath);

				// Ensure that images have their full dimenssions (up to max width of 960 px)
				AddImageDimenssionsToHtml(htmlDoc);

				pagesAsHtml.Add(htmlDoc);
			}

			return pagesAsHtml;
		}

		/// <summary>
		/// Gets the level of each page as dictionary where the key in the dictionary is the
		///		full path of the xhtml page
		/// </summary>
		/// <returns></returns>
		public Dictionary<string, int> GetPagesLevel()
		{
			var pagesLevel = new Dictionary<string, int>();

			// Gets the opf package document
			var opfDoc = new XmlDocument();
			opfDoc.Load(GetOpfFilePath());
			if (opfDoc.DocumentElement == null)
			{
				return pagesLevel;
			}

			// Gets the Navigation page
			XmlNode xmlNavNode = opfDoc.DocumentElement.SelectSingleNode("//*[@properties=\"nav\"]");
			if (xmlNavNode == null || xmlNavNode.Attributes == null || xmlNavNode.Attributes["href"] == null)
			{
				return pagesLevel;
			}

			// Load the navigation page as HtmlDocument
			var navDoc = new HtmlDocument();
			string navPath = Path.Combine(_mContentDir, xmlNavNode.Attributes["href"].Value);
			navDoc.Load(navPath);

			// Gets the directory of all the xhtml files
			string htmlDir = Path.GetDirectoryName(navPath);
			if (htmlDir == null)
			{
				return pagesLevel;
			}

			// Gets the page levels
			HtmlNodeCollection nodes = navDoc.DocumentNode.SelectNodes("//*[@href]");
			foreach (var node in nodes)
			{
				if (!File.Exists(Path.Combine(htmlDir, node.Attributes["href"].Value)))
				{
					continue;
				}
				string filePath = Path.GetFullPath(Path.Combine(htmlDir, node.Attributes["href"].Value));
				int pageLevel = node.AncestorsAndSelf().Count(e => e.Name.Equals("li"));
				if (!pagesLevel.ContainsKey(filePath))
				{
					pagesLevel.Add(filePath, pageLevel);
				}
			}

			return pagesLevel;
		}
		#endregion

		#region Private Methods
		/// <summary>
		/// Get the Opf file path
		/// </summary>
		/// <returns></returns>
		private string GetOpfFilePath()
		{
			string containerXmlPath = Path.Combine(_mEpubDir, "meta-inf\\container.xml");

			XmlDocument xDoc = new XmlDocument();
			xDoc.Load(containerXmlPath);

			XmlNode node = xDoc.SelectSingleNode("//*[@media-type = \"application/oebps-package+xml\"]");

			return node != null && node.Attributes != null && node.Attributes["full-path"] != null
				? Path.Combine(_mEpubDir, node.Attributes["full-path"].Value) 
				: String.Empty;
		}

		/// <summary>
		/// Converts the MathML blocks to be readable by OneNote
		/// </summary>
		/// <param name="htmlDoc"></param>
		private void ConvertMathMl(HtmlDocument htmlDoc)
		{
			HtmlNodeCollection mathNodes = htmlDoc.DocumentNode.SelectNodes("//math");
			if (mathNodes == null)
			{
				return;
			}

			foreach (var mathNode in mathNodes.ToList())
			{
				mathNode.Attributes.RemoveAll();
				HtmlAttribute mathMlNamespaceAttr = htmlDoc.CreateAttribute("xmlns:mml", MathMlNameSpace);
				mathNode.Attributes.Add(mathMlNamespaceAttr);

				foreach (var node in mathNode.DescendantsAndSelf())
				{
					node.Name = "mml:" + node.Name;
				}

				string newMathMlString = String.Format(MathMlOutline, mathNode.OuterHtml);
				HtmlCommentNode newMathNode = htmlDoc.CreateComment(newMathMlString);
				mathNode.ParentNode.ReplaceChild(newMathNode, mathNode);
			}
		}

		/// <summary>
		/// Replace the relative paths with full paths
		/// </summary>
		/// <param name="htmlDoc"></param>
		/// <param name="htmlPath"></param>
		private void UseFullPathForLinks(HtmlDocument htmlDoc, string htmlPath)
		{
			HtmlNodeCollection links = htmlDoc.DocumentNode.SelectNodes("//*[@background or @lowsrc or @src or @href]");
			if (links == null)
			{
				return;
			}

			// Get the directory where the html path exists
			string htmlDir = Path.GetDirectoryName(htmlPath);
			if (htmlDir == null)
			{
				return;
			}

			foreach (var link in links)
			{
				// Background
				if (link.Attributes["background"] != null)
				{
					link.Attributes["background"].Value = Path.GetFullPath(Path.Combine(htmlDir, link.Attributes["background"].Value));
				}

				// Lowsrc
				if (link.Attributes["lowsrc"] != null)
				{
					link.Attributes["lowsrc"].Value = Path.GetFullPath(Path.Combine(htmlDir, link.Attributes["lowsrc"].Value));
				}

				// src
				if (link.Attributes["src"] != null)
				{
					link.Attributes["src"].Value = Path.GetFullPath(Path.Combine(htmlDir, link.Attributes["src"].Value));
				}

				// href
				if (link.Attributes["href"] != null && File.Exists(Path.Combine(htmlDir, link.Attributes["href"].Value)))
				{
					link.Attributes["href"].Value = Path.GetFullPath(Path.Combine(htmlDir, link.Attributes["href"].Value));
				}
				else if (link.Attributes["href"] != null && link.Attributes["href"].Value.Contains('#'))
				{
					int indexOfHash = link.Attributes["href"].Value.LastIndexOf('#');
					string relativePath = link.Attributes["href"].Value.Substring(0, indexOfHash);
					string fullPath = File.Exists(Path.Combine(htmlDir, relativePath))
						? Path.GetFullPath(Path.Combine(htmlDir, relativePath))
						: link.Attributes["href"].Value;
					link.Attributes["href"].Value = indexOfHash == 0 ? htmlPath : fullPath;
				}
			}
		}

		/// <summary>
		/// Add the dimenssions of images to the Html ensuring that the images has max width of 960 pixels
		/// </summary>
		/// <param name="htmlDoc"></param>
		private void AddImageDimenssionsToHtml(HtmlDocument htmlDoc)
		{
			HtmlNodeCollection imgNodes = htmlDoc.DocumentNode.SelectNodes("//img[@src]");
			if (imgNodes == null)
			{
				return;
			}
			foreach (HtmlNode imgNode in imgNodes)
			{
				Image img;
				try
				{
					img = Image.FromFile(imgNode.Attributes["src"].Value);
				}
				catch (Exception)
				{
					Console.WriteLine("Error in AddImageDimenssionsToHtml: No image found at {0}", imgNode.Attributes["src"].Value);
					continue;
				}

				HtmlAttribute heightAttr = htmlDoc.CreateAttribute("height");
				HtmlAttribute widthAttr = htmlDoc.CreateAttribute("width");
				if (img.Width > MaxWidth)
				{
					heightAttr.Value = ((img.Height * MaxWidth) / (img.Width)).ToString(CultureInfo.CurrentCulture);
					widthAttr.Value = MaxWidth.ToString(CultureInfo.CurrentCulture);
				}
				else
				{
					heightAttr.Value = img.Height.ToString();
					widthAttr.Value = img.Width.ToString();
				}
				img.Dispose();

				imgNode.Attributes.Add(heightAttr);
				imgNode.Attributes.Add(widthAttr);
			}
		}
		#endregion
	}
}
