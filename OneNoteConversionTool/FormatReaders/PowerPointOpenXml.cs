using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;


namespace OneNoteConversionTool.FormatReaders
{
	/// <summary>
	/// Obtains data from power point file using OpenXml
	/// </summary>
	public class PowerPointOpenXml
	{
		#region Global Variables
		private const string NewLine = "\n";
		private readonly string _mPath;
		#endregion

		public PowerPointOpenXml(string filePath)
		{
			_mPath = filePath;
		}

		#region All Slides Info
		/// <summary>
		/// Gets the titles of all the slides
		/// </summary>
		/// <returns></returns>
		public List<string> GetAllSlideTitles()
		{
			var titles = new List<string>();

			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				if (pDoc.PresentationPart != null)
				{
					foreach (var relsId in GetSlideRIds(pDoc))
					{
						SlidePart slidePart = pDoc.PresentationPart.GetPartById(relsId) as SlidePart;
						titles.Add(GetSlideTitle(slidePart));
					}
				}
			}

			return titles;
		}

		/// <summary>
		/// Gets the notes of all the slides
		/// </summary>
		/// <returns></returns>
		public List<string> GetAllSlideNotes()
		{
			var notes = new List<string>();

			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				if (pDoc.PresentationPart != null)
				{
					List<StringValue> slideRIds = GetSlideRIds(pDoc);
					foreach (var relsId in slideRIds)
					{
						SlidePart slidePart = pDoc.PresentationPart.GetPartById(relsId) as SlidePart;
						notes.Add(GetSlideNotes(slidePart));
					}
				}
			}

			return notes;
		}

		/// <summary>
		/// Gets the comments of all the slides
		/// </summary>
		/// <returns></returns>
		public List<string> GetAllSlideComments(bool includeAuthors = true)
		{
			var comments = new List<string>();

			var authors = new List<CommentAuthor>();
			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				if (pDoc.PresentationPart != null
					&& pDoc.PresentationPart.CommentAuthorsPart != null
					&& pDoc.PresentationPart.CommentAuthorsPart.CommentAuthorList != null)
				{
					authors = pDoc.PresentationPart.CommentAuthorsPart.CommentAuthorList.Elements<CommentAuthor>().ToList();
				}

				if (pDoc.PresentationPart != null)
				{
					List<StringValue> slideRIds = GetSlideRIds(pDoc);
					foreach (var relsId in slideRIds)
					{
						SlidePart slidePart = pDoc.PresentationPart.GetPartById(relsId) as SlidePart;
						comments.Add(GetSlideComments(slidePart, authors, includeAuthors));
					}
				}
			}

			return comments;
		}

		/// <summary>
		/// Get the number of slides
		/// </summary>
		/// <param name="includeHidden">True (default) if hidden slides should be included, false others</param>
		/// <returns></returns>
		public int NumberOfSlides(bool includeHidden = true)
		{
			int slideCount = 0;

			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				if (pDoc.PresentationPart != null && pDoc.PresentationPart.SlideParts != null)
				{
					slideCount = includeHidden
						? pDoc.PresentationPart.SlideParts.Count()
						: pDoc.PresentationPart.SlideParts.Count(s => !IsHiddenSlide(s));
				}
			}

			return slideCount;
		}
		#endregion

		#region Single Slide Info
		/// <summary>
		/// Gets the title of the slide number
		/// </summary>
		/// <param name="slideNumber">slide number (one index)</param>
		/// <returns>title of the slide</returns>
		public string GetSlideTitle(int slideNumber)
		{
			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				SlidePart slidePart = GetSlidePart(pDoc, slideNumber);
				return GetSlideTitle(slidePart);
			}
		}

		/// <summary>
		/// Gets the notes of the slide
		/// </summary>
		/// <param name="slideNumber">slide number (one index)</param>
		/// <returns>notes on the slide</returns>
		public string GetSlideNotes(int slideNumber)
		{
			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				SlidePart slidePart = GetSlidePart(pDoc, slideNumber);
				return GetSlideNotes(slidePart);
			}
		}

		/// <summary>
		/// Gets the comments of the slide
		/// </summary>
		/// <param name="slideNumber">slide number (one index)</param>
		/// <returns>comments on the slide</returns>
		public string GetSlideComments(int slideNumber, bool includeAuthor = true)
		{
			var authors = new List<CommentAuthor>();
			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				if (pDoc.PresentationPart != null
					&& pDoc.PresentationPart.CommentAuthorsPart != null
					&& pDoc.PresentationPart.CommentAuthorsPart.CommentAuthorList != null)
				{
					authors = pDoc.PresentationPart.CommentAuthorsPart.CommentAuthorList.Elements<CommentAuthor>().ToList();
				}

				SlidePart slidePart = GetSlidePart(pDoc, slideNumber);
				return GetSlideComments(slidePart, authors, includeAuthor);
			}
		}

		/// <summary>
		/// Gets if the slide is hidden or not
		/// </summary>
		/// <param name="slideNumber">slide number (one index)</param>
		/// <returns>True if the slide is hidden, false otherwise</returns>
		public bool IsHiddenSlide(int slideNumber)
		{
			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				SlidePart slidePart = GetSlidePart(pDoc, slideNumber);
				return IsHiddenSlide(slidePart);
			}
		}
		#endregion

		#region Sections
		/// <summary>
		/// Get if the power point presentation has sections or not
		/// </summary>
		/// <returns></returns>
		public bool HasSections()
		{
			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				return GetListOfSections(pDoc).Count() != 0;
			}
		}

		/// <summary>
		/// Get names of all the sections
		/// </summary>
		/// <returns></returns>
		public List<string> GetSectionNames()
		{
			var sectionNames = new List<string>();

			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				List<Section> sections = GetListOfSections(pDoc);
				sectionNames.AddRange(sections.Select(section => GetSectionName(section)));
			}

			return sectionNames;
		}

		/// <summary>
		/// Get the slides that are in each section
		/// </summary>
		/// <returns>list of list of slides for each section
		/// each slide is represented as its slide number (one-indexed)</returns>
		public List<List<int>> GetSlidesInSections()
		{
			var slidesInSections = new List<List<int>>();

			using (var pDoc = PresentationDocument.Open(_mPath, false))
			{
				List<Section> sections = GetListOfSections(pDoc);
				slidesInSections.AddRange(sections.Select(section => GetSlidesInSection(pDoc, section)));
			}

			return slidesInSections;
		} 
		#endregion

		#region Private Methods
		#region Slide Info
		/// <summary>
		/// Get the title of the slide.
		/// Returns String.Empty if it can't find it
		/// </summary>
		/// <param name="slidePart"></param>
		/// <returns></returns>
		private string GetSlideTitle(SlidePart slidePart)
		{
			if (slidePart == null || slidePart.Slide == null)
			{
				return String.Empty;
			}

			Shape shape = slidePart.Slide.Descendants<Shape>().FirstOrDefault(IsTitleShape);
			if (shape == null || shape.TextBody == null)
			{
				return String.Empty;
			}

			return shape.TextBody.InnerText;
		}

		/// <summary>
		/// Get the notes of the slide.
		/// Returns String.Empty if it can't find it
		/// </summary>
		/// <param name="slidePart"></param>
		/// <returns></returns>
		private string GetSlideNotes(SlidePart slidePart)
		{
			string notes = String.Empty;

			// check if this slide has any notes
			if (slidePart == null
			    || slidePart.NotesSlidePart == null
			    || slidePart.NotesSlidePart.NotesSlide == null)
			{
				return notes;
			}

			// Get all body shapes under notesSlide (that's where notes reside)
			List<Shape> shapes = slidePart.NotesSlidePart.NotesSlide.Descendants<Shape>().Where(IsBodyShape).ToList();
			foreach (Shape shape in shapes)
			{
				// Get all paragraphes of the notes (each paragraph is a line in the notes)
				List<Drawing.Paragraph> paragraphs = shape.Descendants<Drawing.Paragraph>().ToList();
				foreach (var paragraph in paragraphs)
				{
					// Get all runs (each run has different style)
					List<Drawing.Run> runs = paragraph.Descendants<Drawing.Run>().ToList();
					foreach (var run in runs)
					{
						List<string> styles = new List<string>();

						styles.Add("font-size:" + RunFontSize(run));
						if (IsRunBold(run))
							styles.Add("font-weight:bold");
						if (IsRunItalic(run))
							styles.Add("font-style:italic");
						if (IsRunUnderlined(run) && IsRunStrikeThrough(run))
							styles.Add("text-decoration:underline line-through");
						else if (IsRunUnderlined(run))
							styles.Add("text-decoration:underline");
						else if (IsRunStrikeThrough(run))
							styles.Add("text-decoration:line-through");

						string style = String.Format("style=\"{0}\"", String.Join(";", styles.ToArray()));
						notes += String.Format("<span {0}>{1}</span>", style, run.InnerText);
					}
					notes += NewLine;
				}
			}

			return Encoding.Default.GetString(Encoding.UTF8.GetBytes(notes));
		}

		/// <summary>
		/// Get the comments of the slide.
		/// Returns String.Empty if it can't find it
		/// </summary>
		/// <param name="slidePart">SlidePart of the slide</param>
		/// <param name="authors">List of comment authors</param>
		/// <param name="includeAuthor"></param>
		/// <returns></returns>
		private string GetSlideComments(SlidePart slidePart, List<CommentAuthor> authors, bool includeAuthor = true)
		{
			string slideComments = String.Empty;

			// check if the slide has any comment list
			if (slidePart == null
			    || slidePart.SlideCommentsPart == null
			    || slidePart.SlideCommentsPart.CommentList == null)
			{
				return slideComments;
			}

			// Get all the comments in the given slide
			List<Comment> comments = slidePart.SlideCommentsPart.CommentList.Elements<Comment>().ToList();
			foreach (var comment in comments)
			{
				// add the author name
				CommentAuthor author = authors.FirstOrDefault(a => a.Id.Value == comment.AuthorId.Value);
				if (author != null && author.Name != null && author.Name.HasValue && includeAuthor)
				{
					slideComments += "Author: " + author.Name.Value.ToString() + NewLine;
				}

				slideComments += comment.Text.InnerText + NewLine;
			}

			return Encoding.Default.GetString(Encoding.UTF8.GetBytes(slideComments));
		}

		/// <summary>
		/// Gets the SlidePart based on the slide number
		/// </summary>
		/// <param name="pDoc"></param>
		/// <param name="slideNumber"></param>
		/// <returns></returns>
		private SlidePart GetSlidePart(PresentationDocument pDoc, int slideNumber)
		{
			int index = slideNumber - 1;
			if (pDoc.PresentationPart != null)
			{
				List<StringValue> slideRIds = GetSlideRIds(pDoc);
				if (0 <= index && index < slideRIds.Count)
				{
					return pDoc.PresentationPart.GetPartById(slideRIds[index]) as SlidePart;
				}
			}

			return null;
		}

		/// <summary>
		/// Get the relationship ids of the slides
		/// </summary>
		/// <param name="pDoc"></param>
		/// <returns>list of relationship ids for slides</returns>
		private List<StringValue> GetSlideRIds(PresentationDocument pDoc)
		{
			var ids = new List<StringValue>();

			if (pDoc.PresentationPart == null
				|| pDoc.PresentationPart.Presentation == null
				|| pDoc.PresentationPart.Presentation.SlideIdList == null)
			{
				return ids;
			}

			ids.AddRange(from SlideId slideId in pDoc.PresentationPart.Presentation.SlideIdList select slideId.RelationshipId);

			return ids;
		}

		/// <summary>
		/// Get the ids of the slides
		/// </summary>
		/// <param name="pDoc"></param>
		/// <returns>list of ids for slides</returns>
		private List<UInt32Value> GetSlideIds(PresentationDocument pDoc)
		{
			var ids = new List<UInt32Value>();

			if (pDoc.PresentationPart == null
				|| pDoc.PresentationPart.Presentation == null
				|| pDoc.PresentationPart.Presentation.SlideIdList == null)
			{
				return ids;
			}

			ids.AddRange(from SlideId slideId in pDoc.PresentationPart.Presentation.SlideIdList select slideId.Id);

			return ids;
		}

		/// <summary>
		/// Gets if the slide is hidden or not
		/// </summary>
		/// <param name="slidePart">SlidePart of the slide</param>
		/// <returns>True if the slide is hidden, false otherwise</returns>
		private bool IsHiddenSlide(SlidePart slidePart)
		{
			return ((slidePart.Slide == null)
				|| (slidePart.Slide.Show != null && slidePart.Slide.Show.HasValue && !slidePart.Slide.Show.Value));
		}

		/// <summary>
		/// Gets if the slide is hidden or not
		/// </summary>
		/// <param name="slide">Slide of interest</param>
		/// <returns>True if the slide is hidden, false otherwise</returns>
		private bool IsHiddenSlide(Slide slide)
		{
			return (slide == null)
				|| (slide.Show != null && slide.Show.HasValue && !slide.Show.Value);
		}
		#endregion

		#region Section Info
		/// <summary>
		/// Get the name of the section
		/// </summary>
		/// <param name="section"></param>
		/// <returns></returns>
		private string GetSectionName(Section section)
		{

			return section.Name != null && section.Name.HasValue
				? section.Name.Value.ToString()
				: String.Empty;
		}

		/// <summary>
		/// Get the slides in a given section
		/// </summary>
		/// <param name="pDoc">the current presentation document</param>
		/// <param name="section">the section of interest</param>
		/// <returns>list of integers representating the slides in each section
		/// the slides are given by their slide number (one-indexed)</returns>
		private List<int> GetSlidesInSection(PresentationDocument pDoc, Section section)
		{
			var slides = new List<int>();

			if (section.SectionSlideIdList == null)
			{
				return slides;
			}

			List<UInt32Value> slideIds = GetSlideIds(pDoc);
			foreach (SectionSlideIdListEntry slideId in section.SectionSlideIdList)
			{
				int index = slideIds.FindIndex(id => id.Value.Equals(slideId.Id.Value));
				if (index != -1)
				{
					int slideNumber = index + 1;
					slides.Add(slideNumber);
				}
			}

			return slides;
		}

		/// <summary>
		/// Get list of all the sections
		/// </summary>
		/// <param name="pDoc"></param>
		/// <returns></returns>
		private List<Section> GetListOfSections(PresentationDocument pDoc)
		{
			var sections = new List<Section>();

			if (pDoc.PresentationPart == null)
			{
				return sections;
			}

			SectionList sectionList = pDoc.PresentationPart.Presentation.PresentationExtensionList.Descendants<SectionList>().FirstOrDefault();
			if (sectionList == null)
			{
				return sections;
			}

			sections.AddRange(sectionList.Select(section => section as Section));

			return sections;
		} 
		#endregion

		/// <summary>
		/// Gets if this shape is for title or not
		/// </summary>
		/// <param name="shape">Shape of interest</param>
		/// <returns>True if it is title shape, false otherwise</returns>
		private bool IsTitleShape(Shape shape)
		{
			if (shape == null
				|| shape.NonVisualShapeProperties == null
				|| shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties == null)
			{
				return false;
			}

			PlaceholderShape phShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();

			return phShape != null
				&& phShape.Type != null
				&& phShape.Type.HasValue
				&& (phShape.Type.Value == PlaceholderValues.Title || phShape.Type.Value == PlaceholderValues.CenteredTitle);
		}

		/// <summary>
		/// Gets if this shape is for body or not
		/// </summary>
		/// <param name="shape">Shape of interest</param>
		/// <returns>True if it is body shape, false otherwise</returns>
		private bool IsBodyShape(Shape shape)
		{
			if (shape == null
				|| shape.NonVisualShapeProperties == null
				|| shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties == null)
			{
				return false;
			}

			PlaceholderShape phShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();

			return phShape != null
				&& phShape.Type != null
				&& phShape.Type.HasValue
				&& phShape.Type.Value == PlaceholderValues.Body;
		}

		/// <summary>
		/// Gets if this run is bold or not
		/// </summary>
		/// <param name="run">Run of interest</param>
		/// <returns>True if it is bold, false otherwise</returns>
		private bool IsRunBold(Drawing.Run run)
		{
			var runProperty = run.Descendants<Drawing.RunProperties>().FirstOrDefault();
			return runProperty.Bold != null && runProperty.Bold.Value;
		}

		/// <summary>
		/// Gets if this run is italic or not
		/// </summary>
		/// <param name="run">Run of interest</param>
		/// <returns>True if it is italic, false otherwise</returns>
		private bool IsRunItalic(Drawing.Run run)
		{
			var runProperty = run.Descendants<Drawing.RunProperties>().FirstOrDefault();
			return runProperty.Italic != null && runProperty.Italic.Value;
		}

		/// <summary>
		/// Gets if this run is underlined or not
		/// </summary>
		/// <param name="run">Run of interest</param>
		/// <returns>True if it is underlined, false otherwise</returns>
		private bool IsRunUnderlined(Drawing.Run run)
		{
			var runProperty = run.Descendants<Drawing.RunProperties>().FirstOrDefault();
			return runProperty.Underline != null
				&& runProperty.Underline.Value != Drawing.TextUnderlineValues.None;
		}

		/// <summary>
		/// Gets if this run is Strike-through or not
		/// </summary>
		/// <param name="run">Run of interest</param>
		/// <returns>True if it is strike-through, false otherwise</returns>
		private bool IsRunStrikeThrough(Drawing.Run run)
		{
			var runProperty = run.Descendants<Drawing.RunProperties>().FirstOrDefault();
			return runProperty.Strike != null 
				&& runProperty.Strike.Value != Drawing.TextStrikeValues.NoStrike;
		}

		/// <summary>
		/// Gets the size of the text in the given run
		/// </summary>
		/// <param name="run">Run of interest</param>
		/// <returns>Size of the font in points,
		/// default is 11.0pt</returns>
		private string RunFontSize(Drawing.Run run)
		{
			var runProperty = run.Descendants<Drawing.RunProperties>().FirstOrDefault();
			if (runProperty.FontSize != null)
				return String.Format("{0:0.0}pt", runProperty.FontSize.Value / 100.0);
			else
				return "11.0pt";
		}
		#endregion
	}
}