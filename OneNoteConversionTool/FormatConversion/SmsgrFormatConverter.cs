using System;
using System.Collections.Generic;
using OneNoteConversionTool.FormatReaders;
using OneNoteConversionTool.OutputGenerator;

namespace OneNoteConversionTool.FormatConversion
{
	/// <summary>
	/// Converter that accepts SMSGR Format
	/// </summary>
	public class SmsgrFormatConverter : GenericFormatConverter
	{
		private const string InputFormat = "OneNote Courseware Output";
		private const string TrainerNotebook = "Trainer Notebook";
		private const string StudentNotebook = "Student Notebook";
		private const string TrainerNotesTitle = "Trainer Notes";
		private const string StudentNotesTitle = "Student Notes";


		/// <summary>
		/// Returns the name of the input format that this IFormatConverter supports
		/// </summary>
		/// <returns></returns>
		public override string GetSupportedInputFormat()
		{
			return InputFormat;
		}

		/// <summary>
		/// Converts PowerPoint presentan to OneNote while converting the sections in power point to main pages, and slides to sub pages
		/// It creates two notebooks, one for the Trainer and one for the student
		/// </summary>
		/// <param name="pptOpenXml"></param>
		/// <param name="imgsPath"></param>
		/// <param name="note"></param>
		/// <param name="sectionName"></param>
		protected override void ConvertPowerPointToOneNote(PowerPointOpenXml pptOpenXml, string imgsPath, OneNoteGenerator note,
			string sectionName)
		{
			string trainerNotebookId = String.Empty;
			string trainerSectionId = String.Empty;
			if (IncludeTrainerNotebook())
			{
				// Create the student notebook
				trainerNotebookId = note.CreateNotebook(TrainerNotebook);
				trainerSectionId = note.CreateSection(sectionName, trainerNotebookId);
			}

			string studentNotebookId = String.Empty;
			string studentSectionId = String.Empty;
			if (IncludeStudentNotebook())
			{
				// Create the student notebook
				studentNotebookId = note.CreateNotebook(StudentNotebook);
				studentSectionId = note.CreateSection(sectionName, studentNotebookId);
			}

			if (pptOpenXml.HasSections())
			{
				List<string> sectionNames = pptOpenXml.GetSectionNames();
				List<List<int>> slidesInSections = pptOpenXml.GetSlidesInSections();

				if (IncludeTrainerNotebook())
					ConvertPowerPointWithSectionsToOneNote(pptOpenXml, imgsPath, note, trainerSectionId, sectionNames, slidesInSections, true);

				if (IncludeStudentNotebook())
					ConvertPowerPointWithSectionsToOneNote(pptOpenXml, imgsPath, note, studentSectionId, sectionNames, slidesInSections, false);
			}
			else
			{
				if (IncludeTrainerNotebook())
					ConvertPowerPointWithoutSectionsToOneNote(pptOpenXml, imgsPath, note, trainerSectionId, true);

				if (IncludeStudentNotebook())
					ConvertPowerPointWithoutSectionsToOneNote(pptOpenXml, imgsPath, note, studentSectionId, false);
			}
		}

		/// <summary>
		/// Gets whether the trainer notebook should be created or not
		/// </summary>
		/// <returns></returns>
		protected virtual bool IncludeTrainerNotebook()
		{
			return true;
		}

		/// <summary>
		/// Gets whether the trainer notebook should be created or not
		/// </summary>
		/// <returns></returns>
		protected virtual bool IncludeStudentNotebook()
		{
			return true;
		}

		#region Helper Methods
		/// <summary>
		/// Helper method to include the common code between the trainer and the student create notebook when converting 
		/// Power Point files that have sections
		/// </summary>
		/// <param name="pptOpenXml"></param>
		/// <param name="imgsPath"></param>
		/// <param name="note"></param>
		/// <param name="sectionId"></param>
		/// <param name="sectionNames"></param>
		/// <param name="slidesInSections"></param>
		/// <param name="isTrainer"></param>
		private void ConvertPowerPointWithSectionsToOneNote(PowerPointOpenXml pptOpenXml, string imgsPath, OneNoteGenerator note, 
			string sectionId, List<string> sectionNames, List<List<int>> slidesInSections, bool isTrainer)
		{
			var pptSectionsPageIds = new List<string>();

			for (int i = 0; i < sectionNames.Count; i++)
			{
				string pptSectionPageId = note.CreatePage(sectionNames[i], sectionId);
				foreach (var slideNumber in slidesInSections[i])
				{
					string pageId;
					if (isTrainer)
					{
						pageId = InsertPowerPointSlideInOneNote(slideNumber, pptOpenXml, imgsPath, note, sectionId,
							true, StudentNotesTitle, true, TrainerNotesTitle);

					}
					else
					{
						pageId = InsertPowerPointSlideInOneNote(slideNumber, pptOpenXml, imgsPath, note, sectionId,
							true, StudentNotesTitle, false);
					}
					if (!pageId.Equals(String.Empty))
					{
						note.SetSubPage(sectionId, pageId);
						note.SetShowDate(pageId, false);
						note.SetShowTime(pageId, false);
					}
				}
				pptSectionsPageIds.Add(pptSectionPageId);
			}

			string tocPageId = note.CreateTableOfContentPage(sectionId);
			note.SetShowDate(tocPageId, false);
			note.SetShowTime(tocPageId, false);

			foreach (var pptSectionPageId in pptSectionsPageIds)
			{
				note.SetCollapsePage(pptSectionPageId);
				note.SetShowDate(pptSectionPageId, false);
				note.SetShowTime(pptSectionPageId, false);
			}
		}

		/// <summary>
		/// Helper method to include the common code between the trainer and the student create notebook when converting 
		/// Power Point files that doesn't have sections
		/// </summary>
		/// <param name="pptOpenXml"></param>
		/// <param name="imgsPath"></param>
		/// <param name="note"></param>
		/// <param name="sectionId"></param>
		/// <param name="isTrainer"></param>
		private void ConvertPowerPointWithoutSectionsToOneNote(PowerPointOpenXml pptOpenXml, string imgsPath, OneNoteGenerator note,
			string sectionId, bool isTrainer)
		{
			for (var i = 1; i <= pptOpenXml.NumberOfSlides(); i++)
			{
				string pageId;
				if (isTrainer)
				{
					pageId = InsertPowerPointSlideInOneNote(i, pptOpenXml, imgsPath, note, sectionId, 
						true, StudentNotesTitle, true, TrainerNotesTitle);
				}
				else
				{
					pageId = InsertPowerPointSlideInOneNote(i, pptOpenXml, imgsPath, note, sectionId, 
						true, StudentNotesTitle, false);
				}
				if (!pageId.Equals(String.Empty))
				{
					note.SetShowDate(pageId, false);
					note.SetShowTime(pageId, false);
				}
			}
			string tocPageId = note.CreateTableOfContentPage(sectionId);
			note.SetShowDate(tocPageId, false);
			note.SetShowTime(tocPageId, false);
		}
		#endregion
	}
}