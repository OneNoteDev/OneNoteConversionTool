using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneNoteConversionTool.FormatConversion
{
	public interface IFormatConverter
	{
		/// <summary>
		/// Converts the Word document input file into OneNote
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		bool ConvertWordToOneNote(string inputFile, string outputDir);

		/// <summary>
		/// Converts the PDF document input file into OneNote
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		bool ConvertPdfToOneNote(string inputFile, string outputDir);

		/// <summary>
		/// Converts the PowerPoint document input file into OneNote
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		bool ConvertPowerPointToOneNote(string inputFile, string outputDir);

		/// <summary>
		/// Converts the InDesign document input file into OneNote
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		/// <returns></returns>
		bool ConvertInDesignToOneNote(string inputFile, string outputDir);


		/// <summary>
		/// Converts the ePub document input file into OneNote
		/// </summary>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		/// <returns></returns>
		bool ConvertEpubToOneNote(string inputFile, string outputDir);

		/// <summary>
		/// Returns the name of the input format that this IFormatConverter supports
		/// </summary>
		/// <returns></returns>
		string GetSupportedInputFormat();
	}
}
