using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace OneNoteConversionTool.FormatConversion
{
	/// <summary>
	/// Handles converting the different input types into OneNote (output)
	/// </summary>
	public class ConversionManager
	{
		private static Dictionary<string, IFormatConverter> _supportedConverters;
		protected static Dictionary<string, IFormatConverter> SupportedConverters
		{
			get
			{
				//Lazy initialization
				if (_supportedConverters == null)
				{
					_supportedConverters = new Dictionary<string, IFormatConverter>();
					string inputFormat = ConfigurationManager.AppSettings.Get("InputFormat");
					IFormatConverter converter;

					//Adding generic as supported format
					converter = new GenericFormatConverter();
					_supportedConverters.Add(converter.GetSupportedInputFormat(), converter);

					if (inputFormat == "Kindercare" || inputFormat == "All")
					{
						//Adding Kindercare as supported format
						converter = new KindercareFormatConverter();
						_supportedConverters.Add(converter.GetSupportedInputFormat(), converter);
					}

					if (inputFormat == "SMSGR" || inputFormat == "All")
					{
						//Adding SMSGR as supported format
						converter = new SmsgrFormatConverter();
						_supportedConverters.Add(converter.GetSupportedInputFormat(), converter);

						//Adding SMSGR for trainer only as supported format
						converter = new SmsgrTrainerOnlyFormatConverter();
						_supportedConverters.Add(converter.GetSupportedInputFormat(), converter);

						//Adding SMSGR for student only as supported format
						converter = new SmsgrStudentOnlyFormatConverter();
						_supportedConverters.Add(converter.GetSupportedInputFormat(), converter);
					}
				}
				return _supportedConverters;
			}
			set { _supportedConverters = value; }
		}

		/// <summary>
		/// Returns the list of all supported input formats
		/// </summary>
		/// <returns></returns>
		public static List<string> GetSupportedFormats()
		{
			return SupportedConverters.Keys.ToList();
		}

		/// <summary>
		/// Converts Input using the specified format
		/// </summary>
		/// <param name="converterType"></param>
		/// <param name="inputPath"></param>
		/// <param name="outputDir"></param>
		public static void ConvertInput(string converterType, string inputPath, string outputDir)
		{
			IFormatConverter converter;
			if (SupportedConverters.TryGetValue(converterType, out converter))
			{
				if ((File.GetAttributes(inputPath) & FileAttributes.Directory) == FileAttributes.Directory)
				{
					ConvertDirectory(converter, inputPath, outputDir);
				}
				else
				{
					ConvertFile(converter, inputPath, outputDir);
				}
			}
			else
			{
				throw new NotSupportedException("Unable to find the converter for this format type.");
			}
		}

		/// <summary>
		/// Helper method to iterate through a directory and convert all the files underneath it
		/// </summary>
		/// <param name="converter"></param>
		/// <param name="inputDir"></param>
		/// <param name="outputDir"></param>
		private static void ConvertDirectory(IFormatConverter converter, string inputDir, string outputDir)
		{
			var notSupportedFiles = ConvertDirectoryImpl(converter, inputDir, outputDir);
			
			if (notSupportedFiles.Count > 0)
			{
				throw new NotSupportedException(string.Format("The following files were not converted:\n{0}",
					string.Join("\n", notSupportedFiles.ToArray())));
			}
		}

		/// <summary>
		/// The internal implementation of the method that goes into the directory and tries to convert all
		/// files underneath it recursively.
		/// </summary>
		/// <param name="converter"></param>
		/// <param name="inputDir"></param>
		/// <param name="outputDir"></param>
		/// <returns></returns>
		private static ICollection<string> ConvertDirectoryImpl(IFormatConverter converter, string inputDir, string outputDir)
		{
			var files = Directory.GetFiles(inputDir);
			var dirs = Directory.GetDirectories(inputDir);
			var notSupportedFiles = new List<string>();

			// Iterate through all files in this directory
			foreach (var file in files)
			{
				try
				{
					ConvertFile(converter, file, outputDir);
				}
				catch (NotSupportedException)
				{
					//Do not block the conversion process, but keep the list of files unable to be converted
					notSupportedFiles.Add(file);
				}
			}

			// Iterate through all directories in this directory
			foreach (var dir in dirs)
			{
				notSupportedFiles.AddRange(ConvertDirectoryImpl(converter, dir, outputDir));
			}

			return notSupportedFiles;
		}

		/// <summary>
		/// Converts a single file into a section in OneNote notebook
		/// </summary>
		/// <param name="converter"></param>
		/// <param name="inputFile"></param>
		/// <param name="outputDir"></param>
		private static void ConvertFile(IFormatConverter converter, string inputFile, string outputDir)
		{
			var file = new FileInfo(inputFile);
			if (Utility.WordSupportedExtenssions.Any(
				ext => ext.Equals(file.Extension, StringComparison.InvariantCultureIgnoreCase)))
			{
				converter.ConvertWordToOneNote(inputFile, outputDir);
			}
			else if (Utility.PowerPointSupportedExtenssions.Any(
				ext => ext.Equals(file.Extension, StringComparison.InvariantCultureIgnoreCase)))
			{
				converter.ConvertPowerPointToOneNote(inputFile, outputDir);
			}
			else if (Utility.PdfSupportedExtenssions.Any(
				ext => ext.Equals(file.Extension, StringComparison.InvariantCultureIgnoreCase)))
			{
				converter.ConvertPdfToOneNote(inputFile, outputDir);
			}
			else if (Utility.EpubSupportedExtenssions.Any(
				ext => ext.Equals(file.Extension, StringComparison.InvariantCultureIgnoreCase)))
			{
				converter.ConvertEpubToOneNote(inputFile, outputDir);
			}
			else if (Utility.InDesignSupportedExtenssions.Any(
				ext => ext.Equals(file.Extension, StringComparison.InvariantCultureIgnoreCase)))
			{
				converter.ConvertInDesignToOneNote(inputFile, outputDir);
			}
		}
	}
}
