using OneNoteConversionTool.FormatConversion;

namespace OneNoteConversionToolUnitTest.FormatConversion
{
	/// <summary>
	/// A mock FormatConverter that implements IFormatConverter for unit testing purposes
	/// </summary>
	class MockFormatConverter : IFormatConverter
	{
		public const string InputFormat = "MockFormat";

		public bool IsWordToOneNoteCalled = false;
		public bool IsPdfToOneNoteCalled = false;
		public bool IsPowerPointToOneNoteCalled = false;
		public bool IsInDesignToOneNoteCalled = false;
		public bool IsEpubToOneNoteCalled = false;

		public bool ConvertWordToOneNote(string inputFile, string outputDir)
		{
			IsWordToOneNoteCalled = true;
			return IsWordToOneNoteCalled;
		}

		public bool ConvertPdfToOneNote(string inputFile, string outputDir)
		{
			IsPdfToOneNoteCalled = true;
			return IsPdfToOneNoteCalled;
		}

		public bool ConvertPowerPointToOneNote(string inputFile, string outputDir)
		{
			IsPowerPointToOneNoteCalled = true;
			return IsPowerPointToOneNoteCalled;
		}

		public bool ConvertInDesignToOneNote(string inputFile, string outputDir)
		{
			IsInDesignToOneNoteCalled = true;
			return IsInDesignToOneNoteCalled;
		}

		public bool ConvertEpubToOneNote(string inputFile, string outputDir)
		{
			IsEpubToOneNoteCalled = true;
			return IsEpubToOneNoteCalled;
		}

		public string GetSupportedInputFormat()
		{
			return InputFormat;
		}
	}
}
