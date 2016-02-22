using System.Collections.Generic;
using OneNoteConversionTool.FormatConversion;

namespace OneNoteConversionToolUnitTest.FormatConversion
{
	/// <summary>
	/// Mock conversion class for unit testing
	/// </summary>
	class MockConversionManager : ConversionManager
	{
		/// <summary>
		/// Replaces the registered converters with mock ones
		/// </summary>
		public static void InitializeWithMockData()
		{
			var converter = new MockFormatConverter();
			SupportedConverters = new Dictionary<string, IFormatConverter>();
			SupportedConverters.Add(converter.GetSupportedInputFormat(), converter);
		}

		/// <summary>
		/// Returns the instance of the MockFormatConverter
		/// </summary>
		/// <returns></returns>
		public static MockFormatConverter GetMockFormatConverter()
		{
			return SupportedConverters[MockFormatConverter.InputFormat] as MockFormatConverter;
		}
	}
}
