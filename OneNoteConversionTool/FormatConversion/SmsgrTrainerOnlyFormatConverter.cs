using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneNoteConversionTool.FormatConversion
{
	class SmsgrTrainerOnlyFormatConverter : SmsgrFormatConverter
	{
		private const string InputFormat = "OneNote Courseware Output - Trainer Only";

		/// <summary>
		/// Returns the name of the input format that this IFormatConverter supports
		/// </summary>
		/// <returns></returns>
		public override string GetSupportedInputFormat()
		{
			return InputFormat;
		}

		/// <summary>
		/// Gets whether the trainer notebook should be created or not
		/// </summary>
		/// <returns></returns>
		protected override bool IncludeTrainerNotebook()
		{
			return true;
		}

		/// <summary>
		/// Gets whether the trainer notebook should be created or not
		/// </summary>
		/// <returns></returns>
		protected override bool IncludeStudentNotebook()
		{
			return false;
		}
	}
}
