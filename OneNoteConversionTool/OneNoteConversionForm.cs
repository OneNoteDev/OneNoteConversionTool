using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;
using OneNoteConversionTool.FormatConversion;

namespace OneNoteConversionTool
{
	public partial class OneNoteConversionForm : Form
	{
		private CommonOpenFileDialog _mInputFolderDialog;
		private CommonOpenFileDialog _mOutputPathDialog;

		public OneNoteConversionForm()
		{
			InitializeComponent();
			InitializeDefaultSelections();
		}

		private void InitializeDefaultSelections()
		{
			// Set filter options and filter index
			openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			openFileDialog.Filter = GetFileDialogFilter();
			openFileDialog.FilterIndex = 0;
			openFileDialog.Multiselect = true;

			// Set default input folder
			_mInputFolderDialog = new CommonOpenFileDialog();
			_mInputFolderDialog.IsFolderPicker = true;
			_mInputFolderDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

			// Set default input format
			fileFormatComboBox.Items.AddRange(ConversionManager.GetSupportedFormats().ToArray());
			fileFormatComboBox.SelectedIndex = 0;

			// Set default output folder
			_mOutputPathDialog = new CommonOpenFileDialog();
			_mOutputPathDialog.IsFolderPicker = true;
			_mOutputPathDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			outputLocationTextBox.Text = _mOutputPathDialog.InitialDirectory;

			// Event handlers
			fileRadioButton.Click += fileRadioButton_isClicked;
			folderRadioButton.Click += folderRadioButton_isClicked;
		}

		private void browseFileButton_Click(object sender, EventArgs e)
		{
			var startingDirectory = string.Empty;
			if (!string.IsNullOrEmpty(openFileTextBox.Text))
			{
				startingDirectory = Path.GetDirectoryName(openFileTextBox.Text);
			}

			if (fileRadioButton.Checked)
			{
				if (!string.IsNullOrEmpty(startingDirectory) && Directory.Exists(startingDirectory))
				{
					openFileDialog.InitialDirectory = startingDirectory;
				}
				else
				{
					openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
					openFileTextBox.Text = string.Empty;
				}

				DialogResult result = openFileDialog.ShowDialog();

				if (result == DialogResult.OK) // Test result.
				{
					openFileTextBox.Text = openFileDialog.FileName;
				}
			}
			else if (folderRadioButton.Checked)
			{
				if (!string.IsNullOrEmpty(startingDirectory) && Directory.Exists(startingDirectory))
				{
					_mInputFolderDialog.InitialDirectory = startingDirectory;
				}
				else
				{
					_mInputFolderDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
					openFileTextBox.Text = string.Empty;
				}

				var result = _mInputFolderDialog.ShowDialog();

				if (result == CommonFileDialogResult.Ok)
				{
					openFileTextBox.Text = _mInputFolderDialog.FileName;
				}
			}
		}

		private void convertButton_Click(object sender, EventArgs e)
		{
			try
			{
                ConversionManager.ConvertInput(fileFormatComboBox.Text, openFileTextBox.Text, outputLocationTextBox.Text);
                MessageBox.Show("Finished conversion. Output File is located at:\n" + outputLocationTextBox.Text);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error during conversion:\n" + ex.Message);
			}
		}


		private void outputBrowseButton_Click(object sender, EventArgs e)
		{
			var result = _mOutputPathDialog.ShowDialog();

			if (result == CommonFileDialogResult.Ok)
			{
				outputLocationTextBox.Text = _mOutputPathDialog.FileName;
			}
		}

		private void fileRadioButton_isClicked(object sender, EventArgs e)
		{
			fileRadioButton.Checked = true;
			folderRadioButton.Checked = false;
		}

		private void folderRadioButton_isClicked(object sender, EventArgs e)
		{
			fileRadioButton.Checked = false;
			folderRadioButton.Checked = true;
		}

		private string GetFileDialogFilter()
		{
			string[] fileDialogFilterStrings =
			{
				"All|*.*",
				"Word Files|*" + String.Join(";*", Utility.WordSupportedExtenssions),
				"PowerPoint Presentation|*" + String.Join(";*", Utility.PowerPointSupportedExtenssions),
				"PDF Files|*" + String.Join(";*", Utility.PdfSupportedExtenssions),
				"Epub|*" + String.Join(";*", Utility.EpubSupportedExtenssions),
				"InDesign|*" + String.Join(";*", Utility.InDesignSupportedExtenssions)
			};
			return String.Join("|", fileDialogFilterStrings);
		}
	}
}
