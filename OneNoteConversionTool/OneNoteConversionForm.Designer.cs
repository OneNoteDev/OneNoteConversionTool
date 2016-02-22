namespace OneNoteConversionTool
{
	partial class OneNoteConversionForm
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.browseFileButton = new System.Windows.Forms.Button();
            this.openFileTextBox = new System.Windows.Forms.TextBox();
            this.fileFormatComboBox = new System.Windows.Forms.ComboBox();
            this.convertButton = new System.Windows.Forms.Button();
            this.openFileLabel = new System.Windows.Forms.Label();
            this.fileFormatLabel = new System.Windows.Forms.Label();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.outputLocationLabel = new System.Windows.Forms.Label();
            this.outputLocationTextBox = new System.Windows.Forms.TextBox();
            this.outputBrowseButton = new System.Windows.Forms.Button();
            this.fileRadioButton = new System.Windows.Forms.RadioButton();
            this.folderRadioButton = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // browseFileButton
            // 
            this.browseFileButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.browseFileButton.Location = new System.Drawing.Point(421, 88);
            this.browseFileButton.Name = "browseFileButton";
            this.browseFileButton.Size = new System.Drawing.Size(75, 23);
            this.browseFileButton.TabIndex = 3;
            this.browseFileButton.Text = "Browse ...";
            this.browseFileButton.UseVisualStyleBackColor = true;
            this.browseFileButton.Click += new System.EventHandler(this.browseFileButton_Click);
            // 
            // openFileTextBox
            // 
            this.openFileTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.openFileTextBox.Location = new System.Drawing.Point(12, 62);
            this.openFileTextBox.Name = "openFileTextBox";
            this.openFileTextBox.Size = new System.Drawing.Size(484, 20);
            this.openFileTextBox.TabIndex = 2;
            // 
            // fileFormatComboBox
            // 
            this.fileFormatComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fileFormatComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.fileFormatComboBox.FormattingEnabled = true;
            this.fileFormatComboBox.Location = new System.Drawing.Point(12, 128);
            this.fileFormatComboBox.Name = "fileFormatComboBox";
            this.fileFormatComboBox.Size = new System.Drawing.Size(484, 21);
            this.fileFormatComboBox.TabIndex = 4;
            // 
            // convertButton
            // 
            this.convertButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.convertButton.Location = new System.Drawing.Point(12, 252);
            this.convertButton.Name = "convertButton";
            this.convertButton.Size = new System.Drawing.Size(484, 58);
            this.convertButton.TabIndex = 7;
            this.convertButton.Text = "Convert to OneNote";
            this.convertButton.UseVisualStyleBackColor = true;
            this.convertButton.Click += new System.EventHandler(this.convertButton_Click);
            // 
            // openFileLabel
            // 
            this.openFileLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.openFileLabel.AutoSize = true;
            this.openFileLabel.Location = new System.Drawing.Point(9, 23);
            this.openFileLabel.Name = "openFileLabel";
            this.openFileLabel.Size = new System.Drawing.Size(106, 13);
            this.openFileLabel.TabIndex = 4;
            this.openFileLabel.Text = "Input path to convert";
            // 
            // fileFormatLabel
            // 
            this.fileFormatLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fileFormatLabel.AutoSize = true;
            this.fileFormatLabel.Location = new System.Drawing.Point(9, 112);
            this.fileFormatLabel.Name = "fileFormatLabel";
            this.fileFormatLabel.Size = new System.Drawing.Size(79, 13);
            this.fileFormatLabel.TabIndex = 5;
            this.fileFormatLabel.Text = "Input file format";
            // 
            // outputLocationLabel
            // 
            this.outputLocationLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.outputLocationLabel.AutoSize = true;
            this.outputLocationLabel.Location = new System.Drawing.Point(10, 169);
            this.outputLocationLabel.Name = "outputLocationLabel";
            this.outputLocationLabel.Size = new System.Drawing.Size(79, 13);
            this.outputLocationLabel.TabIndex = 8;
            this.outputLocationLabel.Text = "Output location";
            // 
            // outputLocationTextBox
            // 
            this.outputLocationTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.outputLocationTextBox.Location = new System.Drawing.Point(12, 185);
            this.outputLocationTextBox.Name = "outputLocationTextBox";
            this.outputLocationTextBox.Size = new System.Drawing.Size(484, 20);
            this.outputLocationTextBox.TabIndex = 5;
            // 
            // outputBrowseButton
            // 
            this.outputBrowseButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.outputBrowseButton.Location = new System.Drawing.Point(421, 211);
            this.outputBrowseButton.Name = "outputBrowseButton";
            this.outputBrowseButton.Size = new System.Drawing.Size(75, 23);
            this.outputBrowseButton.TabIndex = 6;
            this.outputBrowseButton.Text = "Browse ...";
            this.outputBrowseButton.UseVisualStyleBackColor = true;
            this.outputBrowseButton.Click += new System.EventHandler(this.outputBrowseButton_Click);
            // 
            // fileRadioButton
            // 
            this.fileRadioButton.AutoSize = true;
            this.fileRadioButton.Checked = true;
            this.fileRadioButton.Location = new System.Drawing.Point(13, 39);
            this.fileRadioButton.Name = "fileRadioButton";
            this.fileRadioButton.Size = new System.Drawing.Size(41, 17);
            this.fileRadioButton.TabIndex = 0;
            this.fileRadioButton.TabStop = true;
            this.fileRadioButton.Text = "File";
            this.fileRadioButton.UseVisualStyleBackColor = true;
            // 
            // folderRadioButton
            // 
            this.folderRadioButton.AutoSize = true;
            this.folderRadioButton.Location = new System.Drawing.Point(61, 39);
            this.folderRadioButton.Name = "folderRadioButton";
            this.folderRadioButton.Size = new System.Drawing.Size(54, 17);
            this.folderRadioButton.TabIndex = 1;
            this.folderRadioButton.Text = "Folder";
            this.folderRadioButton.UseVisualStyleBackColor = true;
            // 
            // OneNoteConversionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(509, 332);
            this.Controls.Add(this.folderRadioButton);
            this.Controls.Add(this.fileRadioButton);
            this.Controls.Add(this.outputLocationLabel);
            this.Controls.Add(this.outputLocationTextBox);
            this.Controls.Add(this.outputBrowseButton);
            this.Controls.Add(this.fileFormatLabel);
            this.Controls.Add(this.openFileLabel);
            this.Controls.Add(this.convertButton);
            this.Controls.Add(this.fileFormatComboBox);
            this.Controls.Add(this.openFileTextBox);
            this.Controls.Add(this.browseFileButton);
            this.Name = "OneNoteConversionForm";
            this.Text = "OneNote Conversion Tool";
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button browseFileButton;
		private System.Windows.Forms.TextBox openFileTextBox;
		private System.Windows.Forms.ComboBox fileFormatComboBox;
		private System.Windows.Forms.Button convertButton;
		private System.Windows.Forms.Label openFileLabel;
		private System.Windows.Forms.Label fileFormatLabel;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.Label outputLocationLabel;
		private System.Windows.Forms.TextBox outputLocationTextBox;
		private System.Windows.Forms.Button outputBrowseButton;
		private System.Windows.Forms.RadioButton folderRadioButton;
		private System.Windows.Forms.RadioButton fileRadioButton;
	}
}

