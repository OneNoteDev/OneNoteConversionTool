# OneNoteConversionTool

This is a Windows Application for converting documents of various formats into OneNote.

Currently supported formats:
- Microsoft Word (.doc, .docx, .dot, .dotx, .docm, .dotm, .odt)
- Microsoft PowerPoint (.ppt, .pptx, .pot, .potx, .pptm, .potm, .odp)
- PDF (.pdf)
- Epub (.epub)
- InDesign Document (.indd)

## How to open

Open the OneNoteConversionTool.sln file in Visual Studio 2013 or later.

In order to compile, you will need:
- Microsoft Word 2013 installed
- Microsoft PowerPoint 2013 installed

If you would like to add support for InDesign documents, you need to:
- have Adobe InDesign installed
- define `INDESIGN_INSTALLED` in the compiler options
