**MD Converter** is a lightweight Windows tool designed to quickly convert Markdown documents into Microsoft Word (DOCX) files. It is suitable for technical documentation, project reports, user manuals, research notes, and other structured documents.

Main features include:

* Convert Markdown files into well-formatted Word documents
* Support common Markdown syntax, including headings, paragraphs, bold, italic, code blocks, and blockquotes
* Support ordered lists, unordered lists, nested lists, and task lists
* Support Markdown tables and simple aligned-table recognition
* Support hyperlinks, images, and graphical content
* Automatically resize images to prevent them from exceeding Word page boundaries
* Support batch conversion of multiple Markdown files in a specified folder
* Support custom input and output paths
* Support command-line usage in Windows Command Prompt (CMD) and PowerShell
* Support enhanced processing and verification in Microsoft Word COM environments
* Provide both a Windows Installer version and a Portable version
* The Lite version does not include the Playwright browser runtime, resulting in a smaller package size and making it suitable for standard Markdown-to-Word conversion

Default input directory:

`C:\Users\[USERNAME]\Documents\MD_Converter\input`

Place Markdown files to be converted in this folder.

Default output directory:

`C:\Users\[USERNAME]\Documents\MD_Converter\output`

Generated Word documents are saved in this folder.

Basic usage example:

`MD_Converter.exe input/`

The program automatically reads Markdown files from the `input` folder and saves the converted Word documents to the default `output` folder.

MD Converter is designed to provide a **stable, lightweight, and repeatable Markdown-to-Word conversion solution for practical office work and technical documentation workflows**.


1. **Download**

   Download:

   `MD_Converter_v1.1.0_Setup.exe`

   Installer size: approximately **25 MB**

   Disk space used after installation: approximately **70 MB**

2. **Run the Installer**

   Run:

   `C:\Users\[USERNAME]\Downloads\MD_Converter_v1.1.0_Setup.exe`

   No need to create desktop shortcut. 

   Default installation location:

   `C:\Users\[USERNAME]\AppData\Local\Programs\MD_Converter\MD_Converter.exe`

   The installer also creates the following default folders:

   `C:\Users\[USERNAME]\Documents\MD_Converter\input`

   `C:\Users\[USERNAME]\Documents\MD_Converter\output`

4. **Add MD Converter to the PATH Environment Variable**

   Add the following installation directory to the Windows `PATH` environment variable:

   `C:\Users\[USERNAME]\AppData\Local\Programs\MD_Converter`

   Restart the computer after updating `PATH`.

   You can then run MD Converter directly from either **Command Prompt (CMD)** or **PowerShell**.

5. **Usage Example**

   Markdown document:

   `C:\Users\[USERNAME]\Documents\MD_Converter\input\test.md`

   Command:

   `MD_Converter.exe input/`

   The Word document will be generated automatically at:

   `C:\Users\[USERNAME]\Documents\MD_Converter\output\test.docx`

6. **Uninstall**

   Run the following uninstaller:

   `C:\Users\[USERNAME]\AppData\Local\Programs\MD_Converter\unins000.exe`
