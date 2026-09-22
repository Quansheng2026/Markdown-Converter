1. **Download**

   Download:

   `MD_Converter_v1.1.0_Setup.exe`

   Installer size: approximately **25 MB**

   Disk space used after installation: approximately **70 MB**

2. **Run the Installer**

   Run:

   `C:\Users\[USERNAME]\Downloads\MD_Converter_v1.1.0_Setup.exe`

   Default installation location:

   `C:\Users\[USERNAME]\AppData\Local\Programs\MD_Converter\MD_Converter.exe`

   The installer also creates the following default folders:

   `C:\Users\[USERNAME]\Documents\MD_Converter\input`

   `C:\Users\[USERNAME]\Documents\MD_Converter\output`

3. **Add MD Converter to the PATH Environment Variable**

   Add the following installation directory to the Windows `PATH` environment variable:

   `C:\Users\[USERNAME]\AppData\Local\Programs\MD_Converter`

   Restart the computer after updating `PATH`.

   You can then run MD Converter directly from either **Command Prompt (CMD)** or **PowerShell**.

4. **Usage Example**

   Markdown document:

   `C:\Users\[USERNAME]\Documents\MD_Converter\input\test.md`

   Command:

   `MD_Converter.exe input/`

   The Word document will be generated automatically at:

   `C:\Users\[USERNAME]\Documents\MD_Converter\output\test.docx`

5. **Uninstall**

   Run the following uninstaller:

   `C:\Users\[USERNAME]\AppData\Local\Programs\MD_Converter\unins000.exe`
