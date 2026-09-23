# Function & Features
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

# Usage examples
Below is the complete English translation, with the original structure and command examples preserved.

Below are some **common MD Converter usage examples**. Assuming the Windows-installed version has already been added to `PATH`, you can use:

```text
MD_Converter.exe
```

If you are using the Python package version, the corresponding command is usually:

```text
md-converter
```

The usage of the two versions is essentially the same.

1. **Batch convert files using the default `input` / `output` directories**

   Current directory:

   ```text
   C:\Users\[USERNAME]\Documents\projects\Work_Directory\
   ├── input\
   │   ├── report.md
   │   └── notes.md
   └── output\
   ```

   Run:

   ```powershell
   MD_Converter.exe
   ```

   This is equivalent to:

   ```powershell
   MD_Converter.exe input/
   ```

   Output:

   ```text
   output\
   ├── report.docx
   └── notes.docx
   ```

2. **Convert a single Markdown file**

   ```powershell
   MD_Converter.exe README.md
   ```

   The corresponding DOCX file is generated in the default output directory.

3. **Specify a single output file**

   ```powershell
   MD_Converter.exe README.md `
       --output ".\output\README.docx"
   ```

   Short form:

   ```powershell
   MD_Converter.exe README.md `
       -o ".\output\README.docx"
   ```

4. **Batch convert Markdown files in a specified directory**

   ```powershell
   MD_Converter.exe ".\docs"
   ```

   For example, if the directory contains:

   ```text
   docs\
   ├── architecture.md
   ├── installation.md
   └── user_guide.md
   ```

   MD Converter can batch-generate the corresponding Word documents.

5. **Call MD Converter from another project**

   For example, from:

   ```text
   C:\Users\[USERNAME]\Documents\projects\Work_Directory
   ```

   Run:

   ```powershell
   MD_Converter.exe input/
   ```

   MD Converter will use:

   ```text
   Work_Directory\input
   Work_Directory\output
   ```

   This is especially useful when using MD Converter as a general-purpose document conversion tool shared by other projects.

6. **Call the installed version directly without relying on `PATH`**

   ```powershell
   & "$env:LOCALAPPDATA\Programs\MD_Converter\MD_Converter.exe" `
       ".\input"
   ```

   This approach is particularly suitable for scripts, automation workflows, and calls from other applications.

7. **Call the Portable version directly**

   If MD Converter is not installed:

   ```powershell
   & "D:\Tools\MD_Converter_Lite\MD_Converter_Lite.exe" `
       ".\input"
   ```

   The Portable version does not require Python to be installed.

8. **Automatically open the Word file after conversion**

   ```powershell
   MD_Converter.exe report.md --open
   ```

   If Microsoft Word or another default DOCX application is installed, the generated file will be opened automatically.

9. **Prevent the generated file from opening automatically**

   ```powershell
   MD_Converter.exe report.md --no-open
   ```

   This is useful for batch processing and automated workflows.

10. **Use a custom YAML configuration**

    ```powershell
    MD_Converter.exe report.md `
        --config config.yaml
    ```

    Or:

    ```powershell
    MD_Converter.exe report.md `
        -c config.yaml
    ```

    This is suitable when you need to specify a theme, styles, or other conversion parameters.

11. **Specify both configuration and output paths**

    ```powershell
    MD_Converter.exe report.md `
        --config ".\config.yaml" `
        --output ".\build\report.docx"
    ```

    This is a common pattern in project build workflows:

    ```text
    source Markdown
          ↓
    MD Converter
          ↓
    build\report.docx
    ```

12. **Display detailed diagnostic information**

    ```powershell
    MD_Converter.exe report.md --verbose
    ```

    Short form:

    ```powershell
    MD_Converter.exe report.md -v
    ```

    This is useful for troubleshooting images, tables, Mermaid rendering, Word COM processing, or other conversion issues.

13. **Automatically convert ASCII diagrams**

    ```powershell
    MD_Converter.exe architecture.md `
        --ascii-mode auto
    ```

    Currently supported modes:

    ```text
    auto
    interactive
    preview
    ```

14. **Generate ASCII → Mermaid preview files**

    ```powershell
    MD_Converter.exe architecture.md `
        --ascii-mode preview `
        --ascii-preview-dir ".\preview"
    ```

    This is useful for checking the recognition results of complex ASCII diagrams.

15. **Display command-line help**

    ```powershell
    MD_Converter.exe --help
    ```

    Python package version:

    ```powershell
    md-converter --help
    ```

16. **Use MD Converter in a PowerShell automation script**

    ```powershell
    $converter = "$env:LOCALAPPDATA\Programs\MD_Converter\MD_Converter.exe"

    & $converter `
        ".\docs" `
        --verbose
    ```

    This allows the script to operate independently of the current Python virtual environment.

17. **Use the standalone MD Converter from Work_Directory**

    This is particularly suitable for your use case:

    ```text
    Work_Directory
        │
        ├── input\
        ├── output\
        │
        └── call
             ↓
    %LOCALAPPDATA%\Programs\MD_Converter\
        MD_Converter.exe
    ```

    PowerShell:

    ```powershell
    cd C:\Users\[USERNAME]\Documents\projects\Work_Directory

    & "$env:LOCALAPPDATA\Programs\MD_Converter\MD_Converter.exe"
    ```

    MD Converter will then use:

    ```text
    Work_Directory\input
    Work_Directory\output
    ```

    At the same time, it will be completely independent of:

    ```text
    Work_Directory\.venv
    ```

    whether that virtual environment is rebuilt, moved, or becomes corrupted.

For practical day-to-day use, the three most common patterns are:

```powershell
# 1. Batch-convert files in the current project's default directories
MD_Converter.exe

# 2. Convert a single file
MD_Converter.exe report.md

# 3. Use the fixed installed version in project automation
& "$env:LOCALAPPDATA\Programs\MD_Converter\MD_Converter.exe" ".\input"
```

The third approach is particularly suitable when using **MD Converter as a standalone Windows tool shared by other projects such as Work_Directory and QCFP_MTF**.

# Installation
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

# Hardware & OS requirements
## Windows Compatibility Requirements

**Supported Operating Systems**

* Windows 10, 64-bit
* Windows 11, 64-bit

**Architecture**

* x64 (64-bit)
* 32-bit Windows is not supported.
* Windows on ARM has not been officially tested and is not currently supported.

## Hardware Requirements

**Processor**

* 64-bit x64-compatible processor
* Dual-core processor or better recommended

**Memory**

* Minimum: 4 GB RAM
* Recommended: 8 GB RAM or more, especially when processing large Markdown documents, images, or complex Word output

**Disk Space**

* At least 200 MB of free disk space for installation and normal operation
* Additional disk space is required for Markdown source files, generated DOCX documents, images, temporary files, and logs

**Display**

* No special graphics hardware is required for standard Markdown-to-DOCX conversion
* A standard Windows-compatible display is sufficient

**Internet Connection**

* Not required for standard Markdown-to-DOCX conversion
* Some optional or externally dependent features may require Internet access

**Microsoft Word**

* Microsoft Word is not required for basic Markdown-to-DOCX conversion
* The desktop version of Microsoft Word is required for enhanced features that use Word COM automation

## Python

Python does not need to be installed separately.

The required Python runtime and application dependencies are bundled with MD Converter.

## Microsoft Word

Microsoft Word is not required for basic Markdown-to-DOCX conversion.

Some enhanced features that use Microsoft Word COM automation require the desktop version of Microsoft Word to be installed.

## Lite Edition

The Windows Lite edition does not bundle Playwright or Chromium.

Standard Markdown-to-Word conversion is supported, while browser-based Mermaid rendering may not be available in the Lite edition.

## Legacy Windows Versions

Windows 7 is not supported.

Windows 8.1 may be technically compatible with the underlying Python 3.12 runtime, but it has not been officially tested and is therefore not included in the supported operating systems.

## Recommended Environment

* Windows 10 or Windows 11
* 64-bit x64 processor
* 8 GB RAM or more
* At least 200 MB of available disk space
* Microsoft Word recommended if Word COM-based features are required
* Internet access optional for standard Markdown-to-DOCX conversion

# License

MD Converter is proprietary software.

**License Type:** Proprietary Single-User End User License Agreement (EULA)

Each license is granted to one individual user. The software may not be
shared, redistributed, resold, sublicensed, or used as a multi-user,
team-wide, organization-wide, hosted, or SaaS service without separate
authorization.

Personal and business use by the licensed individual is permitted.

Third-party components remain subject to their respective licenses.

