# ImportBMP2Excel

**English** | [简体中文](#简体中文)

A Microsoft Excel macro to import 24-bit BMP image files and display them in Excel by coloring cells.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

## Description

`ImportBMP2Excel.xlsm` contains a VBA macro that allows users to import 24-bit BMP (Bitmap) image files into Microsoft Excel and represent the image by coloring the background of Excel cells.

**Key Features:**

* **BMP File Selection:** Uses the standard Excel file dialog to select the BMP file to import (`\*.bmp`, `\*.dib`).
* **24-bit BMP Support:** Specifically designed for 24-bit BMP image files.
* **Pixel-by-Pixel Representation:** Represents each pixel of the BMP image by setting the background color of the corresponding cell in the Excel worksheet.
* **File Size Limit:** Limits the BMP file size to a maximum of 900 KB to ensure Excel performance.
* **Error Handling:** Displays error messages for invalid file types (non-24-bit BMP) or files exceeding the size limit.
* **Import Button:** Provides an "Import" button in the Excel sheet to trigger the BMP import process.
* **Clear Button:** Provides a "Clear" button in the Excel sheet to clear all imported data.

## Installation

1.  **Download Macro File:** Download the `ImportBMP2Excel.xlsm` file from the GitHub repository.
2.  **Open Excel:** Open Microsoft Excel.
3.  **Open Macro File:** In Excel, go to "File" -> "Open" and locate the downloaded `ImportBMP2Excel.xlsm` file.
4.  **Enable Macros:** When Excel prompts you to enable macros, choose "Enable Content" or allow macros to run.

## How to Use

1.  Open the workbook containing the macro (`ImportBMP2Excel.xlsm`).
2.  **Importing BMP Files:**
    * Click the "Import" button in the Excel sheet.
    * A file dialog will appear, prompting you to select a 24-bit BMP image file (`\*.bmp` or `\*.dib`).
    * After selecting the file, the macro will fill the Excel worksheet with cell colors, starting from column C and row 1. The image dimensions (width and height) determine the range of cells filled.
    * If the file is not a 24-bit BMP or exceeds 900 KB, an error message will be displayed.
    * A "Done!" message box appears upon successful completion.
3.  **Clearing Imported Data:**
    * Click the "Clear" button in the Excel sheet.
    * This will clear all cell formatting and contents from the active worksheet.

## License

This project is dual-licensed under either of the following licenses:

* **MIT License:** See the [LICENSE](LICENSE) file for details.
* **GNU General Public License v3.0 (GPLv3):** See the [LICENSE-GPL-3.0](LICENSE-GPL-3.0) file for details.

Choose the license that best suits your needs. In brief:

* The **MIT License** allows you to use the code in almost any way, as long as you retain the copyright notice and permission notice. It is more permissive and suitable for developers who want to integrate the code into commercial software.
* The **GPLv3 License** is a stronger copyleft license, requiring any work based on your code to also be open-sourced under the GPLv3 license. It aims to protect the open-source nature of the code.

We provide dual licensing to make the code available to a wider audience with different needs.

## Contributing

Contributions are welcome in any form, including but not limited to:

* Bug reports
* Feature requests (e.g., support for other BMP bit depths, custom import location)
* Bug fixes
* Feature implementation
* Code efficiency improvements
* Documentation improvements
* Translations

If you would like to contribute code, please fork the repository, make your changes, and submit a pull request.

## Contact

If you have any questions, suggestions, or feedback, please feel free to contact me via GitHub Issues or email at gsyjz@hotmail.com.

## Acknowledgements

Thanks to everyone who contributes to the open-source community.

