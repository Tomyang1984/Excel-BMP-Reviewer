#   ImportBMP2Excel

[English](README.md)

一个 Microsoft Excel 宏，用于导入 24 位 BMP 图像文件，并通过为单元格着色在 Excel 中显示图像。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

##   描述

`ImportBMP2Excel.xlsm` 包含一个 VBA 宏，它允许用户将 24 位 BMP (Bitmap) 图像文件导入到 Microsoft Excel 中，并通过设置 Excel 单元格的背景颜色来表示图像。

**主要功能:**

* **BMP 文件选择:** 使用标准的 Excel 文件对话框选择要导入的 BMP 文件 (`\*.bmp`, `\*.dib`)。
* **24 位 BMP 支持:** 专门为 24 位 BMP 图像文件设计。
* **逐像素表示:** 通过设置 Excel 工作表中相应单元格的背景颜色来表示 BMP 图像的每个像素。
* **文件大小限制:** 为了确保 Excel 的性能，将 BMP 文件大小限制为最大 900 KB。
* **错误处理:** 显示无效文件类型（非 24 位 BMP）或超出大小限制的文件的错误消息。
* **导入按钮:** 在 Excel 工作表中提供一个“导入”按钮，用于触发 BMP 导入过程。
* **清除按钮:** 在 Excel 工作表中提供一个“清除”按钮，用于清除所有导入的数据。

##   安装

1.  **下载宏文件:** 从 GitHub 存储库下载 `ImportBMP2Excel.xlsm` 文件。
2.  **打开 Excel:** 打开 Microsoft Excel。
3.  **打开宏文件:** 在 Excel 中，转到“文件” -> “打开”，然后找到下载的 `ImportBMP2Excel.xlsm` 文件。
4.  **启用宏:** 当 Excel 提示启用宏时，选择“启用内容”或允许宏运行。

##   如何使用

1.  打开包含宏的工作簿 (`ImportBMP2Excel.xlsm`)。
2.  **导入 BMP 文件:**
    * 单击 Excel 工作表中的“导入”按钮。
    * 将出现一个文件对话框，提示您选择一个 24 位 BMP 图像文件 (`\*.bmp` 或 `\*.dib`)。
    * 选择文件后，宏将使用单元格颜色填充 Excel 工作表，从 C 列和第 1 行开始。图像尺寸（宽度和高度）确定填充的单元格范围。
    * 如果文件不是 24 位 BMP 或超过 900 KB，则会显示一条错误消息。
    * 成功完成后，将显示一个“完成！”消息框。
3.  **清除导入的数据:**
    * 单击 Excel 工作表中的“清除”按钮。
    * 这将清除活动工作表中的所有单元格格式和内容。

##   许可证

本项目采用以下任一双重许可证：

* **MIT 许可证:** 有关详细信息，请参见 [LICENSE](LICENSE) 文件。
* **GNU 通用公共许可证 v3.0 (GPLv3):** 有关详细信息，请参见 [LICENSE-GPL-3.0](LICENSE-GPL-3.0) 文件。

选择最适合您需求的许可证。 简而言之：

* **MIT 许可证** 允许您以几乎任何方式使用代码，只要您保留版权声明和权限声明即可。 它更宽松，适用于希望将代码集成到商业软件中的开发人员。
* **GPLv3 许可证** 是一种更强的 copyleft 许可证，要求任何基于您的代码的作品也必须根据 GPLv3 许可证开源。 它旨在保护代码的开源性质。

我们提供双重许可，以便更广泛的用户可以根据不同的需求使用该代码。

##   贡献

欢迎任何形式的贡献，包括但不限于：

* Bug 报告
* 功能请求（例如，支持其他 BMP 位深度、自定义导入位置）
* Bug 修复
* 功能实现
* 代码效率提升
* 文档改进
* 翻译

如果您想贡献代码，请 fork 该存储库，进行更改并提交拉取请求。

##   联系方式

如果您有任何问题、建议或反馈，请随时通过 GitHub Issues 或发送电子邮件至 gsyjz@hotmail.com 与我联系。

##   致谢

感谢所有为开源社区做出贡献的人们。
