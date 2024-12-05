# PPTX-to-Image-format-PPTX-or-PDF

## 项目介绍

PPTX-to-Image-format-PPTX-or-PDF 是一个用于将 PPTX 文件转换为高分辨率图像，并根据需要将这些图像合成为 PDF 文件或新的 PPTX 文件的工具。

## 功能说明

- **pptx-to-image_PDF.py**：将 PPTX 文件的每张幻灯片导出为高分辨率的 PNG 图像，并将这些图像合成为一个 PDF 文件。
- **pptx-to-image_PPTX.py**：将 PPTX 文件的每张幻灯片导出为高分辨率的 PNG 图像，然后将这些图像插入到一个新的 PPTX 文件中，每张图像作为一张幻灯片。

## 环境依赖

- **操作系统**：仅支持 Windows
- **Python 版本**：Python 3.x
- **需要安装的库**：
  - `pywin32`
  - `Pillow`（PIL）
- **软件要求**：
  - Microsoft Office PowerPoint

## 安装依赖

使用以下命令安装所需的 Python 库：

```Bash
pip install pywin32 Pillow
```

## 使用方法
- 确保已安装 Microsoft Office PowerPoint。
- 运行脚本：

1. 对于生成 PDF：
```Bash
python pptx-to-image_PDF.py
```

2. 对于生成图片版 PPTX：
```Bash
python pptx-to-image_PPTX.py
```
3. 按照提示输入要转换的 PPTX 文件路径和所需的分辨率（DPI）。

## 注意事项
脚本仅在 Windows 系统上运行，因为使用了 Windows COM 接口来控制 PowerPoint。
请确保安装了兼容的 Microsoft Office PowerPoint 版本。
运行脚本前，请关闭所有打开的 PowerPoint 应用程序，以避免可能的冲突。

## 许可协议
本项目使用 MIT 许可证进行许可。
