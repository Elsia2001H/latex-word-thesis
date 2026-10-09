# LaTeX Word Thesis

中文学位论文 LaTeX 转 PDF / Word 工具。使用同一份 TeX 编辑源，生成可编辑的 Word 文档，并尽量保持论文版式一致。

## 功能

- 原生可编辑封面、中文与英文摘要、目录、章节、参考文献、科研成果和致谢。
- 原生 Word 公式、按章编号与交叉引用；不依赖 MathType。
- PDF 图片自动转高分辨率 PNG，支持 PNG、JPG 和带居中子图标识的多图排版。
- 中英文图注和表注、三线表、楷体页眉及横线。
- Bib 文献按首次引用顺序编号，正文引用为上标。
- 简单窗口选择根目录内的 TeX，输出文件名附带时间戳，不覆盖旧结果。
- 修正中英文混排时 Pandoc 首段样式错误覆盖中文字体的问题。

## 源码版与便携版的区别

本仓库发布核心源码、模板和示范，不包含第三方 Python/Pandoc 二进制运行环境。直接下载源码后需要先构建本地便携环境，不能直接双击启动。

本地构建完成后，整个文件夹可迁移至其他 Windows x64 电脑，无需另装 Python 和 Pandoc；目标电脑仍需完整 TeX Live、桌面 Microsoft Word，以及宋体、黑体、楷体、Times New Roman 等字体。

## 首次构建

构建电脑需安装 Python 3.12 x64（包含 pip、Tcl/Tk）、Pandoc、完整 TeX Live 和桌面 Microsoft Word。TeX Live 的 `bin/windows` 应加入 PATH。

下载源码 ZIP 并解压，或者克隆本仓库。在 PowerShell 中进入解压目录，运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\prepare_portable.ps1 `
  -PythonRoot 'C:\你的Python312安装目录' `
  -PandocExe 'C:\Program Files\Pandoc\pandoc.exe'
```

将 PythonRoot 替换为包含 python.exe 的真实目录。脚本只在仓库内复制运行环境、安装固定版本依赖，不修改系统执行策略；需要联网下载 Python 依赖。

完成后双击 `启动转换.bat`，选择 `thesis-template.tex` 并开始转换。PDF、DOCX 和日志保存于 `output`。

命令行也可运行：

```powershell
.\runtime\python.exe .\program\launcher.py --check
.\runtime\python.exe .\program\launcher.py --convert .\thesis-template.tex
```

## 目录

| 路径 | 内容 |
| --- | --- |
| `program/` | 核心转换程序、小窗口与 Word 字段更新脚本 |
| `thesis-template.tex` | 可复用论文模板 |
| `thesis-reference.docx` | Word 样式模板 |
| `references.bib` | 示例文献库 |
| `numeric-citation.csl` | Word 文献著录与引用样式 |
| `figures/` | PDF、PNG、JPG 与多图示例 |
| `examples/` | 稳定版本的 Word 和 PDF 示例 |
| `REFERENCE_GUIDE.md` | 文献维护说明 |
| `THIRD_PARTY.md` | 第三方依赖及许可说明 |

## 适用范围

针对本仓库论文模板，不是任意 LaTeX 项目的通用转换器。保持前置页结构和模板命令，图片放 `figures`，文献放 `references.bib`；复制 TeX 模板后可以改名使用。不要转换不可信来源的 TeX。

Word 和 TeX 的排版引擎不同，不保证分页或每一行逐像素一致。修改应以 TeX 为主，重新导出 Word。

本机 Windows x64、TeX Live 2026、Pandoc 3.11、桌面 Word 16.0 完成测试。包含中文和空格的迁移路径已验证；混排修正版 Word 示例为 20 页，13 个公式对象，8 处公式引用目标全部存在，4 张图片、7 条参考文献。

## 许可

本项目暂未指定自有代码的开源许可证；公开可见不等于授予任意再分发或商业使用许可。后续可由仓库所有者选择许可证。

第三方依赖各自遵循原许可证。本仓库不上传第三方二进制便携包；自行构建或公开再分发时请阅读 `THIRD_PARTY.md`，保留许可声明并满足相应许可证要求。
