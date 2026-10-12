# LaTeX Word Thesis

中文学位论文 LaTeX 转 PDF / Word 工具。使用同一份 TeX 编辑源，生成可编辑的 Word 文档，并尽量保持论文版式一致。

## 功能

- 原生可编辑封面、中文与英文摘要、目录、章节、参考文献、科研成果和致谢。
- MathType 可编辑公式与 Word 原生快速预览，按实际章节编号并保留交叉引用。
- PDF 图片自动转高分辨率 PNG，支持 PNG、JPG 和带居中子图标识的多图排版。
- 中英文图注和表注、三线表、楷体页眉及横线。
- Bib 文献按首次引用顺序编号，正文引用为上标。
- 简单窗口选择根目录内的 TeX，输出文件名附带时间戳，保留历次结果。
- 修正中英文混排时 Pandoc 首段样式错误覆盖中文字体的问题。
- 两份公式样式：正文/独立公式 12 磅，表格/图表注公式 10.5 磅；批量转换进度显示。
- 支持单组 `aligned` 多行公式的关系符号对齐，清理公式和图表引用旁的多余编辑标记。
- v1.2：目录一级条目使用黑体，声明与授权书同页、后接空白页；正文上标文献引用可跳转到末尾对应文献。

## 源码版与便携版的区别

本仓库提供核心源码、模板和示范。下载源码后，先运行构建脚本准备 Python 和 Pandoc 运行环境，再启动转换窗口。

本地构建完成后，整个文件夹可迁移至其他 Windows x64 电脑，无需另装 Python 和 Pandoc；目标电脑仍需完整 TeX Live、桌面 Microsoft Word，以及宋体、黑体、楷体、Times New Roman 等字体。

## 首次构建

构建电脑需安装 Python 3.12 x64（包含 pip、Tcl/Tk）、Pandoc、完整 TeX Live 和桌面 Microsoft Word。TeX Live 的 `bin/windows` 应加入 PATH。

下载源码 ZIP 并解压，或者克隆本仓库。在 PowerShell 中进入解压目录，运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\prepare_portable.ps1 `
  -PythonRoot 'C:\你的Python312安装目录' `
  -PandocExe 'C:\Program Files\Pandoc\pandoc.exe'
```

将 PythonRoot 替换为包含 python.exe 的真实目录。脚本在仓库内复制运行环境、联网安装固定版本依赖；执行选项仅对当前 PowerShell 进程生效。

完成后双击 `启动转换.bat`，选择 `thesis-template.tex` 并开始转换。PDF、DOCX 和日志保存于 `output`。

窗口默认使用 MathType；此模式需要自行安装、授权 MathType，并确保其 Word TeX 转换命令可用。本机验证为 MathType 7.8 与 64 位 API。选择 Word 原生模式可快速预览。详见 `MATHTYPE_GUIDE.md`。

命令行也可运行：

```powershell
.\runtime\python.exe .\program\launcher.py --check
.\runtime\python.exe .\program\launcher.py --convert .\thesis-template.tex
.\runtime\python.exe .\program\launcher.py --convert .\thesis-template.tex --math-mode mathtype
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
| `formula-styles/` | 两份自建 MathType 公式样式 |
| `long-formula-test.tex` | 四章、附录及长公式回归样例 |
| `MATHTYPE_GUIDE.md` | MathType 环境、长公式处理和测试说明 |

## 适用范围

适用于本仓库论文模板。保持前置页结构和模板命令，图片放 `figures`，文献放 `references.bib`；复制 TeX 模板后可以改名使用。请使用来源可信的 TeX 文档。

Word 和 TeX 使用各自的排版引擎，分页和换行可能存在差异。修改以 TeX 为主，再重新导出 Word。

章节编号目前读取主 TeX 文件中的 `\chapter` 与 `\appendix`；使用自定义章节宏或拆分章节文件时，需要先展开到主文件。长公式使用 `aligned`、`array`、`cases` 显式分行。超宽或不支持的公式会中止 MathType 替换并保留原生文档，方便检查。

本机 Windows x64、TeX Live 2026、Pandoc 3.11、桌面 Word 16.0 完成测试。包含中文和空格的迁移路径已验证；v1.2 主示范为 20 页、13 个 MathType 公式，长公式示范为 22 页、21 个 MathType 公式。15 项自动回归测试通过，文献引用保持黑色上标；主示范单独复制后，21 处文献链接通过 Word 实际跳转检查。

## 任务分工

- **Elsia2001H**：需求提出、参考资料与测试素材提供、试用反馈和效果验收。
- **GPT（通过 OpenAI Codex）**：方案设计、代码实现、模板调整、排版调试、回归测试、便携环境整理和 GitHub 发布。

## 许可

项目代码的许可证由仓库所有者后续确定。再分发或商业使用前，请联系仓库所有者确认授权。

第三方依赖遵循各自的许可证。运行环境通过本地构建脚本准备；公开再分发前，请阅读 `THIRD_PARTY.md`，保留许可声明并满足相应要求。
