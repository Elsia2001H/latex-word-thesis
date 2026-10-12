# 随包依赖

运行环境：CPython 3.12.14 Windows x64（runtime/LICENSE.txt）。

转换依赖：PyMuPDF 1.24.11、python-docx 1.2.0、lxml 6.1.3、pandoc-tex-numbering 1.3.4、panflute 2.3.1、pylatexenc 2.11、PyYAML 6.0.3、click 8.5.0、typing_extensions 4.16.0。

各包的许可证及元数据随 runtime/Lib/site-packages 内的包目录、dist-info 目录保留。Pandoc 的版本由 tools/pandoc.exe --version 查询，其许可证见 tools/COPYING.rtf 和 tools/COPYRIGHT.txt。

PyMuPDF 采用 AGPL/商业双许可，Pandoc 采用 GPL。此整理包用于本地科研工作；如进一步公开再分发或商业部署，请遵守所附许可证，并评估对应源代码提供、许可声明及商业许可需求。

Microsoft Word、TeX Live 和 MathType 由用户自行安装。本仓库的两份公式样式为项目自行编写的设置文件。公开源码包含代码、模板及测试示例；Python、Pandoc 和其依赖通过本地构建脚本准备。
