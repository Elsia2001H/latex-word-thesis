# v1.1 公式转换

窗口提供 MathType 和 Word 原生两种模式。MathType 模式保留现有封面、段落、图表、页眉、编号和引用处理，最后替换公式对象，再刷新目录页码。Word 原生模式适合快速预览。

## 环境

Windows x64、桌面 Microsoft Word、完整 TeX Live。MathType 模式另外需要用户自行安装并授权的 MathType，且 Word 中可用 MathType 的 TeX 转换命令。当前验证环境为 Word 16 与 MathType 7.8、64 位公式 API。Python 和 Pandoc 使用本地运行环境。

## 两份样式

- `formula-styles/body.eqp`：正文和独立公式，12 磅、Times New Roman。
- `formula-styles/small.eqp`：表格、图注和表注公式，10.5 磅。

独立公式沿用外部章节编号与书签，双击公式编辑时编号和引用继续由 Word 文档管理。

## 长公式

使用 `equation` 内的 `aligned`、`array` 或 `cases` 明确分行。转换不会自动猜测断行位置；公式超过可用宽度时停止 MathType 替换，保留原生文档供排查。遇到 MathType 不支持的命令，请改写为其支持的 TeX 语法。示例 `long-formula-test.tex` 包含长分数、长根号、四行对齐及六行分段公式。

## 测试结果

2026-10-12：四章及附录样例的 21 个公式转换成功，22 页 Word 页面逐页检查；第四章编号为 4-1 等，附录为 A-1。正文小数常量的保留、四章加两个附录的编号映射通过自动测试。513 个公式分六批转换成功，结构检查确认仅替换公式及其资源，其他文字、编号、引用和版式 XML 保留。批量测试的页面未逐页验收。

具体耗时取决于电脑及公式复杂度；窗口提供公式转换进度。请串行运行转换，日常编辑保持在 TeX，再重新生成 Word。

## 发布内容

源码仓库提供转换代码、两份自建样式和示例。MathType、Word、TeX Live 由用户自行安装。原稳定版本独立保留。
