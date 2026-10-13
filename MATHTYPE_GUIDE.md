# v1.2 公式转换

窗口提供 MathType 和 Word 原生两种模式。MathType 模式保留现有封面、段落、图表、页眉、编号和引用处理，最后替换公式对象，再刷新目录页码。Word 原生模式适合快速预览。

## 环境

Windows x64、桌面 Microsoft Word、完整 TeX Live。MathType 模式另外需要用户自行安装并授权的 MathType，且 Word 中可用 MathType 的 TeX 转换命令。当前验证环境为 Word 16 与 MathType 7.8、64 位公式 API。Python 和 Pandoc 使用本地运行环境。

v1.2 的路径发现补丁通过只读注册信息、常见目录与手动设置寻找接口，环境检查与公式转换共用同一逻辑。找到接口文件不代表版本兼容、授权或 Word 插件已验证；其他 MathType 7.x 尚未全面验证。手动位置的设置方式见 README。

## 两份样式

- `formula-styles/body.eqp`：正文和独立公式，12 磅、Times New Roman。
- `formula-styles/small.eqp`：表格、图注和表注公式，10.5 磅。

独立公式沿用外部章节编号与书签，双击公式编辑时编号和引用继续由 Word 文档管理。

## 长公式

使用 `equation` 内的 `aligned`、`array` 或 `cases` 明确分行。转换不会自动猜测断行位置；公式超过可用宽度时停止 MathType 替换，保留原生文档供排查。遇到 MathType 不支持的命令，请改写为其支持的 TeX 语法。示例 `long-formula-test.tex` 包含长分数、长根号、四行对齐及六行分段公式。

v1.1.1 将单组 `aligned`（每行一个 `&`）转换为 MathType 关系符号对齐结构，兼容外层 `equation` 和标签。复杂嵌套、多组对齐列与显式 `array` 保留原转换路径，需要另行核对。已生成的历史 Word 文件需重新转换后获得修正。

## 推荐 TeX 写法

这些约定针对本工具的转换路径；使用示例中的写法后仍应核对 PDF 和 Word。公式字号交给两份样式管理，优先使用标准数学命令。

### 段内、表格与图表注

短公式用 `$...$`，例如 `$x_{k+1}$`、`$\frac{a+b}{c+d}$`。表格单元格和图表注也采用段内形式，转换时按位置选择小字号。较高的分数和矩阵放入独立公式；表格内公式应小于单元格可用宽度。

### 独立公式及引用

```tex
\begin{equation}
  E=mc^2
  \label{eq:energy}
\end{equation}
由式\eqref{eq:energy}可得……
```

每个编号对应一个独立 `equation`，标签唯一。通过 `\eqref{...}` 引用，由模板生成章节编号；避免手写编号和 `\tag`。需要多条分别编号的公式时，分别使用多个 `equation`。

### 多行共用一个编号

```tex
\begin{equation}
  \begin{aligned}
    J(\mathbf{x}) &= \frac{1}{2}\|A\mathbf{x}-\mathbf{b}\|_2^2,\\
    \nabla J(\mathbf{x}) &= A^{\mathsf{T}}(A\mathbf{x}-\mathbf{b}),\\
    \mathbf{x}_{k+1} &= \mathbf{x}_k-\eta\nabla J(\mathbf{x}_k).
  \end{aligned}
  \label{eq:iteration}
\end{equation}
```

每行一个 `&`，放在待对齐的关系符号前；用 `\\` 换行。一次排列中优先采用同类关系符号。MathType 按每行第一个关系符号对齐；左侧含其他关系符号时需额外检查。最后一行可省略换行。`\\[长度]` 的自定义行距未承诺与 TeX 完全相同，优先使用默认行距。

过长表达式在 TeX 中手动拆行，检查 Word 中公式和右侧编号均未超出版心。续行没有关系符号的特殊对齐、多个 `&` 对齐组、嵌套 `aligned`、`align` / `split` / `gather`、自定义数学宏及局部字号命令，均需先做小样测试；当前自动对齐修正限定为外层 `equation` 内单组 `aligned`。

### 矩阵与分段

矩阵使用 `bmatrix` / `pmatrix`，分段使用 `cases`；这些环境中的 `&` 用来分列，沿用各自转换路径。常用的分数、根号、积分、求和、上下标已在示例测试。上下标多个字符加花括号，如 `x_{k+1}`、`A^{\mathsf{T}}`；括号和环境成对闭合。

避免用大量空格、`\hspace`、`\phantom` 或零宽字符手工补齐位置。字体优先采用 `\mathbf`、`\mathrm`、`\mathsf`；模板的 `\bm` 会转换为 `\mathbf`，其效果需结合符号检查。

## 测试结果

2026-10-12：四章及附录样例的 21 个公式转换成功，22 页 Word 页面逐页检查；第四章编号为 4-1 等，附录为 A-1。正文小数常量的保留、四章加两个附录的编号映射通过自动测试。513 个公式分六批转换成功，结构检查确认仅替换公式及其资源，其他文字、编号、引用和版式 XML 保留。批量测试的页面未逐页验收。

具体耗时取决于电脑及公式复杂度；窗口提供公式转换进度。请串行运行转换，日常编辑保持在 TeX，再重新生成 Word。

## 发布内容

源码仓库提供转换代码、两份自建样式和示例。MathType、Word、TeX Live 由用户自行安装。原稳定版本独立保留。
