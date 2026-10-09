# 参考文献维护说明

文献库是 `references.bib`，正文只写 `\cite{键名}`，不要手动写编号。
PDF 使用 Biber 和 biblatex-gb7714-2015，Word 使用 Pandoc citeproc 和
`numeric-citation.csl`。两条路径都按首次引用顺序编号，重复引用不增加编号。
改动 bib 后请运行完整构建 `py build_thesis.py`，同时刷新 PDF 与 Word。

## 本模板的作者格式

以原 Word 模板为格式依据：英文姓氏完整保留原拼写和大小写，姓在前，
名用大写首字母，多个首字母之间留空格，不加缩写点。例如 `Kalman R E`、
`Åström K J`、`LeCun Y`。不要把所有姓改成全大写，也不要破坏 LeCun 等原始拼写。
中文姓名完整显示，不缩写。超过三名英文作者显示前三名再加 `et al.`。
这属于参考模板的姓名显示约定，不等于宣称严格照搬国标所有默认选项。

bib 中用 `author = {Family, Given names and Family, Given names}`，
中文单作者用 `author = {{胡寿松}}`。姓名的缩写由样式完成，不要混写成
`KALMAN R E`、`R. E. Kalman` 和 `Kalman, Rudolf Emil` 三种显示形式。
未知的中间名不要猜写全，保留出版物提供的首字母。

## 文献类型与样例

- `@book`：`[M]`，含出版地、出版社、年份及非首版版次。
- `@article`：`[J]`，含刊名、年份、卷(期)、页码。
- `@inproceedings`：`[C]//`，含会议文集、出版信息、年份与页码。
- DOI 和来源 URL 保存在库中供核查，本模板不在打印著录中显示。

样例预期首次引用顺序为：胡寿松、Åström 与 Murray、周志华、LeCun 等、
He 等、Kalman、Achenbach。bib 文件故意不是这个顺序，以测试自动排序。

## 书目信息核查来源

1. 胡寿松，自动控制原理，第七版，科学出版社，2019。
   https://www.ecsponline.com/goods.php?id=198822
2. Åström K J, Murray R M. Feedback Systems. Princeton University Press, 2008.
   https://www.cds.caltech.edu/~murray/FBS/First_Edition.html
3. 周志华，机器学习，清华大学出版社，2016。
   https://cs.nju.edu.cn/zhouzh/zhouzh.files/publication/MLbook2016.htm
4. LeCun Y, Bengio Y, Hinton G. Deep learning. Nature, 2015, 521: 436-444.
   https://www.nature.com/articles/nature14539
5. He K, Zhang X, Ren S, et al. Deep Residual Learning for Image Recognition.
   CVPR, 2016: 770-778.
   https://openaccess.thecvf.com/content_cvpr_2016/html/He_Deep_Residual_Learning_CVPR_2016_paper.html
6. Kalman R E. A New Approach to Linear Filtering and Prediction Problems.
   Journal of Basic Engineering, 1960, 82(1): 35-45.
   https://doi.org/10.1115/1.3662552
7. Achenbach J D. Wave Propagation in Elastic Solids. North-Holland, 1973.
   https://www.sciencedirect.com/book/monograph/9780720403251/wave-propagation-in-elastic-solids

## 图表间距

图片前 18 磅不变，英文图注后 12 磅；表格整体前后 12 磅。
图片与中文图注之间保留 6 磅，双语图注之间不另加段距。
TeX 浮动图表与 Word 流式图表的分页方式不同，间距一致不保证逐页相同。
