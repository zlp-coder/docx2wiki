## docx2wiki
jar包，用于转换 word 文档为 wiki 格式，包括文字、图形和表格。支持文字标题的转换，支持嵌入图形转换，支持有单元格合并的复杂表格。

## 用法

下载 target 中编译好的jar包(docx2wiki-1.0.jar)，或者下载源代码自己编译

需要将pom文件中的引用加入到主工程文件中

调用的例子在 op.tools.docx2wiki.java


### 已知问题

* 仅支持docx，不支持doc
* word 文档中如果有图片是emf格式，并且有中文的，转换后的图片中文乱码
