# Typst docx

这是一个用于在WPS Office中直接使用Typst的服务端脚本。

## 使用方法

```bash
cargo install --git https://github.com/yyhhenry/typst-docx
```

保证电脑中已经安装了pandoc（添加到PATH），Rust环境，以及WPS Office。

在WPS Office的Normal.dotm中添加宏，然后根据需求添加到快捷访问工具栏（笔者使用Alt+7）。

本地运行此项目的Rust后端，在WPS Office中直接输入一段Typst源码，选中后Alt+7即可将其转换为对应的docx格式。

主要用于输入公式，或借助Typst脚本进行一定的计算。

请尽量以段落为单位使用，默认情况下，此脚本会嵌入一个段落（末尾有换行符），如果你分别选中每个词，每次都需要手动删除多余的换行符。

## Word

可以尝试scripts/macro.vba中的宏PasteTypstDocx，将其添加到Normal.dotm中，然后在Word中加入快捷访问工具栏。

这个宏与上述的WPS Office中的JS宏有略微的不同，不是作为一个转换功能，而是根据剪贴板中的Typst源码，粘贴为docx格式。
