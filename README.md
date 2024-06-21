# Paste Typst docx

这是一个用于在WPS/Word中直接粘贴Typst内容的宏。

This is a macro for pasting Typst content directly into WPS/Word.

## 使用方法 Usage

保证电脑中已经安装了pandoc（添加到PATH），Rust环境，以及WPS/Word。

Make sure you have pandoc (added to PATH), Rust environment, and WPS/Word installed on your computer.

```bash
# install pandoc and rust
scoop install pandoc rustup

# install typst-docx
cargo install --git https://github.com/yyhhenry/typst-docx
```

在Normal.dotm中添加[Word VBA宏](scripts/macro.vba)或，然后根据需求添加到快捷访问工具栏。

Add [Word VBA macro](scripts/macro.vba) or [WPS JS macro](scripts/macro.js) to Normal.dotm, then add it to the Quick Access Toolbar as needed.

如果调用宏时选中了一些内容，会将选中的内容剪切。

If some content is selected when the macro is called, the selected content will be cut.

后端会编译剪切板中的Typst源码到Word内容，并由宏插入到当前光标处。样式会来自插入前光标处的样式，如果获取的样式不在样式库中，则会使用默认样式。

The backend will compile the Typst source code in the clipboard into Word content, and the macro will insert it at the current cursor position. The style will come from the style at the cursor before insertion. If the obtained style is not in the style library, the default style will be used.

运行宏时，如果后台未启动，会自动启动后台。

When running the macro, if the backend is not started, the backend will be started automatically.

## WPS兼容模式 WPS Compatibility Mode

使用[WPS JS宏](scripts/macro.js)，其他设置方法与上述相同。

Use [WPS JS macro](scripts/macro.js), other settings are the same as above.

只能使用剪切板，不会自动剪切选中内容。

Only the clipboard can be used, and the selected content will not be cut automatically.

默认情况下会生成额外的末尾换行，当你在段落的中间插入Typst内容时，会断开段落。

By default, extra line breaks will be generated at the end, which will break the paragraph when you insert Typst content in the middle of a paragraph.

默认情况下会插入默认正文样式，请手动调整。

By default, the default body style will be inserted, please adjust manually.

WPS无法自动启动后台，需要手动启动后台。

WPS cannot start the backend automatically, you need to start the backend manually.

```bash
# start the backend manually
typst-docx
```
