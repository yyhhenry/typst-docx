# Paste Typst docx

这是一个用于在WPS/Word中直接粘贴Typst内容的宏。

This is a macro for pasting Typst content directly into WPS/Word.

## 使用方法

保证电脑中已经安装了pandoc（添加到PATH），Rust环境，以及WPS/Word。

Make sure you have pandoc (added to PATH), Rust environment, and WPS/Word installed on your computer.

```bash
# install pandoc and rust
scoop install pandoc rustup

# install typst-docx
cargo install --git https://github.com/yyhhenry/typst-docx
```

在Normal.dotm中添加[VBA宏](scripts/macro.vba)，然后根据需求添加到快捷访问工具栏。

Add [VBA macro](scripts/macro.vba) to Normal.dotm, then add it to the Quick Access Toolbar as needed.

运行此宏会将剪贴板中的Typst内容以docx格式粘贴到当前光标处，使用默认样式，末尾有换行符。

Running this macro will paste the Typst content in the clipboard into the current cursor position in docx format, using the default style, with a newline at the end.

你可以手动启动后端，或者宏会自动启动后端。

You can start the backend manually, or the macro will start the backend automatically.
