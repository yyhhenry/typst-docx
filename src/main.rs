use std::{io::Write, path::PathBuf};

use anyhow::Result;
use clap::Parser;
use std::fs::File;

/// Convert typst file to docx
#[derive(Parser)]
struct Args {
    /// The input typst file (*.typ), if not provided, clipboard text will be used
    input_file: Option<String>,
}

fn generate_temp_path(key: &str, ext: &str) -> Result<PathBuf> {
    let temp_dir = std::env::temp_dir();
    let mut temp_file = temp_dir.clone();
    temp_file.push(format!("{}-{}.{}", key, uuid::Uuid::new_v4(), ext));
    Ok(temp_file)
}

fn create_temp_file(key: &str, ext: &str) -> Result<(PathBuf, File)> {
    let path = generate_temp_path(key, ext)?;
    let file = std::fs::File::create(&path)?;
    Ok((path, file))
}

fn clipboard_to_temp_file(key: &str, ext: &str) -> Result<std::path::PathBuf> {
    let clipboard = arboard::Clipboard::new()?.get_text()?;
    let (path, mut file) = create_temp_file(key, ext)?;
    file.write_all(clipboard.as_bytes())?;
    Ok(path)
}

fn main() -> Result<()> {
    let args = Args::parse();
    let path = match args.input_file {
        Some(file) => PathBuf::from(file),
        None => clipboard_to_temp_file("typst-docx", "typ")?,
    };
    if !path.is_file() || path.extension() != Some("typ".as_ref()) {
        anyhow::bail!("Invalid input file");
    }
    let docx_path = generate_temp_path("typst-docx", "docx")?;
    // pandoc $path -o $docx_path
    let output = std::process::Command::new("pandoc")
        .arg(&path)
        .arg("-o")
        .arg(&docx_path)
        .output()?;
    if !output.status.success() {
        anyhow::bail!(
            "Failed to convert typst to docx, Consider installing pandoc: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }
    // open docx with default application (Using powershell)
    let ps_command = format!(
        "Start-Process -FilePath \"{}\"",
        docx_path.to_string_lossy()
    );
    let output = std::process::Command::new("powershell")
        .arg("-Command")
        .arg(ps_command)
        .output()?;
    if !output.status.success() {
        anyhow::bail!(
            "Failed to open docx file, Consider installing powershell and Word: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }
    println!("Docx file created: {}", docx_path.to_string_lossy());
    Ok(())
}
