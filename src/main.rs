use actix_web::{App, HttpServer, Responder};
use anyhow::Result;
use fs2::FileExt;
use std::fs::File;
use std::{io::Write, path::PathBuf};

fn generate_temp_path(key: &str, ext: &str) -> Result<PathBuf> {
    let temp_dir = std::env::temp_dir();
    let mut temp_file = temp_dir.clone();
    let uuid = uuid::Uuid::new_v4();
    temp_file.push(format!("{}-{}.{}", key, uuid, ext));
    Ok(temp_file)
}

fn clipboard_to_temp_file(key: &str, ext: &str) -> Result<std::path::PathBuf> {
    let clipboard = arboard::Clipboard::new()?.get_text()?;
    println!("Clipboard text: {}", clipboard);
    let path = generate_temp_path(key, ext)?;
    let mut file = File::create(&path)?;
    file.write_all(clipboard.as_bytes())?;
    Ok(path)
}

fn clipboard_to_docx() -> Result<String> {
    let input_file = clipboard_to_temp_file("typst-docx", "typ")?;
    let docx_file = generate_temp_path("typst-docx", "docx")?;
    let output = std::process::Command::new("pandoc")
        .arg(&input_file)
        .arg("-o")
        .arg(&docx_file)
        .output()?;
    if !output.status.success() {
        let err = String::from_utf8_lossy(&output.stderr).replace(
            &format!("\"{}\"", input_file.to_string_lossy()),
            "Typst code",
        );
        anyhow::bail!("{}", err);
    }
    let docx_file = docx_file.to_string_lossy().to_string();
    println!("Docx file created: {}", docx_file);
    Ok(docx_file)
}

#[actix_web::get("/")]
async fn typst_docx() -> impl Responder {
    match clipboard_to_docx() {
        Ok(docx_file) => docx_file,
        Err(e) => {
            let err = format!("Error: {}", e.to_string());
            println!("{}", err);
            err
        }
    }
}

/// Err if another instance is running
fn unique_lock() -> Result<File> {
    let lock_file = std::env::temp_dir().join("typst-docx.lock");
    let file = File::options().write(true).create(true).open(&lock_file)?;
    match file.try_lock_exclusive() {
        Ok(_) => Ok(file),
        Err(_) => anyhow::bail!("Another instance is running."),
    }
}

static MACRO_VBA: &str = include_str!("../scripts/macro.vba");

#[actix_web::main]
async fn main() -> Result<()> {
    println!("Typst Docx");
    let _lock = unique_lock()?;
    let addr = "127.0.0.1:5180";
    let server = HttpServer::new(|| App::new().service(typst_docx))
        .bind(addr)?
        .run();

    println!("Server started at: http://{}", addr);
    println!("You should download pandoc and add it to your PATH to use this service.");
    println!("Use this macro in your Word (Normal.dotm) to paste typst code as docx.");
    println!("-----\n\n{}", MACRO_VBA);

    server.await?;
    Ok(())
}
