use actix_web::{App, HttpServer, Responder};
use anyhow::Result;
use std::fs::File;
use std::{io::Write, path::PathBuf};

fn generate_temp_path(key: &str, ext: &str) -> Result<PathBuf> {
    let temp_dir = std::env::temp_dir();
    let mut temp_file = temp_dir.clone();
    temp_file.push(format!("{}.{}", key, ext));
    Ok(temp_file)
}

fn generate_uuid_temp_path(key: &str, ext: &str) -> Result<PathBuf> {
    let uuid = uuid::Uuid::new_v4();
    generate_temp_path(&format!("{}-{}", key, uuid), ext)
}

fn clipboard_to_temp_file(key: &str, ext: &str) -> Result<std::path::PathBuf> {
    let clipboard = arboard::Clipboard::new()?.get_text()?;
    println!("Clipboard text: {}", clipboard);
    let path = generate_uuid_temp_path(key, ext)?;
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
        anyhow::bail!(
            "Failed to convert typst to docx, Consider installing pandoc: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }
    let docx_file = docx_file.to_string_lossy().to_string();
    println!("Docx file created: {}", docx_file);
    Ok(docx_file)
}

#[actix_web::get("/")]
async fn typst_docx() -> impl Responder {
    clipboard_to_docx().unwrap_or("".to_string())
}

#[actix_web::main]
async fn main() -> Result<()> {
    let addr = "127.0.0.1:5180";
    println!("Server started at: http://{}", addr);
    println!("Use the macro at js/macro.js in your WPS Office to convert typst to docx");
    println!("You should download pandoc and add it to your PATH to use this service.");
    HttpServer::new(|| App::new().service(typst_docx))
        .bind(addr)?
        .run()
        .await?;
    Ok(())
}
