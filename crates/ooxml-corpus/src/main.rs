//! Corpus analysis tool for OOXML documents.
//!
//! Analyzes collections of DOCX/XLSX/PPTX files to discover patterns,
//! edge cases, and test our parser against real-world documents.

use rayon::prelude::*;
use serde::Serialize;
use std::collections::HashMap;
use std::fs::{self, File};
use std::io::{BufReader, Read, Seek};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::Instant;

/// Results from analyzing a corpus of documents.
#[derive(Debug, Default, Serialize)]
struct CorpusAnalysis {
    /// Total files found
    total_files: usize,
    /// Successfully parsed
    successes: usize,
    /// Failed to parse
    failures: usize,
    /// Skipped (not DOCX, or couldn't read)
    skipped: usize,
    /// Error messages grouped by type
    errors: HashMap<String, Vec<String>>,
    /// Time taken in seconds
    duration_secs: f64,
}

/// Result from analyzing a single document.
#[derive(Debug)]
#[allow(dead_code)] // Fields kept for debugging/verbose mode
enum DocResult {
    Success { path: String },
    Failure { path: String, error: String },
    Skipped { path: String, reason: String },
}

fn main() {
    let args: Vec<String> = std::env::args().collect();

    if args.len() < 2 {
        eprintln!("Usage: ooxml-corpus <path> [--json] [--limit N]");
        eprintln!();
        eprintln!("  <path>     Directory of DOCX files, or a ZIP archive");
        eprintln!("  --json     Output results as JSON");
        eprintln!("  --limit N  Only process first N files");
        eprintln!();
        eprintln!("Examples:");
        eprintln!("  ooxml-corpus ./corpora/napierone/DOCX/");
        eprintln!("  ooxml-corpus ./corpora/napierone/DOCX-total.zip --limit 100");
        std::process::exit(1);
    }

    let path = &args[1];
    let json_output = args.iter().any(|a| a == "--json");
    let limit = args
        .iter()
        .position(|a| a == "--limit")
        .and_then(|i| args.get(i + 1))
        .and_then(|s| s.parse().ok());

    let path = Path::new(path);

    if !path.exists() {
        eprintln!("Error: path does not exist: {}", path.display());
        std::process::exit(1);
    }

    let analysis = if path.is_dir() {
        analyze_directory(path, limit)
    } else if path.extension().is_some_and(|e| e == "zip") {
        analyze_zip(path, limit)
    } else {
        eprintln!("Error: path must be a directory or ZIP file");
        std::process::exit(1);
    };

    if json_output {
        println!("{}", serde_json::to_string_pretty(&analysis).unwrap());
    } else {
        print_summary(&analysis);
    }
}

fn analyze_directory(dir: &Path, limit: Option<usize>) -> CorpusAnalysis {
    eprintln!("Scanning directory: {}", dir.display());

    let files: Vec<PathBuf> = collect_docx_files(dir);
    let file_count = limit.map_or(files.len(), |l| l.min(files.len()));

    eprintln!(
        "Found {} DOCX files, analyzing {}...",
        files.len(),
        file_count
    );

    let start = Instant::now();
    let processed = AtomicUsize::new(0);

    let results: Vec<DocResult> = files
        .into_par_iter()
        .take(limit.unwrap_or(usize::MAX))
        .map(|path| {
            let count = processed.fetch_add(1, Ordering::Relaxed) + 1;
            if count.is_multiple_of(100) {
                eprintln!("  Processed {}/{}", count, file_count);
            }
            analyze_docx_file(&path)
        })
        .collect();

    let mut analysis = CorpusAnalysis {
        total_files: file_count,
        duration_secs: start.elapsed().as_secs_f64(),
        ..Default::default()
    };

    for result in results {
        match result {
            DocResult::Success { .. } => analysis.successes += 1,
            DocResult::Failure { path, error } => {
                analysis.failures += 1;
                let error_type = categorize_error(&error);
                analysis
                    .errors
                    .entry(error_type)
                    .or_default()
                    .push(format!("{}: {}", path, error));
            }
            DocResult::Skipped { .. } => analysis.skipped += 1,
        }
    }

    analysis
}

fn analyze_zip(zip_path: &Path, limit: Option<usize>) -> CorpusAnalysis {
    eprintln!("Opening ZIP archive: {}", zip_path.display());

    let file = match File::open(zip_path) {
        Ok(f) => f,
        Err(e) => {
            eprintln!("Error opening ZIP: {}", e);
            return CorpusAnalysis::default();
        }
    };

    let mut archive = match zip::ZipArchive::new(BufReader::new(file)) {
        Ok(a) => a,
        Err(e) => {
            eprintln!("Error reading ZIP: {}", e);
            return CorpusAnalysis::default();
        }
    };

    // Collect DOCX file indices
    let docx_indices: Vec<usize> = (0..archive.len())
        .filter(|&i| {
            archive
                .by_index(i)
                .ok()
                .is_some_and(|f| f.name().ends_with(".docx") || f.name().ends_with(".DOCX"))
        })
        .collect();

    let file_count = limit.map_or(docx_indices.len(), |l| l.min(docx_indices.len()));
    eprintln!(
        "Found {} DOCX files in archive, analyzing {}...",
        docx_indices.len(),
        file_count
    );

    let start = Instant::now();
    let mut analysis = CorpusAnalysis {
        total_files: file_count,
        ..Default::default()
    };

    // ZIP archives aren't thread-safe, so we process sequentially
    for (count, &idx) in docx_indices.iter().take(file_count).enumerate() {
        if (count + 1).is_multiple_of(100) {
            eprintln!("  Processed {}/{}", count + 1, file_count);
        }

        let result = analyze_docx_from_zip(&mut archive, idx);
        match result {
            DocResult::Success { .. } => analysis.successes += 1,
            DocResult::Failure { path, error } => {
                analysis.failures += 1;
                let error_type = categorize_error(&error);
                analysis
                    .errors
                    .entry(error_type)
                    .or_default()
                    .push(format!("{}: {}", path, error));
            }
            DocResult::Skipped { .. } => analysis.skipped += 1,
        }
    }

    analysis.duration_secs = start.elapsed().as_secs_f64();
    analysis
}

fn collect_docx_files(dir: &Path) -> Vec<PathBuf> {
    let mut files = Vec::new();
    collect_docx_files_recursive(dir, &mut files);
    files
}

fn collect_docx_files_recursive(dir: &Path, files: &mut Vec<PathBuf>) {
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_docx_files_recursive(&path, files);
            } else if path.extension().is_some_and(|e| e == "docx" || e == "DOCX") {
                files.push(path);
            }
        }
    }
}

fn analyze_docx_file(path: &Path) -> DocResult {
    let path_str = path.display().to_string();

    match ooxml_wml::Document::open(path) {
        Ok(doc) => {
            // Try to access the body to ensure full parsing
            let _ = doc.body().text();
            DocResult::Success { path: path_str }
        }
        Err(e) => DocResult::Failure {
            path: path_str,
            error: e.to_string(),
        },
    }
}

fn analyze_docx_from_zip<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    idx: usize,
) -> DocResult {
    let mut file = match archive.by_index(idx) {
        Ok(f) => f,
        Err(e) => {
            return DocResult::Skipped {
                path: format!("index {}", idx),
                reason: e.to_string(),
            };
        }
    };

    let name = file.name().to_string();

    // Read the entire DOCX into memory
    let mut buffer = Vec::new();
    if let Err(e) = file.read_to_end(&mut buffer) {
        return DocResult::Failure {
            path: name,
            error: format!("Failed to read from ZIP: {}", e),
        };
    }

    // Parse with ooxml-wml
    let cursor = std::io::Cursor::new(buffer);
    match ooxml_wml::Document::from_reader(cursor) {
        Ok(doc) => {
            // Try to access the body to ensure full parsing
            let _ = doc.body().text();
            DocResult::Success { path: name }
        }
        Err(e) => DocResult::Failure {
            path: name,
            error: e.to_string(),
        },
    }
}

fn categorize_error(error: &str) -> String {
    if error.contains("zip") || error.contains("Zip") || error.contains("ZIP") {
        "ZIP/Archive error".to_string()
    } else if error.contains("xml") || error.contains("XML") || error.contains("Xml") {
        "XML parsing error".to_string()
    } else if error.contains("missing") || error.contains("Missing") {
        "Missing part error".to_string()
    } else if error.contains("invalid") || error.contains("Invalid") {
        "Invalid format error".to_string()
    } else if error.contains("unsupported") || error.contains("Unsupported") {
        "Unsupported feature".to_string()
    } else {
        "Other error".to_string()
    }
}

fn print_summary(analysis: &CorpusAnalysis) {
    println!();
    println!("=== Corpus Analysis Results ===");
    println!();
    println!("Total files:  {}", analysis.total_files);
    println!(
        "Successes:    {} ({:.1}%)",
        analysis.successes,
        if analysis.total_files > 0 {
            analysis.successes as f64 / analysis.total_files as f64 * 100.0
        } else {
            0.0
        }
    );
    println!(
        "Failures:     {} ({:.1}%)",
        analysis.failures,
        if analysis.total_files > 0 {
            analysis.failures as f64 / analysis.total_files as f64 * 100.0
        } else {
            0.0
        }
    );
    println!("Skipped:      {}", analysis.skipped);
    println!("Duration:     {:.2}s", analysis.duration_secs);

    if !analysis.errors.is_empty() {
        println!();
        println!("=== Errors by Category ===");
        for (category, errors) in &analysis.errors {
            println!();
            println!("{}: {} occurrences", category, errors.len());
            // Show first 3 examples
            for error in errors.iter().take(3) {
                println!("  - {}", truncate(error, 100));
            }
            if errors.len() > 3 {
                println!("  ... and {} more", errors.len() - 3);
            }
        }
    }

    println!();
    if analysis.total_files > 0 {
        let rate = analysis.total_files as f64 / analysis.duration_secs;
        println!("Processing rate: {:.1} files/sec", rate);
    }
}

fn truncate(s: &str, max_len: usize) -> String {
    if s.len() <= max_len {
        s.to_string()
    } else {
        format!("{}...", &s[..max_len - 3])
    }
}
