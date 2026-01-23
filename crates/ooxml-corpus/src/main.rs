//! Corpus analysis tool for OOXML documents.
//!
//! Analyzes collections of DOCX/XLSX/PPTX files to discover patterns,
//! edge cases, and test our parser against real-world documents.

use ooxml_corpus::{
    CorpusCoverage, CorpusDatabase, CorpusEdgeCaseStats, CorpusFeatureStats, CorpusValidationStats,
    CoverageReport, DocumentEdgeCases, DocumentFeatures, FixtureCriteria, ValidationResult,
    analyze_error, detect_edge_cases, extract_coverage_from_file, extract_features,
    extract_fixture, validate_document,
};
use rayon::prelude::*;
use serde::Serialize;
use std::collections::HashMap;
use std::fs::{self, File};
use std::io::{BufReader, Read, Seek};
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::Instant;

/// Results from analyzing a corpus of documents.
#[derive(Debug, Default, Serialize)]
struct CorpusAnalysis {
    /// Total files found.
    total_files: usize,
    /// Successfully parsed.
    successes: usize,
    /// Failed to parse.
    failures: usize,
    /// Skipped (not DOCX, or couldn't read).
    skipped: usize,
    /// Error messages grouped by category.
    errors: HashMap<String, Vec<String>>,
    /// Time taken in seconds.
    duration_secs: f64,
    /// Feature statistics (if --features enabled).
    #[serde(skip_serializing_if = "Option::is_none")]
    feature_stats: Option<CorpusFeatureStats>,
    /// Edge case statistics (if --edge-cases enabled).
    #[serde(skip_serializing_if = "Option::is_none")]
    edge_case_stats: Option<CorpusEdgeCaseStats>,
    /// Validation statistics (if --validate enabled).
    #[serde(skip_serializing_if = "Option::is_none")]
    validation_stats: Option<CorpusValidationStats>,
}

/// Result from analyzing a single document.
#[derive(Debug)]
enum DocResult {
    Success {
        #[allow(dead_code)]
        path: String,
        features: Option<Box<DocumentFeatures>>,
        edge_cases: Option<DocumentEdgeCases>,
        validation: Option<ValidationResult>,
    },
    Failure {
        path: String,
        error: ooxml_wml::Error,
    },
    Skipped {
        #[allow(dead_code)]
        path: String,
        #[allow(dead_code)]
        reason: String,
    },
}

/// Configuration for corpus analysis.
struct AnalysisConfig {
    /// Extract feature statistics.
    extract_features: bool,
    /// Detect edge cases.
    detect_edge_cases: bool,
    /// Run validation checks.
    validate: bool,
    /// Database for storing results.
    db: Option<CorpusDatabase>,
    /// Corpus name for database storage.
    corpus_name: String,
}

fn main() {
    let args: Vec<String> = std::env::args().collect();

    if args.len() < 2 {
        print_usage();
        std::process::exit(1);
    }

    // Check for subcommands
    match args[1].as_str() {
        "stats" => {
            if args.len() < 3 {
                eprintln!("Usage: ooxml-corpus stats <database.db> [corpus-name]");
                std::process::exit(1);
            }
            run_stats(&args[2], args.get(3).map(|s| s.as_str()));
            return;
        }
        "list" => {
            if args.len() < 3 {
                eprintln!("Usage: ooxml-corpus list <database.db>");
                std::process::exit(1);
            }
            run_list(&args[2]);
            return;
        }
        "failures" => {
            if args.len() < 4 {
                eprintln!("Usage: ooxml-corpus failures <database.db> <corpus-name>");
                std::process::exit(1);
            }
            run_failures(&args[2], &args[3]);
            return;
        }
        "extract" => {
            if args.len() < 4 {
                eprintln!("Usage: ooxml-corpus extract <source-path> <output-dir> [options]");
                eprintln!();
                eprintln!("Options:");
                eprintln!("  --corpus NAME    Corpus name for manifest");
                eprintln!("  --interesting    Extract only unusual/rare documents");
                eprintln!("  --minimal        Extract only small documents (<50 paragraphs)");
                eprintln!("  --max N          Maximum number of fixtures to extract");
                eprintln!(
                    "  --feature FEAT   Extract documents with feature (tables/images/lists/hyperlinks)"
                );
                std::process::exit(1);
            }
            run_extract(&args[2..]);
            return;
        }
        "coverage" => {
            if args.len() < 3 {
                eprintln!("Usage: ooxml-corpus coverage <path> [options]");
                eprintln!();
                eprintln!("Options:");
                eprintln!("  --limit N    Only process first N files");
                eprintln!("  --json       Output as JSON");
                std::process::exit(1);
            }
            let limit = args
                .iter()
                .position(|a| a == "--limit")
                .and_then(|i| args.get(i + 1))
                .and_then(|s| s.parse().ok());
            let json_output = args.iter().any(|a| a == "--json");
            run_coverage(&args[2], limit, json_output);
            return;
        }
        "sml-attrs" => {
            if args.len() < 3 {
                eprintln!("Usage: ooxml-corpus sml-attrs <path> [options]");
                eprintln!();
                eprintln!("Analyze SML (SpreadsheetML) attribute usage in XLSX files.");
                eprintln!("Useful for determining which attributes are commonly vs rarely used.");
                eprintln!();
                eprintln!("Options:");
                eprintln!("  --limit N    Only process first N files");
                eprintln!("  --json       Output as JSON");
                std::process::exit(1);
            }
            let limit = args
                .iter()
                .position(|a| a == "--limit")
                .and_then(|i| args.get(i + 1))
                .and_then(|s| s.parse().ok());
            let json_output = args.iter().any(|a| a == "--json");
            run_sml_attrs(&args[2], limit, json_output);
            return;
        }
        "--help" | "-h" | "help" => {
            print_usage();
            return;
        }
        _ => {}
    }

    let path = &args[1];
    let json_output = args.iter().any(|a| a == "--json");
    let extract_features_flag = args.iter().any(|a| a == "--features");
    let detect_edge_cases_flag = args.iter().any(|a| a == "--edge-cases");
    let validate_flag = args.iter().any(|a| a == "--validate");
    let limit = args
        .iter()
        .position(|a| a == "--limit")
        .and_then(|i| args.get(i + 1))
        .and_then(|s| s.parse().ok());

    let db_path = args
        .iter()
        .position(|a| a == "--db")
        .and_then(|i| args.get(i + 1));

    let corpus_name = args
        .iter()
        .position(|a| a == "--corpus")
        .and_then(|i| args.get(i + 1).cloned());

    let path = Path::new(path);

    if !path.exists() {
        eprintln!("Error: path does not exist: {}", path.display());
        std::process::exit(1);
    }

    // Determine corpus name from path if not provided
    let corpus_name = corpus_name.unwrap_or_else(|| {
        path.file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("unknown")
            .to_string()
    });

    // Open database if specified
    let db = if let Some(db_path) = db_path {
        match CorpusDatabase::open(db_path) {
            Ok(db) => {
                eprintln!("Storing results in: {}", db_path);
                Some(db)
            }
            Err(e) => {
                eprintln!("Error opening database: {}", e);
                std::process::exit(1);
            }
        }
    } else {
        None
    };

    // Edge cases require features to be extracted first
    let needs_features = extract_features_flag || detect_edge_cases_flag || db.is_some();

    let config = AnalysisConfig {
        extract_features: needs_features,
        detect_edge_cases: detect_edge_cases_flag,
        validate: validate_flag,
        db,
        corpus_name,
    };

    let analysis = if path.is_dir() {
        analyze_directory(path, limit, config)
    } else if path.extension().is_some_and(|e| e == "zip") {
        analyze_zip(path, limit, config)
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

fn analyze_directory(dir: &Path, limit: Option<usize>, config: AnalysisConfig) -> CorpusAnalysis {
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
    let feature_stats = Mutex::new(CorpusFeatureStats::new());
    let edge_case_stats = Mutex::new(CorpusEdgeCaseStats::new());
    let validation_stats = Mutex::new(CorpusValidationStats::new());

    let results: Vec<DocResult> = files
        .into_par_iter()
        .take(limit.unwrap_or(usize::MAX))
        .map(|path| {
            let count = processed.fetch_add(1, Ordering::Relaxed) + 1;
            if count.is_multiple_of(100) {
                eprintln!("  Processed {}/{}", count, file_count);
            }
            analyze_docx_file(
                &path,
                config.extract_features,
                config.detect_edge_cases,
                config.validate,
            )
        })
        .collect();

    let mut analysis = CorpusAnalysis {
        total_files: file_count,
        duration_secs: start.elapsed().as_secs_f64(),
        ..Default::default()
    };

    for result in results {
        match result {
            DocResult::Success {
                ref path,
                ref features,
                ref edge_cases,
                ref validation,
            } => {
                analysis.successes += 1;
                if let Some(f) = features {
                    feature_stats.lock().unwrap().add(f);
                    // Store in database if available
                    if let Some(ref db) = config.db {
                        let _ = db.insert_success(path, &config.corpus_name, f);
                    }
                }
                if let Some(ec) = edge_cases {
                    edge_case_stats.lock().unwrap().add(ec);
                }
                if let Some(v) = validation {
                    validation_stats.lock().unwrap().add(v);
                }
            }
            DocResult::Failure {
                ref path,
                ref error,
            } => {
                analysis.failures += 1;
                let analyzed = analyze_error(error);
                let category = analyzed.category.as_str().to_string();
                analysis
                    .errors
                    .entry(category)
                    .or_default()
                    .push(format!("{}: {}", path, analyzed.message));
                // Store failure in database if available
                if let Some(ref db) = config.db {
                    let _ = db.insert_failure(path, &config.corpus_name, &analyzed);
                }
            }
            DocResult::Skipped { .. } => analysis.skipped += 1,
        }
    }

    if config.extract_features {
        analysis.feature_stats = Some(feature_stats.into_inner().unwrap());
    }
    if config.detect_edge_cases {
        analysis.edge_case_stats = Some(edge_case_stats.into_inner().unwrap());
    }
    if config.validate {
        analysis.validation_stats = Some(validation_stats.into_inner().unwrap());
    }

    analysis
}

fn analyze_zip(zip_path: &Path, limit: Option<usize>, config: AnalysisConfig) -> CorpusAnalysis {
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
    let mut feature_stats = CorpusFeatureStats::new();
    let mut edge_case_stats = CorpusEdgeCaseStats::new();
    let mut validation_stats = CorpusValidationStats::new();

    // ZIP archives aren't thread-safe, so we process sequentially
    for (count, &idx) in docx_indices.iter().take(file_count).enumerate() {
        if (count + 1).is_multiple_of(100) {
            eprintln!("  Processed {}/{}", count + 1, file_count);
        }

        let result = analyze_docx_from_zip(
            &mut archive,
            idx,
            config.extract_features,
            config.detect_edge_cases,
            config.validate,
        );
        match result {
            DocResult::Success {
                ref path,
                ref features,
                ref edge_cases,
                ref validation,
            } => {
                analysis.successes += 1;
                if let Some(f) = features {
                    feature_stats.add(f);
                    // Store in database if available
                    if let Some(ref db) = config.db {
                        let _ = db.insert_success(path, &config.corpus_name, f);
                    }
                }
                if let Some(ec) = edge_cases {
                    edge_case_stats.add(ec);
                }
                if let Some(v) = validation {
                    validation_stats.add(v);
                }
            }
            DocResult::Failure {
                ref path,
                ref error,
            } => {
                analysis.failures += 1;
                let analyzed = analyze_error(error);
                let category = analyzed.category.as_str().to_string();
                analysis
                    .errors
                    .entry(category)
                    .or_default()
                    .push(format!("{}: {}", path, analyzed.message));
                // Store failure in database if available
                if let Some(ref db) = config.db {
                    let _ = db.insert_failure(path, &config.corpus_name, &analyzed);
                }
            }
            DocResult::Skipped { .. } => analysis.skipped += 1,
        }
    }

    analysis.duration_secs = start.elapsed().as_secs_f64();

    if config.extract_features {
        analysis.feature_stats = Some(feature_stats);
    }
    if config.detect_edge_cases {
        analysis.edge_case_stats = Some(edge_case_stats);
    }
    if config.validate {
        analysis.validation_stats = Some(validation_stats);
    }

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

fn analyze_docx_file(
    path: &Path,
    extract_features_flag: bool,
    detect_edge_cases_flag: bool,
    validate_flag: bool,
) -> DocResult {
    let path_str = path.display().to_string();

    match ooxml_wml::Document::open(path) {
        Ok(doc) => {
            // Try to access the body to ensure full parsing
            let _ = doc.body().text();

            let features = if extract_features_flag {
                Some(Box::new(extract_features(&doc)))
            } else {
                None
            };

            let edge_cases = if detect_edge_cases_flag {
                features.as_ref().map(|f| detect_edge_cases(f))
            } else {
                None
            };

            let validation = if validate_flag {
                Some(validate_document(&doc))
            } else {
                None
            };

            DocResult::Success {
                path: path_str,
                features,
                edge_cases,
                validation,
            }
        }
        Err(e) => DocResult::Failure {
            path: path_str,
            error: e,
        },
    }
}

fn analyze_docx_from_zip<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    idx: usize,
    extract_features_flag: bool,
    detect_edge_cases_flag: bool,
    validate_flag: bool,
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
            error: ooxml_wml::Error::Io(e),
        };
    }

    // Parse with ooxml-wml
    let cursor = std::io::Cursor::new(buffer);
    match ooxml_wml::Document::from_reader(cursor) {
        Ok(doc) => {
            // Try to access the body to ensure full parsing
            let _ = doc.body().text();

            let features = if extract_features_flag {
                Some(Box::new(extract_features(&doc)))
            } else {
                None
            };

            let edge_cases = if detect_edge_cases_flag {
                features.as_ref().map(|f| detect_edge_cases(f))
            } else {
                None
            };

            let validation = if validate_flag {
                Some(validate_document(&doc))
            } else {
                None
            };

            DocResult::Success {
                path: name,
                features,
                edge_cases,
                validation,
            }
        }
        Err(e) => DocResult::Failure {
            path: name,
            error: e,
        },
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

        // Sort by count (descending)
        let mut error_vec: Vec<_> = analysis.errors.iter().collect();
        error_vec.sort_by(|a, b| b.1.len().cmp(&a.1.len()));

        for (category, errors) in error_vec {
            println!();
            println!(
                "{}: {} occurrences",
                format_category(category),
                errors.len()
            );
            // Show first 3 examples
            for error in errors.iter().take(3) {
                println!("  - {}", truncate(error, 100));
            }
            if errors.len() > 3 {
                println!("  ... and {} more", errors.len() - 3);
            }
        }
    }

    if let Some(stats) = &analysis.feature_stats {
        println!();
        println!("=== Feature Statistics ===");
        println!();
        println!(
            "Documents with tables:     {} ({:.1}%)",
            stats.with_tables,
            stats.percentage(stats.with_tables)
        );
        println!(
            "Documents with images:     {} ({:.1}%)",
            stats.with_images,
            stats.percentage(stats.with_images)
        );
        println!(
            "Documents with hyperlinks: {} ({:.1}%)",
            stats.with_hyperlinks,
            stats.percentage(stats.with_hyperlinks)
        );
        println!(
            "Documents with lists:      {} ({:.1}%)",
            stats.with_lists,
            stats.percentage(stats.with_lists)
        );
        println!(
            "Documents with bold:       {} ({:.1}%)",
            stats.with_bold,
            stats.percentage(stats.with_bold)
        );
        println!(
            "Documents with italic:     {} ({:.1}%)",
            stats.with_italic,
            stats.percentage(stats.with_italic)
        );
        println!(
            "Documents with color:      {} ({:.1}%)",
            stats.with_color,
            stats.percentage(stats.with_color)
        );
        println!(
            "Documents with alignment:  {} ({:.1}%)",
            stats.with_alignment,
            stats.percentage(stats.with_alignment)
        );
        println!();
        println!("Total paragraphs: {}", stats.total_paragraphs);
        println!("Total tables:     {}", stats.total_tables);
        println!("Total images:     {}", stats.total_images);
        println!("Total hyperlinks: {}", stats.total_hyperlinks);
        println!();
        println!("Max paragraphs in one doc: {}", stats.max_paragraphs);
        println!("Max tables in one doc:     {}", stats.max_tables);
        println!("Max table nesting depth:   {}", stats.max_table_nesting);

        if !stats.font_usage.is_empty() {
            println!();
            println!("Top fonts used:");
            let mut fonts: Vec<_> = stats.font_usage.iter().collect();
            fonts.sort_by(|a, b| b.1.cmp(a.1));
            for (font, count) in fonts.iter().take(10) {
                println!("  {}: {} docs", font, count);
            }
        }

        if !stats.style_usage.is_empty() {
            println!();
            println!("Top styles used:");
            let mut styles: Vec<_> = stats.style_usage.iter().collect();
            styles.sort_by(|a, b| b.1.cmp(a.1));
            for (style, count) in styles.iter().take(10) {
                println!("  {}: {} docs", style, count);
            }
        }
    }

    if let Some(stats) = &analysis.edge_case_stats {
        println!();
        println!("=== Edge Case Statistics ===");
        println!();
        println!(
            "Documents with edge cases: {} ({:.1}%)",
            stats.documents_with_edge_cases,
            stats.edge_case_percentage()
        );
        println!("Total documents checked:   {}", stats.total_documents);

        if !stats.by_type.is_empty() {
            println!();
            println!("Edge cases by type:");
            let mut by_type: Vec<_> = stats.by_type.iter().collect();
            by_type.sort_by(|a, b| b.1.cmp(a.1));
            for (edge_type, count) in by_type.iter().take(10) {
                println!("  {}: {} docs", edge_type, count);
            }
        }

        if !stats.by_severity.is_empty() {
            println!();
            println!("By severity:");
            for severity in &["rare", "unusual", "noteworthy", "info"] {
                if let Some(count) = stats.by_severity.get(*severity) {
                    println!("  {}: {}", severity, count);
                }
            }
        }

        if !stats.max_values.is_empty() {
            println!();
            println!("Maximum values observed:");
            for (key, value) in &stats.max_values {
                println!("  {}: {}", key, value);
            }
        }
    }

    if let Some(stats) = &analysis.validation_stats {
        println!();
        println!("=== Validation Statistics ===");
        println!();
        println!(
            "Valid documents:           {} ({:.1}%)",
            stats.valid_documents,
            stats.valid_percentage()
        );
        println!(
            "Documents with warnings:   {}",
            stats.documents_with_warnings
        );
        println!("Documents with errors:     {}", stats.documents_with_errors);

        if !stats.warnings_by_code.is_empty() {
            println!();
            println!("Warnings by type:");
            let mut warnings: Vec<_> = stats.warnings_by_code.iter().collect();
            warnings.sort_by(|a, b| b.1.cmp(a.1));
            for (code, count) in warnings {
                println!("  {}: {}", code, count);
            }
        }

        if !stats.errors_by_code.is_empty() {
            println!();
            println!("Errors by type:");
            let mut errors: Vec<_> = stats.errors_by_code.iter().collect();
            errors.sort_by(|a, b| b.1.cmp(a.1));
            for (code, count) in errors {
                println!("  {}: {}", code, count);
            }
        }
    }

    println!();
    if analysis.total_files > 0 {
        let rate = analysis.total_files as f64 / analysis.duration_secs;
        println!("Processing rate: {:.1} files/sec", rate);
    }
}

fn format_category(category: &str) -> &str {
    match category {
        "zip_corruption" => "ZIP Corruption",
        "zip_unsupported_compression" => "Unsupported ZIP Compression",
        "missing_required_part" => "Missing Required Part",
        "invalid_content_type" => "Invalid Content Type",
        "xml_malformed" => "Malformed XML",
        "xml_encoding" => "XML Encoding Error",
        "xml_namespace" => "XML Namespace Error",
        "unexpected_element" => "Unexpected Element",
        "missing_attribute" => "Missing Attribute",
        "invalid_attribute" => "Invalid Attribute",
        "broken_relationship" => "Broken Relationship",
        "missing_target" => "Missing Target",
        "unsupported_feature" => "Unsupported Feature",
        "not_implemented" => "Not Implemented",
        "io_error" => "I/O Error",
        "unknown" => "Unknown",
        other => other,
    }
}

fn truncate(s: &str, max_len: usize) -> String {
    if s.len() <= max_len {
        s.to_string()
    } else {
        format!("{}...", &s[..max_len - 3])
    }
}

fn print_usage() {
    eprintln!("Usage: ooxml-corpus <command|path> [options]");
    eprintln!();
    eprintln!("Commands:");
    eprintln!("  stats <db> [corpus]       Show statistics for a corpus in the database");
    eprintln!("  list <db>                 List all corpora in the database");
    eprintln!("  failures <db> <corpus>    List all failures for a corpus");
    eprintln!("  extract <src> <out> [opts] Extract test fixtures from documents");
    eprintln!("  sml-attrs <path> [opts]   Analyze SML attribute usage in XLSX files");
    eprintln!("  coverage <path> [opts]    Analyze WML element coverage in DOCX files");
    eprintln!();
    eprintln!("Analyze Mode:");
    eprintln!("  ooxml-corpus <path> [options]");
    eprintln!();
    eprintln!("Analyze Options:");
    eprintln!("  --json         Output results as JSON");
    eprintln!("  --limit N      Only process first N files");
    eprintln!("  --features     Extract feature statistics");
    eprintln!("  --edge-cases   Detect edge cases and unusual patterns");
    eprintln!("  --validate     Run validation checks");
    eprintln!("  --db FILE      Store results in SQLite database");
    eprintln!("  --corpus NAME  Corpus name for database (default: directory name)");
    eprintln!();
    eprintln!("Extract Options:");
    eprintln!("  --interesting  Extract only unusual/rare documents");
    eprintln!("  --minimal      Extract only small documents (<50 paragraphs)");
    eprintln!("  --max N        Maximum fixtures to extract (default: 10)");
    eprintln!("  --feature FEAT Extract docs with feature (tables/images/lists/hyperlinks)");
    eprintln!("  --corpus NAME  Corpus name for manifest");
    eprintln!();
    eprintln!("Examples:");
    eprintln!("  ooxml-corpus ./corpora/napierone/DOCX/");
    eprintln!("  ooxml-corpus ./corpora/napierone/DOCX/ --features --json");
    eprintln!("  ooxml-corpus ./corpus.zip --limit 100");
    eprintln!("  ooxml-corpus ./DOCX/ --db corpus.db --corpus napierone-docx");
    eprintln!("  ooxml-corpus ./DOCX/ --edge-cases --validate");
    eprintln!("  ooxml-corpus extract ./DOCX/ ./fixtures --interesting --max 5");
    eprintln!("  ooxml-corpus extract ./DOCX/ ./fixtures --feature tables --minimal");
    eprintln!("  ooxml-corpus stats corpus.db napierone-docx");
}

fn run_stats(db_path: &str, corpus_name: Option<&str>) {
    let db = match CorpusDatabase::open(db_path) {
        Ok(db) => db,
        Err(e) => {
            eprintln!("Error opening database: {}", e);
            std::process::exit(1);
        }
    };

    // If no corpus specified, show all corpora
    if corpus_name.is_none() {
        match db.list_corpora() {
            Ok(corpora) => {
                if corpora.is_empty() {
                    println!("No corpora in database.");
                    return;
                }
                println!("=== All Corpora ===\n");
                for corpus in corpora {
                    print_corpus_stats(&db, &corpus);
                    println!();
                }
            }
            Err(e) => {
                eprintln!("Error listing corpora: {}", e);
                std::process::exit(1);
            }
        }
        return;
    }

    print_corpus_stats(&db, corpus_name.unwrap());
}

fn print_corpus_stats(db: &CorpusDatabase, corpus: &str) {
    match db.get_corpus_stats(corpus) {
        Ok(stats) => {
            println!("Corpus: {}", corpus);
            println!("  Total documents: {}", stats.total_documents);
            println!(
                "  Successes:       {} ({:.1}%)",
                stats.successes,
                stats.success_rate()
            );
            println!(
                "  Failures:        {} ({:.1}%)",
                stats.failures,
                100.0 - stats.success_rate()
            );

            if !stats.error_categories.is_empty() {
                println!("  Error breakdown:");
                for (category, count) in &stats.error_categories {
                    println!("    {}: {}", format_category(category), count);
                }
            }
        }
        Err(e) => {
            eprintln!("Error getting stats for {}: {}", corpus, e);
        }
    }
}

fn run_list(db_path: &str) {
    let db = match CorpusDatabase::open(db_path) {
        Ok(db) => db,
        Err(e) => {
            eprintln!("Error opening database: {}", e);
            std::process::exit(1);
        }
    };

    match db.list_corpora() {
        Ok(corpora) => {
            if corpora.is_empty() {
                println!("No corpora in database.");
            } else {
                println!("Corpora in {}:", db_path);
                for corpus in corpora {
                    if let Ok(stats) = db.get_corpus_stats(&corpus) {
                        println!(
                            "  {} ({} docs, {:.1}% success)",
                            corpus,
                            stats.total_documents,
                            stats.success_rate()
                        );
                    } else {
                        println!("  {}", corpus);
                    }
                }
            }
        }
        Err(e) => {
            eprintln!("Error listing corpora: {}", e);
            std::process::exit(1);
        }
    }
}

fn run_failures(db_path: &str, corpus_name: &str) {
    let db = match CorpusDatabase::open(db_path) {
        Ok(db) => db,
        Err(e) => {
            eprintln!("Error opening database: {}", e);
            std::process::exit(1);
        }
    };

    match db.get_failures(corpus_name) {
        Ok(failures) => {
            if failures.is_empty() {
                println!("No failures in corpus '{}'.", corpus_name);
            } else {
                println!(
                    "Failures in '{}' ({} total):\n",
                    corpus_name,
                    failures.len()
                );
                for doc in failures {
                    println!("Path: {}", doc.path);
                    if let Some(err) = doc.parse_error {
                        println!("  Category: {}", format_category(&err.category));
                        println!("  Message:  {}", truncate(&err.message, 80));
                    }
                    println!();
                }
            }
        }
        Err(e) => {
            eprintln!("Error getting failures: {}", e);
            std::process::exit(1);
        }
    }
}

fn run_extract(args: &[String]) {
    let source_path = Path::new(&args[0]);
    let output_dir = Path::new(&args[1]);

    // Parse options
    let corpus_name = args
        .iter()
        .position(|a| a == "--corpus")
        .and_then(|i| args.get(i + 1).cloned())
        .unwrap_or_else(|| "extracted".to_string());

    let interesting = args.iter().any(|a| a == "--interesting");
    let minimal = args.iter().any(|a| a == "--minimal");
    let max_fixtures = args
        .iter()
        .position(|a| a == "--max")
        .and_then(|i| args.get(i + 1))
        .and_then(|s| s.parse().ok())
        .unwrap_or(10);

    let feature = args
        .iter()
        .position(|a| a == "--feature")
        .and_then(|i| args.get(i + 1).cloned());

    // Build criteria
    let mut criteria = if interesting {
        FixtureCriteria::interesting()
    } else if minimal {
        FixtureCriteria::minimal()
    } else if let Some(feat) = &feature {
        FixtureCriteria::with_feature(feat)
    } else {
        FixtureCriteria::default()
    };

    // Apply feature filter if specified alongside other criteria
    if let Some(feat) = &feature {
        match feat.as_str() {
            "tables" => criteria.has_tables = true,
            "images" => criteria.has_images = true,
            "hyperlinks" => criteria.has_hyperlinks = true,
            "lists" => criteria.has_lists = true,
            _ => {}
        }
    }

    if !source_path.exists() {
        eprintln!(
            "Error: source path does not exist: {}",
            source_path.display()
        );
        std::process::exit(1);
    }

    // Create output directory
    if let Err(e) = fs::create_dir_all(output_dir) {
        eprintln!("Error creating output directory: {}", e);
        std::process::exit(1);
    }

    eprintln!("Scanning for fixtures in: {}", source_path.display());
    eprintln!("Output directory: {}", output_dir.display());
    eprintln!("Max fixtures: {}", max_fixtures);

    // Collect and analyze documents
    let files = if source_path.is_dir() {
        collect_docx_files(source_path)
    } else {
        vec![source_path.to_path_buf()]
    };

    eprintln!("Found {} DOCX files, analyzing...", files.len());

    let mut extracted = 0;
    let mut analyzed = 0;

    for file_path in files {
        if extracted >= max_fixtures {
            break;
        }

        analyzed += 1;
        if analyzed % 100 == 0 {
            eprintln!("  Analyzed {}, extracted {}", analyzed, extracted);
        }

        // Parse and analyze the document
        let doc = match ooxml_wml::Document::open(&file_path) {
            Ok(d) => d,
            Err(_) => continue,
        };

        let features = extract_features(&doc);
        let edge_cases = detect_edge_cases(&features);

        // Check if it matches criteria
        if !criteria.matches(&features, &edge_cases) {
            continue;
        }

        // Determine why it's interesting
        let reason = ooxml_corpus::determine_interest_reason(&features, &edge_cases)
            .unwrap_or_else(|| "Matches extraction criteria".to_string());

        // Create fixture directory
        let fixture_name = format!("fixture_{:04}", extracted);
        let fixture_dir = output_dir.join(&fixture_name);

        // Extract the fixture
        match extract_fixture(
            &file_path,
            &fixture_dir,
            &corpus_name,
            &features,
            &edge_cases,
            &reason,
        ) {
            Ok(result) => {
                println!("Extracted: {} -> {}", file_path.display(), fixture_name);
                println!("  Reason: {}", result.manifest.reason);
                println!("  Tags: {}", result.manifest.tags.join(", "));
                extracted += 1;
            }
            Err(e) => {
                eprintln!("Warning: Failed to extract {}: {}", file_path.display(), e);
            }
        }
    }

    println!();
    println!("Extraction complete!");
    println!("  Analyzed: {} documents", analyzed);
    println!("  Extracted: {} fixtures", extracted);
    println!("  Output: {}", output_dir.display());
}

fn run_coverage(path_str: &str, limit: Option<usize>, json_output: bool) {
    let path = Path::new(path_str);

    if !path.exists() {
        eprintln!("Error: path does not exist: {}", path.display());
        std::process::exit(1);
    }

    eprintln!("Scanning for DOCX files in: {}", path.display());

    let files = if path.is_dir() {
        collect_docx_files(path)
    } else {
        vec![path.to_path_buf()]
    };

    let file_count = limit.map_or(files.len(), |l| l.min(files.len()));
    eprintln!(
        "Found {} DOCX files, analyzing {}...",
        files.len(),
        file_count
    );

    let mut corpus_coverage = CorpusCoverage::new();
    let mut processed = 0;
    let mut errors = 0;

    for file_path in files.iter().take(file_count) {
        processed += 1;
        if processed % 100 == 0 {
            eprintln!("  Processed {}/{}", processed, file_count);
        }

        match extract_coverage_from_file(file_path) {
            Ok(doc_coverage) => {
                corpus_coverage.add(&doc_coverage);
            }
            Err(_) => {
                errors += 1;
            }
        }
    }

    let report = CoverageReport::from_corpus(&corpus_coverage);

    if json_output {
        println!("{}", serde_json::to_string_pretty(&report).unwrap());
    } else {
        print_coverage_report(&report, errors);
    }
}

fn run_sml_attrs(path_str: &str, limit: Option<usize>, json_output: bool) {
    let path = Path::new(path_str);

    if !path.exists() {
        eprintln!("Error: path does not exist: {}", path.display());
        std::process::exit(1);
    }

    eprintln!("Analyzing SML attribute usage in: {}", path.display());

    let stats = if path.is_dir() {
        ooxml_corpus::analyze_xlsx_directory(path, limit)
    } else {
        match ooxml_corpus::analyze_xlsx_file(path) {
            Ok(s) => s,
            Err(e) => {
                eprintln!("Error analyzing file: {}", e);
                std::process::exit(1);
            }
        }
    };

    if json_output {
        println!("{}", serde_json::to_string_pretty(&stats).unwrap());
    } else {
        stats.print_report();
    }
}

fn print_coverage_report(report: &CoverageReport, errors: usize) {
    println!();
    println!("=== OOXML Element Coverage Report ===");
    println!();
    println!("Documents analyzed: {}", report.total_documents);
    if errors > 0 {
        println!("Parse errors:       {}", errors);
    }
    println!();
    println!(
        "WML Coverage: {:.1}% ({}/{} elements)",
        report.wml_coverage_percent, report.wml_elements_seen, report.wml_elements_total
    );

    println!();
    println!("=== Top Elements by Frequency ===");
    println!();
    for (name, count, pct) in report.top_elements.iter().take(20) {
        println!("  {:30} {:5} docs ({:5.1}%)", name, count, pct);
    }

    if !report.missing_elements.is_empty() {
        println!();
        println!(
            "=== Missing WML Elements ({}) ===",
            report.missing_elements.len()
        );
        println!();
        for chunk in report.missing_elements.chunks(5) {
            println!("  {}", chunk.join(", "));
        }
    }

    if !report.unknown_elements.is_empty() {
        println!();
        println!("=== Unknown Elements (Top 20) ===");
        println!();
        for (name, count) in &report.unknown_elements {
            println!("  {:40} {} docs", name, count);
        }
    }
}
