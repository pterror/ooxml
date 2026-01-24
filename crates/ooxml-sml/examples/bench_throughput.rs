//! Benchmark: measure cell throughput with generated parsers.
//!
//! Usage: cargo run -p ooxml-sml --release --example bench_throughput -- <path> [--limit N]

use ooxml_sml::{CellResolveExt, RowExt, Workbook};
use std::path::{Path, PathBuf};
use std::time::Instant;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <path-to-xlsx-dir> [--limit N]", args[0]);
        std::process::exit(1);
    }

    let path = Path::new(&args[1]);
    let limit: Option<usize> = args
        .iter()
        .position(|a| a == "--limit")
        .and_then(|i| args.get(i + 1))
        .and_then(|s| s.parse().ok());

    let files = collect_xlsx_files(path, limit);
    eprintln!("Found {} XLSX files", files.len());

    let start = Instant::now();
    let mut total_cells: u64 = 0;
    let mut total_rows: u64 = 0;
    let mut total_sheets: u64 = 0;
    let mut success_count = 0;
    let mut fail_count = 0;

    for (i, file) in files.iter().enumerate() {
        if (i + 1) % 50 == 0 {
            eprintln!("  Processing {}/{}", i + 1, files.len());
        }

        match process_file(file) {
            Ok((sheets, rows, cells)) => {
                total_sheets += sheets;
                total_rows += rows;
                total_cells += cells;
                success_count += 1;
            }
            Err(_) => {
                fail_count += 1;
            }
        }
    }

    let elapsed = start.elapsed();
    let secs = elapsed.as_secs_f64();

    println!();
    println!("=== BENCHMARK RESULTS ===");
    println!();
    println!(
        "Files:     {} success, {} failed",
        success_count, fail_count
    );
    println!("Sheets:    {}", total_sheets);
    println!("Rows:      {}", total_rows);
    println!("Cells:     {}", total_cells);
    println!();
    println!("Time:      {:.2}s", secs);
    println!("Cells/sec: {:.0}", total_cells as f64 / secs);
    println!("Rows/sec:  {:.0}", total_rows as f64 / secs);
}

fn process_file(path: &Path) -> Result<(u64, u64, u64), ooxml_sml::Error> {
    let mut workbook = Workbook::open(path)?;
    let mut sheets = 0u64;
    let mut rows = 0u64;
    let mut cells = 0u64;

    for sheet in workbook.resolved_sheets()? {
        sheets += 1;
        for row in sheet.rows() {
            rows += 1;
            for cell in row.cells_iter() {
                cells += 1;
                // Access the value to ensure it's actually parsed
                let _ = cell.value_as_string(sheet.context());
            }
        }
    }

    Ok((sheets, rows, cells))
}

fn collect_xlsx_files(path: &Path, limit: Option<usize>) -> Vec<PathBuf> {
    let mut files = Vec::new();
    collect_recursive(path, &mut files);
    if let Some(max) = limit {
        files.truncate(max);
    }
    files
}

fn collect_recursive(dir: &Path, files: &mut Vec<PathBuf>) {
    if dir.is_file() {
        if dir
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
        {
            files.push(dir.to_path_buf());
        }
        return;
    }

    if let Ok(entries) = std::fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_recursive(&path, files);
            } else if path
                .extension()
                .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
            {
                files.push(path);
            }
        }
    }
}
