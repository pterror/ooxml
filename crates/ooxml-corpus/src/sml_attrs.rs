//! SML (SpreadsheetML) attribute usage analysis.
//!
//! Analyzes real XLSX files to determine which attributes are commonly used
//! vs rarely used. This data can inform codegen decisions about which fields
//! to parse eagerly vs lazily.

use quick_xml::Reader;
use quick_xml::events::Event;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read};
use std::path::Path;
use zip::ZipArchive;

/// Statistics about SML attribute usage across a corpus.
#[derive(Debug, Default, Serialize, Deserialize)]
pub struct SmlAttrStats {
    /// Files analyzed.
    pub files_analyzed: u64,
    /// Worksheets analyzed.
    pub total_worksheets: u64,
    /// Total rows seen.
    pub total_rows: u64,
    /// Total cells seen.
    pub total_cells: u64,
    /// Total formulas seen.
    pub total_formulas: u64,

    /// Row attribute counts: attr_name -> count.
    pub row_attrs: HashMap<String, u64>,
    /// Cell attribute counts: attr_name -> count.
    pub cell_attrs: HashMap<String, u64>,
    /// Formula attribute counts: attr_name -> count.
    pub formula_attrs: HashMap<String, u64>,
    /// Worksheet child element counts: element_name -> count.
    pub worksheet_children: HashMap<String, u64>,
}

impl SmlAttrStats {
    pub fn new() -> Self {
        Self::default()
    }

    /// Merge another stats instance into this one.
    pub fn merge(&mut self, other: &SmlAttrStats) {
        self.files_analyzed += other.files_analyzed;
        self.total_worksheets += other.total_worksheets;
        self.total_rows += other.total_rows;
        self.total_cells += other.total_cells;
        self.total_formulas += other.total_formulas;

        for (k, v) in &other.row_attrs {
            *self.row_attrs.entry(k.clone()).or_default() += v;
        }
        for (k, v) in &other.cell_attrs {
            *self.cell_attrs.entry(k.clone()).or_default() += v;
        }
        for (k, v) in &other.formula_attrs {
            *self.formula_attrs.entry(k.clone()).or_default() += v;
        }
        for (k, v) in &other.worksheet_children {
            *self.worksheet_children.entry(k.clone()).or_default() += v;
        }
    }

    /// Print human-readable statistics.
    pub fn print_report(&self) {
        println!("\n=== SML ATTRIBUTE USAGE STATISTICS ===\n");
        println!("Files analyzed:  {}", self.files_analyzed);
        println!("Worksheets:      {}", self.total_worksheets);
        println!("Rows:            {}", self.total_rows);
        println!("Cells:           {}", self.total_cells);
        println!("Formulas:        {}", self.total_formulas);

        println!("\n--- ROW ATTRIBUTES ({} rows) ---", self.total_rows);
        print_attr_stats(&self.row_attrs, self.total_rows);

        println!("\n--- CELL ATTRIBUTES ({} cells) ---", self.total_cells);
        print_attr_stats(&self.cell_attrs, self.total_cells);

        println!(
            "\n--- FORMULA ATTRIBUTES ({} formulas) ---",
            self.total_formulas
        );
        print_attr_stats(&self.formula_attrs, self.total_formulas);

        println!(
            "\n--- WORKSHEET CHILD ELEMENTS ({} worksheets) ---",
            self.total_worksheets
        );
        print_attr_stats(&self.worksheet_children, self.total_worksheets);
    }
}

fn print_attr_stats(attrs: &HashMap<String, u64>, total: u64) {
    let mut sorted: Vec<_> = attrs.iter().collect();
    sorted.sort_by(|a, b| b.1.cmp(a.1));

    for (attr, count) in sorted {
        let pct = if total > 0 {
            (*count as f64 / total as f64) * 100.0
        } else {
            0.0
        };
        let marker = if pct > 50.0 {
            "***"
        } else if pct > 10.0 {
            "**"
        } else if pct > 1.0 {
            "*"
        } else {
            ""
        };
        println!("  {:20} {:>10} ({:>6.2}%) {}", attr, count, pct, marker);
    }
}

/// Analyze a single XLSX file and return its attribute statistics.
pub fn analyze_xlsx_file(path: &Path) -> Result<SmlAttrStats, String> {
    let file = File::open(path).map_err(|e| e.to_string())?;
    let mut archive = ZipArchive::new(BufReader::new(file)).map_err(|e| e.to_string())?;

    let mut stats = SmlAttrStats::new();
    stats.files_analyzed = 1;

    // Find and analyze all worksheet XML files
    let names: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            archive.by_index(i).ok().and_then(|f| {
                let name = f.name().to_string();
                if name.contains("sheet") && name.ends_with(".xml") && !name.contains("rels") {
                    Some(name)
                } else {
                    None
                }
            })
        })
        .collect();

    for name in names {
        if let Ok(mut entry) = archive.by_name(&name) {
            let mut xml = String::new();
            if entry.read_to_string(&mut xml).is_ok() {
                analyze_worksheet_xml(&xml, &mut stats);
            }
        }
    }

    Ok(stats)
}

/// Analyze worksheet XML content and accumulate statistics.
fn analyze_worksheet_xml(xml: &str, stats: &mut SmlAttrStats) {
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut in_worksheet = false;
    let mut depth = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                if name == "worksheet" {
                    in_worksheet = true;
                    stats.total_worksheets += 1;
                    depth = 0;
                }

                if in_worksheet && depth == 1 {
                    *stats.worksheet_children.entry(name.clone()).or_default() += 1;
                }

                process_element(&name, &e, stats);

                if in_worksheet {
                    depth += 1;
                }
            }
            Ok(Event::Empty(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                if in_worksheet && depth == 1 {
                    *stats.worksheet_children.entry(name.clone()).or_default() += 1;
                }

                process_element(&name, &e, stats);
            }
            Ok(Event::End(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                if name == "worksheet" {
                    in_worksheet = false;
                }
                if in_worksheet {
                    depth -= 1;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn process_element(name: &str, e: &quick_xml::events::BytesStart, stats: &mut SmlAttrStats) {
    match name {
        "row" => {
            stats.total_rows += 1;
            for attr in e.attributes().filter_map(|a| a.ok()) {
                let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                *stats.row_attrs.entry(key).or_default() += 1;
            }
        }
        "c" => {
            stats.total_cells += 1;
            for attr in e.attributes().filter_map(|a| a.ok()) {
                let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                *stats.cell_attrs.entry(key).or_default() += 1;
            }
        }
        "f" => {
            stats.total_formulas += 1;
            for attr in e.attributes().filter_map(|a| a.ok()) {
                let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                *stats.formula_attrs.entry(key).or_default() += 1;
            }
        }
        _ => {}
    }
}

/// Analyze a directory of XLSX files.
pub fn analyze_xlsx_directory(dir: &Path, limit: Option<usize>) -> SmlAttrStats {
    use rayon::prelude::*;
    use std::sync::atomic::{AtomicUsize, Ordering};

    let mut files = Vec::new();
    collect_xlsx_files_recursive(dir, &mut files);

    if let Some(max) = limit {
        files.truncate(max);
    }

    let total = files.len();
    eprintln!("Found {} XLSX files to analyze", total);

    let processed = AtomicUsize::new(0);

    let results: Vec<SmlAttrStats> = files
        .par_iter()
        .filter_map(|path| {
            let count = processed.fetch_add(1, Ordering::Relaxed) + 1;
            if count.is_multiple_of(100) {
                eprintln!("  Processed {}/{}", count, total);
            }
            analyze_xlsx_file(path).ok()
        })
        .collect();

    let mut combined = SmlAttrStats::new();
    for stats in results {
        combined.merge(&stats);
    }

    combined
}

fn collect_xlsx_files_recursive(dir: &Path, files: &mut Vec<std::path::PathBuf>) {
    if let Ok(entries) = std::fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_xlsx_files_recursive(&path, files);
            } else if path
                .extension()
                .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx"))
            {
                files.push(path);
            }
        }
    }
}
