//! SQLite persistence for corpus analysis results.
//!
//! This module provides storage and retrieval of document analysis results,
//! enabling incremental analysis, querying, and statistical aggregation.

use crate::{AnalyzedError, DocumentFeatures};
use rusqlite::{Connection, Result as SqlResult, params};
use std::path::Path;

/// Database for storing corpus analysis results.
pub struct CorpusDatabase {
    conn: Connection,
}

/// A stored document analysis result.
#[derive(Debug, Clone)]
pub struct StoredDocument {
    /// Database ID.
    pub id: i64,
    /// File path within corpus.
    pub path: String,
    /// Corpus name/identifier.
    pub corpus: String,
    /// Whether parsing succeeded.
    pub parse_success: bool,
    /// Error details if parsing failed.
    pub parse_error: Option<StoredError>,
    /// Feature data if parsing succeeded.
    pub features: Option<DocumentFeatures>,
}

/// Stored error information.
#[derive(Debug, Clone)]
pub struct StoredError {
    /// Error category string.
    pub category: String,
    /// Error message.
    pub message: String,
    /// Raw error string.
    pub raw_error: String,
}

impl CorpusDatabase {
    /// Open or create a database at the given path.
    pub fn open(path: impl AsRef<Path>) -> SqlResult<Self> {
        let conn = Connection::open(path)?;
        let db = Self { conn };
        db.init_schema()?;
        Ok(db)
    }

    /// Create an in-memory database (useful for testing).
    pub fn in_memory() -> SqlResult<Self> {
        let conn = Connection::open_in_memory()?;
        let db = Self { conn };
        db.init_schema()?;
        Ok(db)
    }

    /// Initialize the database schema.
    fn init_schema(&self) -> SqlResult<()> {
        self.conn.execute_batch(
            r#"
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY,
                path TEXT NOT NULL,
                corpus TEXT NOT NULL,
                parse_success INTEGER NOT NULL,
                error_category TEXT,
                error_message TEXT,
                error_raw TEXT,
                features_json TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(path, corpus)
            );

            CREATE INDEX IF NOT EXISTS idx_documents_corpus ON documents(corpus);
            CREATE INDEX IF NOT EXISTS idx_documents_success ON documents(parse_success);
            CREATE INDEX IF NOT EXISTS idx_documents_error_category ON documents(error_category);

            CREATE TABLE IF NOT EXISTS corpus_stats (
                id INTEGER PRIMARY KEY,
                corpus TEXT NOT NULL UNIQUE,
                total_documents INTEGER NOT NULL DEFAULT 0,
                successes INTEGER NOT NULL DEFAULT 0,
                failures INTEGER NOT NULL DEFAULT 0,
                stats_json TEXT,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            "#,
        )
    }

    /// Insert or update a successful document analysis.
    pub fn insert_success(
        &self,
        path: &str,
        corpus: &str,
        features: &DocumentFeatures,
    ) -> SqlResult<i64> {
        let features_json = serde_json::to_string(features).unwrap_or_default();

        self.conn.execute(
            r#"
            INSERT INTO documents (path, corpus, parse_success, features_json)
            VALUES (?1, ?2, 1, ?3)
            ON CONFLICT(path, corpus) DO UPDATE SET
                parse_success = 1,
                error_category = NULL,
                error_message = NULL,
                error_raw = NULL,
                features_json = ?3,
                created_at = CURRENT_TIMESTAMP
            "#,
            params![path, corpus, features_json],
        )?;

        Ok(self.conn.last_insert_rowid())
    }

    /// Insert or update a failed document analysis.
    pub fn insert_failure(
        &self,
        path: &str,
        corpus: &str,
        error: &AnalyzedError,
    ) -> SqlResult<i64> {
        self.conn.execute(
            r#"
            INSERT INTO documents (path, corpus, parse_success, error_category, error_message, error_raw)
            VALUES (?1, ?2, 0, ?3, ?4, ?5)
            ON CONFLICT(path, corpus) DO UPDATE SET
                parse_success = 0,
                error_category = ?3,
                error_message = ?4,
                error_raw = ?5,
                features_json = NULL,
                created_at = CURRENT_TIMESTAMP
            "#,
            params![
                path,
                corpus,
                error.category.as_str(),
                error.message,
                error.raw_error
            ],
        )?;

        Ok(self.conn.last_insert_rowid())
    }

    /// Get a document by path and corpus.
    pub fn get_document(&self, path: &str, corpus: &str) -> SqlResult<Option<StoredDocument>> {
        let mut stmt = self.conn.prepare(
            r#"
            SELECT id, path, corpus, parse_success, error_category, error_message, error_raw, features_json
            FROM documents
            WHERE path = ?1 AND corpus = ?2
            "#,
        )?;

        let mut rows = stmt.query(params![path, corpus])?;

        if let Some(row) = rows.next()? {
            Ok(Some(self.row_to_document(row)?))
        } else {
            Ok(None)
        }
    }

    /// Check if a document exists in the database.
    pub fn document_exists(&self, path: &str, corpus: &str) -> SqlResult<bool> {
        let count: i64 = self.conn.query_row(
            "SELECT COUNT(*) FROM documents WHERE path = ?1 AND corpus = ?2",
            params![path, corpus],
            |row| row.get(0),
        )?;
        Ok(count > 0)
    }

    /// Get all documents for a corpus.
    pub fn get_corpus_documents(&self, corpus: &str) -> SqlResult<Vec<StoredDocument>> {
        let mut stmt = self.conn.prepare(
            r#"
            SELECT id, path, corpus, parse_success, error_category, error_message, error_raw, features_json
            FROM documents
            WHERE corpus = ?1
            ORDER BY path
            "#,
        )?;

        let mut docs = Vec::new();
        let mut rows = stmt.query(params![corpus])?;

        while let Some(row) = rows.next()? {
            docs.push(self.row_to_document(row)?);
        }

        Ok(docs)
    }

    /// Get all failed documents for a corpus.
    pub fn get_failures(&self, corpus: &str) -> SqlResult<Vec<StoredDocument>> {
        let mut stmt = self.conn.prepare(
            r#"
            SELECT id, path, corpus, parse_success, error_category, error_message, error_raw, features_json
            FROM documents
            WHERE corpus = ?1 AND parse_success = 0
            ORDER BY error_category, path
            "#,
        )?;

        let mut docs = Vec::new();
        let mut rows = stmt.query(params![corpus])?;

        while let Some(row) = rows.next()? {
            docs.push(self.row_to_document(row)?);
        }

        Ok(docs)
    }

    /// Get documents by error category.
    pub fn get_by_error_category(
        &self,
        corpus: &str,
        category: &str,
    ) -> SqlResult<Vec<StoredDocument>> {
        let mut stmt = self.conn.prepare(
            r#"
            SELECT id, path, corpus, parse_success, error_category, error_message, error_raw, features_json
            FROM documents
            WHERE corpus = ?1 AND error_category = ?2
            ORDER BY path
            "#,
        )?;

        let mut docs = Vec::new();
        let mut rows = stmt.query(params![corpus, category])?;

        while let Some(row) = rows.next()? {
            docs.push(self.row_to_document(row)?);
        }

        Ok(docs)
    }

    /// Get corpus statistics.
    pub fn get_corpus_stats(&self, corpus: &str) -> SqlResult<CorpusStats> {
        let (total, successes, failures): (i64, i64, i64) = self.conn.query_row(
            r#"
            SELECT
                COUNT(*),
                SUM(CASE WHEN parse_success = 1 THEN 1 ELSE 0 END),
                SUM(CASE WHEN parse_success = 0 THEN 1 ELSE 0 END)
            FROM documents
            WHERE corpus = ?1
            "#,
            params![corpus],
            |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?)),
        )?;

        // Get error category breakdown
        let mut stmt = self.conn.prepare(
            r#"
            SELECT error_category, COUNT(*)
            FROM documents
            WHERE corpus = ?1 AND parse_success = 0
            GROUP BY error_category
            ORDER BY COUNT(*) DESC
            "#,
        )?;

        let mut error_categories = Vec::new();
        let mut rows = stmt.query(params![corpus])?;

        while let Some(row) = rows.next()? {
            let category: String = row.get(0)?;
            let count: i64 = row.get(1)?;
            error_categories.push((category, count as u64));
        }

        Ok(CorpusStats {
            total_documents: total as u64,
            successes: successes as u64,
            failures: failures as u64,
            error_categories,
        })
    }

    /// List all corpora in the database.
    pub fn list_corpora(&self) -> SqlResult<Vec<String>> {
        let mut stmt = self
            .conn
            .prepare("SELECT DISTINCT corpus FROM documents ORDER BY corpus")?;

        let mut corpora = Vec::new();
        let mut rows = stmt.query([])?;

        while let Some(row) = rows.next()? {
            corpora.push(row.get(0)?);
        }

        Ok(corpora)
    }

    /// Delete all documents for a corpus.
    pub fn delete_corpus(&self, corpus: &str) -> SqlResult<usize> {
        self.conn
            .execute("DELETE FROM documents WHERE corpus = ?1", params![corpus])
    }

    /// Convert a database row to a StoredDocument.
    fn row_to_document(&self, row: &rusqlite::Row) -> SqlResult<StoredDocument> {
        let id: i64 = row.get(0)?;
        let path: String = row.get(1)?;
        let corpus: String = row.get(2)?;
        let parse_success: bool = row.get(3)?;
        let error_category: Option<String> = row.get(4)?;
        let error_message: Option<String> = row.get(5)?;
        let error_raw: Option<String> = row.get(6)?;
        let features_json: Option<String> = row.get(7)?;

        let parse_error =
            if let (Some(cat), Some(msg), Some(raw)) = (error_category, error_message, error_raw) {
                Some(StoredError {
                    category: cat,
                    message: msg,
                    raw_error: raw,
                })
            } else {
                None
            };

        let features = features_json.and_then(|json| serde_json::from_str(&json).ok());

        Ok(StoredDocument {
            id,
            path,
            corpus,
            parse_success,
            parse_error,
            features,
        })
    }
}

/// Summary statistics for a corpus.
#[derive(Debug, Clone)]
pub struct CorpusStats {
    /// Total documents in corpus.
    pub total_documents: u64,
    /// Successfully parsed documents.
    pub successes: u64,
    /// Failed documents.
    pub failures: u64,
    /// Error category breakdown (category, count).
    pub error_categories: Vec<(String, u64)>,
}

impl CorpusStats {
    /// Get success percentage.
    pub fn success_rate(&self) -> f64 {
        if self.total_documents == 0 {
            0.0
        } else {
            (self.successes as f64 / self.total_documents as f64) * 100.0
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ErrorCategory;

    #[test]
    fn test_create_database() {
        let db = CorpusDatabase::in_memory().unwrap();
        let corpora = db.list_corpora().unwrap();
        assert!(corpora.is_empty());
    }

    #[test]
    fn test_insert_and_retrieve_success() {
        let db = CorpusDatabase::in_memory().unwrap();

        let features = DocumentFeatures {
            paragraph_count: 10,
            table_count: 2,
            has_bold: true,
            ..Default::default()
        };

        db.insert_success("/path/to/doc.docx", "test-corpus", &features)
            .unwrap();

        let doc = db
            .get_document("/path/to/doc.docx", "test-corpus")
            .unwrap()
            .unwrap();

        assert!(doc.parse_success);
        assert_eq!(doc.path, "/path/to/doc.docx");
        assert_eq!(doc.corpus, "test-corpus");
        assert!(doc.parse_error.is_none());
        assert!(doc.features.is_some());

        let f = doc.features.unwrap();
        assert_eq!(f.paragraph_count, 10);
        assert_eq!(f.table_count, 2);
        assert!(f.has_bold);
    }

    #[test]
    fn test_insert_and_retrieve_failure() {
        let db = CorpusDatabase::in_memory().unwrap();

        let error = AnalyzedError {
            category: ErrorCategory::MissingRequiredPart,
            subcategory: Some("/word/document.xml".to_string()),
            message: "Missing document.xml".to_string(),
            location: None,
            raw_error: "Missing required part: /word/document.xml".to_string(),
        };

        db.insert_failure("/path/to/bad.docx", "test-corpus", &error)
            .unwrap();

        let doc = db
            .get_document("/path/to/bad.docx", "test-corpus")
            .unwrap()
            .unwrap();

        assert!(!doc.parse_success);
        assert!(doc.features.is_none());
        assert!(doc.parse_error.is_some());

        let err = doc.parse_error.unwrap();
        assert_eq!(err.category, "missing_required_part");
        assert_eq!(err.message, "Missing document.xml");
    }

    #[test]
    fn test_corpus_stats() {
        let db = CorpusDatabase::in_memory().unwrap();

        // Insert some successes
        for i in 0..8 {
            let features = DocumentFeatures::default();
            db.insert_success(&format!("/doc{}.docx", i), "test", &features)
                .unwrap();
        }

        // Insert some failures
        let error = AnalyzedError {
            category: ErrorCategory::XmlMalformed,
            subcategory: None,
            message: "Bad XML".to_string(),
            location: None,
            raw_error: "XML error".to_string(),
        };
        db.insert_failure("/bad1.docx", "test", &error).unwrap();
        db.insert_failure("/bad2.docx", "test", &error).unwrap();

        let stats = db.get_corpus_stats("test").unwrap();
        assert_eq!(stats.total_documents, 10);
        assert_eq!(stats.successes, 8);
        assert_eq!(stats.failures, 2);
        assert!((stats.success_rate() - 80.0).abs() < 0.01);
    }

    #[test]
    fn test_document_exists() {
        let db = CorpusDatabase::in_memory().unwrap();

        assert!(!db.document_exists("/doc.docx", "test").unwrap());

        db.insert_success("/doc.docx", "test", &DocumentFeatures::default())
            .unwrap();

        assert!(db.document_exists("/doc.docx", "test").unwrap());
        assert!(!db.document_exists("/doc.docx", "other").unwrap());
    }

    #[test]
    fn test_upsert_behavior() {
        let db = CorpusDatabase::in_memory().unwrap();

        // Insert as failure first
        let error = AnalyzedError {
            category: ErrorCategory::Unknown,
            subcategory: None,
            message: "Unknown error".to_string(),
            location: None,
            raw_error: "Error".to_string(),
        };
        db.insert_failure("/doc.docx", "test", &error).unwrap();

        let doc = db.get_document("/doc.docx", "test").unwrap().unwrap();
        assert!(!doc.parse_success);

        // Update to success
        let features = DocumentFeatures {
            paragraph_count: 5,
            ..Default::default()
        };
        db.insert_success("/doc.docx", "test", &features).unwrap();

        let doc = db.get_document("/doc.docx", "test").unwrap().unwrap();
        assert!(doc.parse_success);
        assert!(doc.parse_error.is_none());
        assert_eq!(doc.features.unwrap().paragraph_count, 5);
    }
}
