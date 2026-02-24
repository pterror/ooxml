//! Extension traits for DrawingML types.
//!
//! Provides convenience methods for working with generated DML types.

#[cfg(feature = "dml-text")]
use crate::types::*;

/// Extension trait for [`TextBody`] providing convenience methods.
#[cfg(feature = "dml-text")]
pub trait TextBodyExt {
    /// Get all paragraphs in the text body.
    fn paragraphs(&self) -> &[TextParagraph];

    /// Extract all text from the text body.
    fn text(&self) -> String;
}

#[cfg(feature = "dml-text")]
impl TextBodyExt for TextBody {
    fn paragraphs(&self) -> &[TextParagraph] {
        &self.p
    }

    fn text(&self) -> String {
        self.p
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// Extension trait for [`TextParagraph`] providing convenience methods.
#[cfg(feature = "dml-text")]
pub trait TextParagraphExt {
    /// Get all text runs in the paragraph.
    fn runs(&self) -> Vec<&TextRun>;

    /// Extract all text from the paragraph.
    fn text(&self) -> String;

    /// Get the paragraph level (for bullets/numbering).
    fn level(&self) -> Option<i32>;

    /// Get the text alignment.
    fn alignment(&self) -> Option<STTextAlignType>;
}

#[cfg(feature = "dml-text")]
impl TextParagraphExt for TextParagraph {
    fn runs(&self) -> Vec<&TextRun> {
        self.text_run
            .iter()
            .filter_map(|tr| match tr {
                EGTextRun::R(run) => Some(run.as_ref()),
                _ => None,
            })
            .collect()
    }

    fn text(&self) -> String {
        self.text_run
            .iter()
            .filter_map(|tr| match tr {
                EGTextRun::R(run) => Some(run.t.as_str()),
                EGTextRun::Br(_) => Some("\n"),
                EGTextRun::Fld(fld) => fld.t.as_deref(),
            })
            .collect()
    }

    fn level(&self) -> Option<i32> {
        self.p_pr.as_ref().and_then(|p| p.lvl)
    }

    fn alignment(&self) -> Option<STTextAlignType> {
        self.p_pr.as_ref().and_then(|p| p.algn)
    }
}

/// Extension trait for [`TextRun`] providing convenience methods.
#[cfg(feature = "dml-text")]
pub trait TextRunExt {
    /// Get the text content.
    fn text(&self) -> &str;

    /// Check if the text is bold.
    fn is_bold(&self) -> bool;

    /// Check if the text is italic.
    fn is_italic(&self) -> bool;

    /// Check if the text is underlined.
    fn is_underlined(&self) -> bool;

    /// Check if the text has strikethrough.
    fn is_strikethrough(&self) -> bool;

    /// Get the font size in hundredths of a point.
    fn font_size(&self) -> Option<i32>;

    /// Check if the run has a hyperlink.
    fn has_hyperlink(&self) -> bool;

    /// Get the hyperlink relationship ID.
    fn hyperlink_rel_id(&self) -> Option<&str>;
}

#[cfg(feature = "dml-text")]
impl TextRunExt for TextRun {
    fn text(&self) -> &str {
        &self.t
    }

    fn is_bold(&self) -> bool {
        self.r_pr.as_ref().and_then(|p| p.b).unwrap_or(false)
    }

    fn is_italic(&self) -> bool {
        self.r_pr.as_ref().and_then(|p| p.i).unwrap_or(false)
    }

    fn is_underlined(&self) -> bool {
        self.r_pr
            .as_ref()
            .and_then(|p| p.u.as_ref())
            .is_some_and(|u| *u != STTextUnderlineType::None)
    }

    fn is_strikethrough(&self) -> bool {
        self.r_pr
            .as_ref()
            .and_then(|p| p.strike.as_ref())
            .is_some_and(|s| *s != STTextStrikeType::NoStrike)
    }

    fn font_size(&self) -> Option<i32> {
        self.r_pr.as_ref().and_then(|p| p.sz)
    }

    fn has_hyperlink(&self) -> bool {
        self.r_pr
            .as_ref()
            .and_then(|p| p.hlink_click.as_ref())
            .is_some()
    }

    fn hyperlink_rel_id(&self) -> Option<&str> {
        self.r_pr
            .as_ref()
            .and_then(|p| p.hlink_click.as_ref())
            .and_then(|h| h.id.as_deref())
    }
}

/// Extension trait for [`CTTable`] providing convenience methods.
#[cfg(feature = "dml-tables")]
pub trait TableExt {
    /// Get all rows in the table.
    fn rows(&self) -> &[CTTableRow];

    /// Get the number of rows.
    fn row_count(&self) -> usize;

    /// Get the number of columns (from grid, or first row if empty).
    fn col_count(&self) -> usize;

    /// Get a cell by row and column index (0-based).
    fn cell(&self, row: usize, col: usize) -> Option<&CTTableCell>;

    /// Get all cell text as a 2D vector.
    fn to_text_grid(&self) -> Vec<Vec<String>>;

    /// Get plain text representation (tab-separated values).
    fn text(&self) -> String;
}

#[cfg(feature = "dml-tables")]
impl TableExt for CTTable {
    fn rows(&self) -> &[CTTableRow] {
        &self.tr
    }

    fn row_count(&self) -> usize {
        self.tr.len()
    }

    fn col_count(&self) -> usize {
        self.tbl_grid.grid_col.len()
    }

    fn cell(&self, row: usize, col: usize) -> Option<&CTTableCell> {
        self.tr.get(row).and_then(|r| r.tc.get(col))
    }

    fn to_text_grid(&self) -> Vec<Vec<String>> {
        self.tr
            .iter()
            .map(|row| row.tc.iter().map(|c| c.text()).collect())
            .collect()
    }

    fn text(&self) -> String {
        self.tr
            .iter()
            .map(|row| {
                row.tc
                    .iter()
                    .map(|c| c.text())
                    .collect::<Vec<_>>()
                    .join("\t")
            })
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// Extension trait for [`CTTableRow`] providing convenience methods.
#[cfg(feature = "dml-tables")]
pub trait TableRowExt {
    /// Get all cells in this row.
    fn cells(&self) -> &[CTTableCell];

    /// Get a cell by column index (0-based).
    fn cell(&self, col: usize) -> Option<&CTTableCell>;

    /// Get the row height in EMUs (if specified).
    fn height_emu(&self) -> Option<i64>;
}

#[cfg(feature = "dml-tables")]
impl TableRowExt for CTTableRow {
    fn cells(&self) -> &[CTTableCell] {
        &self.tc
    }

    fn cell(&self, col: usize) -> Option<&CTTableCell> {
        self.tc.get(col)
    }

    fn height_emu(&self) -> Option<i64> {
        self.height.parse::<i64>().ok()
    }
}

/// Extension trait for [`CTTableCell`] providing convenience methods.
#[cfg(feature = "dml-tables")]
pub trait TableCellExt {
    /// Get the text body (paragraphs) if present.
    fn text_body(&self) -> Option<&TextBody>;

    /// Get the cell text (paragraphs joined with newlines).
    fn text(&self) -> String;

    /// Get the row span (number of rows this cell spans).
    fn row_span(&self) -> u32;

    /// Get the column span (number of columns this cell spans).
    fn col_span(&self) -> u32;

    /// Check if this cell spans multiple rows.
    fn has_row_span(&self) -> bool;

    /// Check if this cell spans multiple columns.
    fn has_col_span(&self) -> bool;

    /// Check if this cell is merged horizontally (continuation of previous cell).
    fn is_h_merge(&self) -> bool;

    /// Check if this cell is merged vertically (continuation of cell above).
    fn is_v_merge(&self) -> bool;
}

#[cfg(feature = "dml-tables")]
impl TableCellExt for CTTableCell {
    fn text_body(&self) -> Option<&TextBody> {
        self.tx_body.as_deref()
    }

    fn text(&self) -> String {
        self.tx_body
            .as_ref()
            .map(|tb| tb.text())
            .unwrap_or_default()
    }

    fn row_span(&self) -> u32 {
        self.row_span.map(|s| s.max(1) as u32).unwrap_or(1)
    }

    fn col_span(&self) -> u32 {
        self.grid_span.map(|s| s.max(1) as u32).unwrap_or(1)
    }

    fn has_row_span(&self) -> bool {
        self.row_span.is_some_and(|s| s > 1)
    }

    fn has_col_span(&self) -> bool {
        self.grid_span.is_some_and(|s| s > 1)
    }

    fn is_h_merge(&self) -> bool {
        self.h_merge.unwrap_or(false)
    }

    fn is_v_merge(&self) -> bool {
        self.v_merge.unwrap_or(false)
    }
}

/// The kind of chart contained in a plot area.
///
/// Corresponds to the chart element types defined in ECMA-376 §21.2.
#[cfg(feature = "dml-charts")]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartKind {
    /// Bar or column chart (CT_BarChart).
    Bar,
    /// 3D bar chart (CT_Bar3DChart).
    Bar3D,
    /// Line chart (CT_LineChart).
    Line,
    /// 3D line chart (CT_Line3DChart).
    Line3D,
    /// Pie chart (CT_PieChart).
    Pie,
    /// 3D pie chart (CT_Pie3DChart).
    Pie3D,
    /// Scatter / XY chart (CT_ScatterChart).
    Scatter,
    /// Area chart (CT_AreaChart).
    Area,
    /// 3D area chart (CT_Area3DChart).
    Area3D,
    /// Bubble chart (CT_BubbleChart).
    Bubble,
    /// Doughnut chart (CT_DoughnutChart).
    Doughnut,
    /// Radar / spider chart (CT_RadarChart).
    Radar,
    /// Stock chart (CT_StockChart).
    Stock,
    /// Surface chart (CT_SurfaceChart).
    Surface,
    /// 3D surface chart (CT_Surface3DChart).
    Surface3D,
    /// Pie-of-pie or bar-of-pie chart (CT_OfPieChart).
    OfPie,
}

/// Extension trait for [`ChartSpace`] (the root element of a chart part).
///
/// Corresponds to ECMA-376 §21.2.2.29 (CT_ChartSpace).
#[cfg(feature = "dml-charts")]
pub trait ChartSpaceExt {
    /// The inner chart definition.
    fn chart(&self) -> &crate::types::Chart;

    /// All chart kinds present in this chart space's plot area.
    ///
    /// A chart space can contain multiple chart types (e.g. a combined bar+line chart).
    fn chart_types(&self) -> Vec<ChartKind>;

    /// The chart title text, if the title contains rich text content.
    ///
    /// Returns `None` if there is no title, or if the title references an external
    /// cell range rather than inline text.
    fn title_text(&self) -> Option<String>;
}

#[cfg(feature = "dml-charts")]
impl ChartSpaceExt for crate::types::ChartSpace {
    fn chart(&self) -> &crate::types::Chart {
        &self.chart
    }

    fn chart_types(&self) -> Vec<ChartKind> {
        self.chart.chart_types()
    }

    fn title_text(&self) -> Option<String> {
        self.chart.title_text()
    }
}

/// Extension trait for [`Chart`] providing convenience access to chart content.
///
/// Corresponds to ECMA-376 §21.2.2.27 (CT_Chart).
#[cfg(feature = "dml-charts")]
pub trait ChartExt {
    /// The plot area containing the chart series and axes.
    fn plot_area(&self) -> &crate::types::PlotArea;

    /// The chart legend, if present.
    fn legend(&self) -> Option<&crate::types::Legend>;

    /// All chart kinds present in this chart's plot area.
    fn chart_types(&self) -> Vec<ChartKind>;

    /// The chart title text, if the title contains rich text content.
    fn title_text(&self) -> Option<String>;
}

#[cfg(feature = "dml-charts")]
impl ChartExt for crate::types::Chart {
    fn plot_area(&self) -> &crate::types::PlotArea {
        &self.plot_area
    }

    fn legend(&self) -> Option<&crate::types::Legend> {
        self.legend.as_deref()
    }

    fn chart_types(&self) -> Vec<ChartKind> {
        self.plot_area.chart_types()
    }

    fn title_text(&self) -> Option<String> {
        self.title.as_deref().and_then(|t| t.title_text())
    }
}

/// Extension trait for [`PlotArea`] providing access to contained chart types.
///
/// Corresponds to ECMA-376 §21.2.2.145 (CT_PlotArea).
#[cfg(feature = "dml-charts")]
pub trait PlotAreaExt {
    /// All chart kinds present in this plot area.
    ///
    /// Returns one entry per chart type present. A combined chart (e.g. bar + line)
    /// returns multiple entries in the order they appear in the XML.
    fn chart_types(&self) -> Vec<ChartKind>;
}

#[cfg(feature = "dml-charts")]
impl PlotAreaExt for crate::types::PlotArea {
    fn chart_types(&self) -> Vec<ChartKind> {
        let mut kinds = Vec::new();
        if !self.bar_chart.is_empty() {
            kinds.push(ChartKind::Bar);
        }
        if !self.bar3_d_chart.is_empty() {
            kinds.push(ChartKind::Bar3D);
        }
        if !self.line_chart.is_empty() {
            kinds.push(ChartKind::Line);
        }
        if !self.line3_d_chart.is_empty() {
            kinds.push(ChartKind::Line3D);
        }
        if !self.pie_chart.is_empty() {
            kinds.push(ChartKind::Pie);
        }
        if !self.pie3_d_chart.is_empty() {
            kinds.push(ChartKind::Pie3D);
        }
        if !self.scatter_chart.is_empty() {
            kinds.push(ChartKind::Scatter);
        }
        if !self.area_chart.is_empty() {
            kinds.push(ChartKind::Area);
        }
        if !self.area3_d_chart.is_empty() {
            kinds.push(ChartKind::Area3D);
        }
        if !self.bubble_chart.is_empty() {
            kinds.push(ChartKind::Bubble);
        }
        if !self.doughnut_chart.is_empty() {
            kinds.push(ChartKind::Doughnut);
        }
        if !self.radar_chart.is_empty() {
            kinds.push(ChartKind::Radar);
        }
        if !self.stock_chart.is_empty() {
            kinds.push(ChartKind::Stock);
        }
        if !self.surface_chart.is_empty() {
            kinds.push(ChartKind::Surface);
        }
        if !self.surface3_d_chart.is_empty() {
            kinds.push(ChartKind::Surface3D);
        }
        if !self.of_pie_chart.is_empty() {
            kinds.push(ChartKind::OfPie);
        }
        kinds
    }
}

/// Extension trait for [`ChartTitle`] providing text extraction.
///
/// Corresponds to ECMA-376 §21.2.2.210 (CT_Title).
#[cfg(feature = "dml-charts")]
pub trait ChartTitleExt {
    /// Extract the title text if it is stored as inline rich text.
    ///
    /// Returns `None` if:
    /// - There is no `tx` child element.
    /// - The `tx` element references a cell range (via `strRef`) rather than inline text.
    fn title_text(&self) -> Option<String>;
}

#[cfg(all(feature = "dml-charts", feature = "dml-text"))]
impl ChartTitleExt for crate::types::ChartTitle {
    fn title_text(&self) -> Option<String> {
        self.tx.as_deref().and_then(|tx| {
            tx.rich.as_deref().map(|body| {
                body.p
                    .iter()
                    .map(|p| {
                        p.text_run
                            .iter()
                            .filter_map(|tr| match tr {
                                EGTextRun::R(run) => Some(run.t.as_str()),
                                EGTextRun::Br(_) => Some("\n"),
                                EGTextRun::Fld(fld) => fld.t.as_deref(),
                            })
                            .collect::<String>()
                    })
                    .collect::<Vec<_>>()
                    .join("\n")
            })
        })
    }
}

// When dml-text is not enabled, provide a fallback that always returns None.
#[cfg(all(feature = "dml-charts", not(feature = "dml-text")))]
impl ChartTitleExt for crate::types::ChartTitle {
    fn title_text(&self) -> Option<String> {
        None
    }
}

/// Extension trait for [`DataModel`] providing convenience access to SmartArt content.
///
/// Corresponds to ECMA-376 §21.4.2.8 (CT_DataModel).
#[cfg(feature = "dml-diagrams")]
pub trait DataModelExt {
    /// Returns all diagram points that represent actual content nodes.
    ///
    /// Filters out connector/transition points (`parTrans`, `sibTrans`, `pres`) and
    /// returns only points of type `node`, `asst`, or `doc`.
    fn content_points(&self) -> Vec<&crate::types::DiagramPoint>;

    /// Returns all connections between diagram points.
    fn connections(&self) -> Vec<&crate::types::DiagramConnection>;

    /// Extracts all text from diagram content nodes, in order.
    ///
    /// Each node's text paragraphs are joined with newlines. Nodes are separated
    /// by newlines in the returned string.
    fn text(&self) -> Vec<String>;
}

#[cfg(feature = "dml-diagrams")]
fn diagram_content_points(model: &crate::types::DataModel) -> Vec<&crate::types::DiagramPoint> {
    use crate::types::STPtType;
    model
        .pt_lst
        .pt
        .iter()
        .filter(|pt| {
            // Include points without a type (defaults to "node") and explicit
            // node/asst/doc types. Exclude parTrans, sibTrans, pres connectors.
            matches!(
                pt.r#type,
                None | Some(STPtType::Node) | Some(STPtType::Asst) | Some(STPtType::Doc)
            )
        })
        .collect()
}

#[cfg(all(feature = "dml-diagrams", feature = "dml-text"))]
impl DataModelExt for crate::types::DataModel {
    fn content_points(&self) -> Vec<&crate::types::DiagramPoint> {
        diagram_content_points(self)
    }

    fn connections(&self) -> Vec<&crate::types::DiagramConnection> {
        self.cxn_lst
            .as_deref()
            .map(|lst| lst.cxn.iter().collect())
            .unwrap_or_default()
    }

    fn text(&self) -> Vec<String> {
        self.content_points()
            .iter()
            .filter_map(|pt| {
                pt.t.as_deref().map(|body| {
                    body.p
                        .iter()
                        .map(|p| {
                            p.text_run
                                .iter()
                                .filter_map(|tr| match tr {
                                    crate::types::EGTextRun::R(run) => Some(run.t.as_str()),
                                    crate::types::EGTextRun::Br(_) => Some("\n"),
                                    crate::types::EGTextRun::Fld(fld) => fld.t.as_deref(),
                                })
                                .collect::<String>()
                        })
                        .collect::<Vec<_>>()
                        .join("\n")
                })
            })
            .filter(|s| !s.is_empty())
            .collect()
    }
}

// Fallback implementation when dml-text is not enabled: text() always returns empty.
#[cfg(all(feature = "dml-diagrams", not(feature = "dml-text")))]
impl DataModelExt for crate::types::DataModel {
    fn content_points(&self) -> Vec<&crate::types::DiagramPoint> {
        diagram_content_points(self)
    }

    fn connections(&self) -> Vec<&crate::types::DiagramConnection> {
        self.cxn_lst
            .as_deref()
            .map(|lst| lst.cxn.iter().collect())
            .unwrap_or_default()
    }

    fn text(&self) -> Vec<String> {
        Vec::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_paragraph_text() {
        let para = TextParagraph {
            #[cfg(feature = "dml-text")]
            p_pr: None,
            text_run: vec![
                EGTextRun::R(Box::new(TextRun {
                    #[cfg(feature = "dml-text")]
                    r_pr: None,
                    #[cfg(feature = "dml-text")]
                    t: "Hello ".to_string(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                })),
                EGTextRun::R(Box::new(TextRun {
                    #[cfg(feature = "dml-text")]
                    r_pr: None,
                    #[cfg(feature = "dml-text")]
                    t: "World".to_string(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                })),
            ],
            #[cfg(feature = "dml-text")]
            end_para_r_pr: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        assert_eq!(para.text(), "Hello World");
        assert_eq!(para.runs().len(), 2);
    }

    #[cfg(feature = "dml-text")]
    #[test]
    fn test_text_run_formatting() {
        let run = TextRun {
            r_pr: Some(Box::new(TextCharacterProperties {
                b: Some(true),
                i: Some(true),
                ..Default::default()
            })),
            t: "Bold Italic".to_string(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        assert!(run.is_bold());
        assert!(run.is_italic());
        assert!(!run.is_underlined());
        assert_eq!(run.text(), "Bold Italic");
    }

    #[cfg(feature = "dml-tables")]
    fn make_cell(text: &str) -> CTTableCell {
        CTTableCell {
            row_span: None,
            grid_span: None,
            h_merge: None,
            v_merge: None,
            id: None,
            tx_body: Some(Box::new(TextBody {
                body_pr: Box::new(CTTextBodyProperties::default()),
                lst_style: None,
                p: vec![TextParagraph {
                    #[cfg(feature = "dml-text")]
                    p_pr: None,
                    text_run: vec![EGTextRun::R(Box::new(TextRun {
                        #[cfg(feature = "dml-text")]
                        r_pr: None,
                        #[cfg(feature = "dml-text")]
                        t: text.to_string(),
                        #[cfg(feature = "extra-children")]
                        extra_children: Vec::new(),
                    }))],
                    #[cfg(feature = "dml-text")]
                    end_para_r_pr: None,
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                }],
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            })),
            tc_pr: None,
            ext_lst: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }
    }

    #[cfg(feature = "dml-tables")]
    fn make_table(data: &[&[&str]]) -> CTTable {
        let rows: Vec<CTTableRow> = data
            .iter()
            .map(|row_data| CTTableRow {
                height: "370840".to_string(), // ~0.26 inches
                tc: row_data.iter().map(|&text| make_cell(text)).collect(),
                ext_lst: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            })
            .collect();

        let col_count = data.first().map(|r| r.len()).unwrap_or(0);
        let grid_cols: Vec<CTTableCol> = (0..col_count)
            .map(|_| CTTableCol {
                width: "914400".to_string(), // 1 inch
                ext_lst: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            })
            .collect();

        CTTable {
            tbl_pr: None,
            tbl_grid: Box::new(CTTableGrid {
                grid_col: grid_cols,
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }),
            tr: rows,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }
    }

    #[cfg(feature = "dml-tables")]
    #[test]
    fn test_table_ext() {
        let table = make_table(&[&["A", "B", "C"], &["1", "2", "3"], &["X", "Y", "Z"]]);

        assert_eq!(table.row_count(), 3);
        assert_eq!(table.col_count(), 3);
        assert_eq!(table.rows().len(), 3);

        // Test cell access
        assert_eq!(table.cell(0, 0).unwrap().text(), "A");
        assert_eq!(table.cell(1, 1).unwrap().text(), "2");
        assert_eq!(table.cell(2, 2).unwrap().text(), "Z");
        assert!(table.cell(3, 0).is_none()); // Out of bounds

        // Test text grid
        let grid = table.to_text_grid();
        assert_eq!(
            grid,
            vec![
                vec!["A", "B", "C"],
                vec!["1", "2", "3"],
                vec!["X", "Y", "Z"],
            ]
        );

        // Test text output
        assert_eq!(table.text(), "A\tB\tC\n1\t2\t3\nX\tY\tZ");
    }

    #[cfg(feature = "dml-tables")]
    #[test]
    fn test_table_row_ext() {
        let table = make_table(&[&["Hello", "World"]]);
        let row = &table.tr[0];

        assert_eq!(row.cells().len(), 2);
        assert_eq!(row.cell(0).unwrap().text(), "Hello");
        assert_eq!(row.cell(1).unwrap().text(), "World");
        assert!(row.cell(2).is_none());
        assert_eq!(row.height_emu(), Some(370840));
    }

    #[cfg(feature = "dml-tables")]
    #[test]
    fn test_table_cell_ext() {
        let cell = make_cell("Test Content");

        assert_eq!(cell.text(), "Test Content");
        assert!(cell.text_body().is_some());
        assert_eq!(cell.row_span(), 1);
        assert_eq!(cell.col_span(), 1);
        assert!(!cell.has_row_span());
        assert!(!cell.has_col_span());
        assert!(!cell.is_h_merge());
        assert!(!cell.is_v_merge());
    }

    #[cfg(feature = "dml-tables")]
    #[test]
    fn test_table_cell_spanning() {
        let mut cell = make_cell("Merged");
        cell.row_span = Some(2);
        cell.grid_span = Some(3);
        cell.h_merge = Some(true);

        assert_eq!(cell.row_span(), 2);
        assert_eq!(cell.col_span(), 3);
        assert!(cell.has_row_span());
        assert!(cell.has_col_span());
        assert!(cell.is_h_merge());
        assert!(!cell.is_v_merge());
    }

    // -------------------------------------------------------------------------
    // Chart extension trait tests
    // -------------------------------------------------------------------------

    #[cfg(feature = "dml-charts")]
    fn make_chart_with_bar() -> crate::types::Chart {
        use crate::types::*;
        Chart {
            #[cfg(feature = "dml-charts")]
            title: None,
            #[cfg(feature = "dml-charts")]
            auto_title_deleted: None,
            #[cfg(feature = "dml-charts")]
            pivot_fmts: None,
            #[cfg(feature = "dml-charts")]
            view3_d: None,
            #[cfg(feature = "dml-charts")]
            floor: None,
            #[cfg(feature = "dml-charts")]
            side_wall: None,
            #[cfg(feature = "dml-charts")]
            back_wall: None,
            #[cfg(feature = "dml-charts")]
            plot_area: Box::new(PlotArea {
                #[cfg(feature = "dml-charts")]
                bar_chart: vec![BarChart {
                    #[cfg(feature = "dml-charts")]
                    bar_dir: Box::new(BarDirection::default()),
                    #[cfg(feature = "dml-charts")]
                    grouping: None,
                    #[cfg(feature = "dml-charts")]
                    vary_colors: None,
                    #[cfg(feature = "dml-charts")]
                    ser: Vec::new(),
                    #[cfg(feature = "dml-charts")]
                    d_lbls: None,
                    #[cfg(feature = "dml-charts")]
                    gap_width: None,
                    #[cfg(feature = "dml-charts")]
                    overlap: None,
                    #[cfg(feature = "dml-charts")]
                    ser_lines: Vec::new(),
                    #[cfg(feature = "dml-charts")]
                    ax_id: Vec::new(),
                    #[cfg(feature = "dml-charts")]
                    ext_lst: None,
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                }],
                ..Default::default()
            }),
            #[cfg(feature = "dml-charts")]
            legend: None,
            #[cfg(feature = "dml-charts")]
            plot_vis_only: None,
            #[cfg(feature = "dml-charts")]
            disp_blanks_as: None,
            #[cfg(feature = "dml-charts")]
            show_d_lbls_over_max: None,
            #[cfg(feature = "dml-charts")]
            ext_lst: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }
    }

    #[cfg(feature = "dml-charts")]
    #[test]
    fn test_chart_kind_from_bar_chart() {
        use crate::ext::{ChartExt, ChartKind};
        let chart = make_chart_with_bar();
        let kinds = chart.chart_types();
        assert_eq!(kinds, vec![ChartKind::Bar]);
    }

    #[cfg(feature = "dml-charts")]
    #[test]
    fn test_plot_area_empty_has_no_kinds() {
        use crate::ext::PlotAreaExt;
        use crate::types::PlotArea;
        let area = PlotArea::default();
        assert!(area.chart_types().is_empty());
    }

    #[cfg(feature = "dml-charts")]
    #[test]
    fn test_chart_space_delegates_to_chart() {
        use crate::ext::{ChartKind, ChartSpaceExt};
        use crate::types::ChartSpace;
        let space = ChartSpace {
            #[cfg(feature = "dml-charts")]
            date1904: None,
            #[cfg(feature = "dml-charts")]
            lang: None,
            #[cfg(feature = "dml-charts")]
            rounded_corners: None,
            #[cfg(feature = "dml-charts")]
            style: None,
            #[cfg(feature = "dml-charts")]
            clr_map_ovr: None,
            #[cfg(feature = "dml-charts")]
            pivot_source: None,
            #[cfg(feature = "dml-charts")]
            protection: None,
            #[cfg(feature = "dml-charts")]
            chart: Box::new(make_chart_with_bar()),
            #[cfg(feature = "dml-charts")]
            sp_pr: None,
            #[cfg(feature = "dml-charts")]
            tx_pr: None,
            #[cfg(feature = "dml-charts")]
            external_data: None,
            #[cfg(feature = "dml-charts")]
            print_settings: None,
            #[cfg(feature = "dml-charts")]
            user_shapes: None,
            #[cfg(feature = "dml-charts")]
            ext_lst: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        assert_eq!(space.chart_types(), vec![ChartKind::Bar]);
        assert!(space.title_text().is_none());
    }

    #[cfg(all(feature = "dml-charts", feature = "dml-text"))]
    #[test]
    fn test_chart_title_text_rich() {
        use crate::ext::ChartTitleExt;
        use crate::types::*;

        // Build a chart title with inline rich text
        let title = ChartTitle {
            #[cfg(feature = "dml-charts")]
            tx: Some(Box::new(ChartText {
                #[cfg(feature = "dml-charts")]
                str_ref: None,
                #[cfg(feature = "dml-charts")]
                rich: Some(Box::new(TextBody {
                    body_pr: Box::new(CTTextBodyProperties::default()),
                    lst_style: None,
                    p: vec![TextParagraph {
                        #[cfg(feature = "dml-text")]
                        p_pr: None,
                        text_run: vec![EGTextRun::R(Box::new(TextRun {
                            #[cfg(feature = "dml-text")]
                            r_pr: None,
                            #[cfg(feature = "dml-text")]
                            t: "Sales Report".to_string(),
                            #[cfg(feature = "extra-children")]
                            extra_children: Vec::new(),
                        }))],
                        #[cfg(feature = "dml-text")]
                        end_para_r_pr: None,
                        #[cfg(feature = "extra-children")]
                        extra_children: Vec::new(),
                    }],
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                })),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            })),
            #[cfg(feature = "dml-charts")]
            layout: None,
            #[cfg(feature = "dml-charts")]
            overlay: None,
            #[cfg(feature = "dml-charts")]
            sp_pr: None,
            #[cfg(feature = "dml-charts")]
            tx_pr: None,
            #[cfg(feature = "dml-charts")]
            ext_lst: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        assert_eq!(title.title_text(), Some("Sales Report".to_string()));
    }

    // -------------------------------------------------------------------------
    // DataModelExt tests
    // -------------------------------------------------------------------------

    #[cfg(feature = "dml-diagrams")]
    fn make_data_model(points: Vec<crate::types::DiagramPoint>) -> crate::types::DataModel {
        use crate::types::*;
        DataModel {
            #[cfg(feature = "dml-diagrams")]
            pt_lst: Box::new(DiagramPointList {
                pt: points,
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }),
            #[cfg(feature = "dml-diagrams")]
            cxn_lst: None,
            #[cfg(feature = "dml-diagrams")]
            bg: None,
            #[cfg(feature = "dml-diagrams")]
            whole: None,
            #[cfg(feature = "dml-diagrams")]
            ext_lst: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }
    }

    #[cfg(feature = "dml-diagrams")]
    fn make_diagram_point(
        id: &str,
        pt_type: Option<crate::types::STPtType>,
    ) -> crate::types::DiagramPoint {
        use crate::types::*;
        DiagramPoint {
            #[cfg(feature = "dml-diagrams")]
            model_id: id.to_string(),
            #[cfg(feature = "dml-diagrams")]
            r#type: pt_type,
            #[cfg(feature = "dml-diagrams")]
            cxn_id: None,
            #[cfg(feature = "dml-diagrams")]
            pr_set: None,
            #[cfg(feature = "dml-diagrams")]
            sp_pr: None,
            #[cfg(feature = "dml-diagrams")]
            t: None,
            #[cfg(feature = "dml-diagrams")]
            ext_lst: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }
    }

    #[cfg(feature = "dml-diagrams")]
    #[test]
    fn test_data_model_content_points() {
        use crate::ext::DataModelExt;
        use crate::types::STPtType;

        let points = vec![
            make_diagram_point("1", None),                     // node (default)
            make_diagram_point("2", Some(STPtType::Node)),     // explicit node
            make_diagram_point("3", Some(STPtType::Asst)),     // assistant
            make_diagram_point("4", Some(STPtType::ParTrans)), // connector — excluded
            make_diagram_point("5", Some(STPtType::SibTrans)), // connector — excluded
            make_diagram_point("6", Some(STPtType::Pres)),     // presentation — excluded
        ];

        let model = make_data_model(points);
        let content = model.content_points();

        // Only node/asst/doc types are returned
        assert_eq!(content.len(), 3);
        assert_eq!(content[0].model_id, "1");
        assert_eq!(content[1].model_id, "2");
        assert_eq!(content[2].model_id, "3");
    }

    #[cfg(feature = "dml-diagrams")]
    #[test]
    fn test_data_model_connections_empty() {
        use crate::ext::DataModelExt;

        let model = make_data_model(vec![]);
        assert!(model.connections().is_empty());
    }

    #[cfg(all(feature = "dml-diagrams", feature = "dml-text"))]
    #[test]
    fn test_data_model_text() {
        use crate::ext::DataModelExt;
        use crate::types::*;

        let mut pt = make_diagram_point("1", None);
        pt.t = Some(Box::new(TextBody {
            body_pr: Box::new(CTTextBodyProperties::default()),
            lst_style: None,
            p: vec![TextParagraph {
                #[cfg(feature = "dml-text")]
                p_pr: None,
                text_run: vec![EGTextRun::R(Box::new(TextRun {
                    #[cfg(feature = "dml-text")]
                    r_pr: None,
                    #[cfg(feature = "dml-text")]
                    t: "SmartArt Node".to_string(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                }))],
                #[cfg(feature = "dml-text")]
                end_para_r_pr: None,
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }],
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }));

        let model = make_data_model(vec![pt]);
        let texts = model.text();

        assert_eq!(texts.len(), 1);
        assert_eq!(texts[0], "SmartArt Node");
    }

    #[cfg(feature = "dml-charts")]
    #[test]
    fn test_chart_title_none_when_no_tx() {
        use crate::ext::ChartTitleExt;
        use crate::types::ChartTitle;
        let title = ChartTitle {
            #[cfg(feature = "dml-charts")]
            tx: None,
            #[cfg(feature = "dml-charts")]
            layout: None,
            #[cfg(feature = "dml-charts")]
            overlay: None,
            #[cfg(feature = "dml-charts")]
            sp_pr: None,
            #[cfg(feature = "dml-charts")]
            tx_pr: None,
            #[cfg(feature = "dml-charts")]
            ext_lst: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        assert!(title.title_text().is_none());
    }
}
