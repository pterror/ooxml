//! Extension traits for generated PML types.
//!
//! These traits provide convenient accessor methods for the generated types,
//! similar to the handwritten API but working with the schema-generated structures.

use crate::generated::*;
use ooxml_dml::ext::{TextBodyExt, TextParagraphExt};
use ooxml_dml::types::TextParagraph;

/// Extension trait for Shape (p:sp element).
pub trait ShapeExt {
    /// Get the shape name from cNvPr.
    fn name(&self) -> &str;

    /// Get the shape description/alt text from cNvPr.
    fn description(&self) -> Option<&str>;

    /// Get the text body if present.
    fn text_body(&self) -> Option<&ooxml_dml::types::TextBody>;

    /// Get the text paragraphs from the text body.
    fn paragraphs(&self) -> &[TextParagraph];

    /// Get all text content joined by newlines.
    fn text(&self) -> Option<String>;

    /// Check if the shape has text content.
    fn has_text(&self) -> bool;

    /// Get the drawing element ID (`cNvPr@id`).
    fn shape_id(&self) -> u32;

    /// Check if this shape is a placeholder (has a `ph` element).
    fn is_placeholder(&self) -> bool;

    /// Get the placeholder type (title, body, etc.) if this is a placeholder.
    fn placeholder_type(&self) -> Option<&STPlaceholderType>;

    /// Get the placeholder index if this is a placeholder.
    fn placeholder_index(&self) -> Option<u32>;

    /// Get the position offset in EMU as (x, y).
    fn offset_emu(&self) -> Option<(i64, i64)>;

    /// Get the extent (width, height) in EMU.
    fn extent_emu(&self) -> Option<(i64, i64)>;

    /// Get the rotation angle in degrees.
    fn rotation_angle_deg(&self) -> Option<f64>;
}

impl ShapeExt for Shape {
    fn name(&self) -> &str {
        &self.non_visual_properties.c_nv_pr.name
    }

    fn description(&self) -> Option<&str> {
        self.non_visual_properties.c_nv_pr.descr.as_deref()
    }

    fn text_body(&self) -> Option<&ooxml_dml::types::TextBody> {
        self.text_body.as_deref()
    }

    fn paragraphs(&self) -> &[TextParagraph] {
        static EMPTY: &[TextParagraph] = &[];
        self.text_body
            .as_ref()
            .map(|tb| tb.paragraphs())
            .unwrap_or(EMPTY)
    }

    fn text(&self) -> Option<String> {
        self.text_body.as_ref().map(|tb| {
            tb.paragraphs()
                .iter()
                .map(|p| p.text())
                .collect::<Vec<_>>()
                .join("\n")
        })
    }

    fn has_text(&self) -> bool {
        self.text_body
            .as_ref()
            .is_some_and(|tb| !tb.paragraphs().is_empty())
    }

    fn shape_id(&self) -> u32 {
        self.non_visual_properties.c_nv_pr.id
    }

    fn is_placeholder(&self) -> bool {
        self.non_visual_properties.nv_pr.ph.is_some()
    }

    fn placeholder_type(&self) -> Option<&STPlaceholderType> {
        self.non_visual_properties
            .nv_pr
            .ph
            .as_deref()
            .and_then(|p| p.r#type.as_ref())
    }

    fn placeholder_index(&self) -> Option<u32> {
        self.non_visual_properties
            .nv_pr
            .ph
            .as_deref()
            .and_then(|p| p.idx)
    }

    fn offset_emu(&self) -> Option<(i64, i64)> {
        use ooxml_dml::ext::ShapePropertiesExt;
        self.shape_properties.offset_emu()
    }

    fn extent_emu(&self) -> Option<(i64, i64)> {
        use ooxml_dml::ext::ShapePropertiesExt;
        self.shape_properties.extent_emu()
    }

    fn rotation_angle_deg(&self) -> Option<f64> {
        use ooxml_dml::ext::ShapePropertiesExt;
        self.shape_properties.rotation_angle_deg()
    }
}

/// Extension trait for Picture (p:pic element).
pub trait PictureExt {
    /// Get the picture name from cNvPr.
    fn name(&self) -> &str;

    /// Get the picture description/alt text from cNvPr.
    fn description(&self) -> Option<&str>;

    /// Get the relationship ID for the embedded image (from blipFill/blip@r:embed).
    fn embed_rel_id(&self) -> Option<&str>;

    /// Get the position offset in EMU as (x, y).
    fn offset_emu(&self) -> Option<(i64, i64)>;

    /// Get the extent (width, height) in EMU.
    fn extent_emu(&self) -> Option<(i64, i64)>;

    /// Get the crop rectangle, if any.
    ///
    /// Specifies what portion of the image is displayed.
    /// Requires the `dml-fills` feature.
    #[cfg(feature = "dml-fills")]
    fn crop_rect(&self) -> Option<&ooxml_dml::types::CTRelativeRect>;
}

impl PictureExt for Picture {
    fn name(&self) -> &str {
        &self.non_visual_picture_properties.c_nv_pr.name
    }

    fn description(&self) -> Option<&str> {
        self.non_visual_picture_properties.c_nv_pr.descr.as_deref()
    }

    fn embed_rel_id(&self) -> Option<&str> {
        self.blip_fill
            .blip
            .as_ref()
            .and_then(|b| b.embed.as_deref())
    }

    fn offset_emu(&self) -> Option<(i64, i64)> {
        use ooxml_dml::ext::ShapePropertiesExt;
        self.shape_properties.offset_emu()
    }

    fn extent_emu(&self) -> Option<(i64, i64)> {
        use ooxml_dml::ext::ShapePropertiesExt;
        self.shape_properties.extent_emu()
    }

    #[cfg(feature = "dml-fills")]
    fn crop_rect(&self) -> Option<&ooxml_dml::types::CTRelativeRect> {
        self.blip_fill.src_rect.as_deref()
    }
}

/// Extension trait for Connector (p:cxnSp element).
pub trait ConnectorExt {
    /// Get the connector name from cNvPr.
    fn name(&self) -> &str;

    /// Get the connector description/alt text from cNvPr.
    fn description(&self) -> Option<&str>;
}

impl ConnectorExt for Connector {
    fn name(&self) -> &str {
        &self.non_visual_connector_properties.c_nv_pr.name
    }

    fn description(&self) -> Option<&str> {
        self.non_visual_connector_properties
            .c_nv_pr
            .descr
            .as_deref()
    }
}

/// Extension trait for GraphicalObjectFrame (p:graphicFrame element).
pub trait GraphicalObjectFrameExt {
    /// Get the frame name from cNvPr.
    fn name(&self) -> &str;

    /// Get the frame description/alt text from cNvPr.
    fn description(&self) -> Option<&str>;
}

impl GraphicalObjectFrameExt for GraphicalObjectFrame {
    fn name(&self) -> &str {
        &self.nv_graphic_frame_pr.c_nv_pr.name
    }

    fn description(&self) -> Option<&str> {
        self.nv_graphic_frame_pr.c_nv_pr.descr.as_deref()
    }
}

/// Extension trait for GroupShape (p:grpSp and p:spTree elements).
///
/// This trait works for both group shapes and the root shape tree,
/// since p:spTree is defined as a GroupShape in the schema.
pub trait GroupShapeExt {
    /// Get the group name from cNvPr.
    fn name(&self) -> &str;

    /// Get the group description/alt text from cNvPr.
    fn description(&self) -> Option<&str>;

    /// Get all shapes in this group.
    fn shapes(&self) -> &[Shape];

    /// Get all pictures in this group.
    fn pictures(&self) -> &[Picture];

    /// Get all connectors in this group.
    fn connectors(&self) -> &[Connector];

    /// Get all nested group shapes.
    fn group_shapes(&self) -> &[GroupShape];

    /// Get all graphical object frames (charts, tables, etc.).
    fn graphic_frames(&self) -> &[GraphicalObjectFrame];

    /// Get all text from shapes in this group (recursively).
    fn text(&self) -> String;

    /// Collect all shapes recursively (including from nested group shapes).
    fn all_shapes_recursive(&self) -> Vec<&Shape>;

    /// Collect all text from all shapes recursively, joined by newlines.
    fn all_text_recursive(&self) -> String;
}

impl GroupShapeExt for GroupShape {
    fn name(&self) -> &str {
        &self.non_visual_group_properties.c_nv_pr.name
    }

    fn description(&self) -> Option<&str> {
        self.non_visual_group_properties.c_nv_pr.descr.as_deref()
    }

    fn shapes(&self) -> &[Shape] {
        &self.shape
    }

    fn pictures(&self) -> &[Picture] {
        &self.picture
    }

    fn connectors(&self) -> &[Connector] {
        &self.connector
    }

    fn group_shapes(&self) -> &[GroupShape] {
        &self.group_shape
    }

    fn graphic_frames(&self) -> &[GraphicalObjectFrame] {
        &self.graphic_frame
    }

    fn text(&self) -> String {
        let mut texts = Vec::new();

        // Collect text from shapes
        for shape in &self.shape {
            if let Some(t) = shape.text() {
                texts.push(t);
            }
        }

        // Recursively collect from nested groups
        for group in &self.group_shape {
            let t = group.text();
            if !t.is_empty() {
                texts.push(t);
            }
        }

        texts.join("\n")
    }

    fn all_shapes_recursive(&self) -> Vec<&Shape> {
        let mut shapes: Vec<&Shape> = self.shape.iter().collect();
        for group in &self.group_shape {
            shapes.extend(group.all_shapes_recursive());
        }
        shapes
    }

    fn all_text_recursive(&self) -> String {
        self.all_shapes_recursive()
            .iter()
            .filter_map(|s| s.text())
            .filter(|t| !t.is_empty())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// Extension trait for CommonSlideData (p:cSld element).
pub trait CommonSlideDataExt {
    /// Get the shape tree (which is a GroupShape).
    fn shape_tree(&self) -> &GroupShape;

    /// Get all shapes from the shape tree.
    fn shapes(&self) -> &[Shape];

    /// Get all pictures from the shape tree.
    fn pictures(&self) -> &[Picture];

    /// Get all text content from all shapes.
    fn text(&self) -> String;
}

impl CommonSlideDataExt for CommonSlideData {
    fn shape_tree(&self) -> &GroupShape {
        &self.shape_tree
    }

    fn shapes(&self) -> &[Shape] {
        self.shape_tree.shapes()
    }

    fn pictures(&self) -> &[Picture] {
        self.shape_tree.pictures()
    }

    fn text(&self) -> String {
        self.shape_tree.text()
    }
}

/// Extension trait for Slide (p:sld element).
pub trait SlideExt {
    /// Get the common slide data.
    fn common_slide_data(&self) -> &CommonSlideData;

    /// Get the shape tree (which is a GroupShape).
    fn shape_tree(&self) -> &GroupShape;

    /// Get all shapes on the slide.
    fn shapes(&self) -> &[Shape];

    /// Get all pictures on the slide.
    fn pictures(&self) -> &[Picture];

    /// Get all text content from the slide.
    fn text(&self) -> String;

    /// Get the slide transition (if any).
    #[cfg(feature = "pml-transitions")]
    fn transition(&self) -> Option<&SlideTransition>;

    /// Get the slide background, if one is explicitly set.
    ///
    /// Requires the `pml-styling` feature.
    #[cfg(feature = "pml-styling")]
    fn background(&self) -> Option<&CTBackground>;

    /// Check if the slide is hidden (show=false).
    fn is_hidden(&self) -> bool;
}

impl SlideExt for Slide {
    fn common_slide_data(&self) -> &CommonSlideData {
        &self.common_slide_data
    }

    fn shape_tree(&self) -> &GroupShape {
        self.common_slide_data.shape_tree()
    }

    fn shapes(&self) -> &[Shape] {
        self.common_slide_data.shapes()
    }

    fn pictures(&self) -> &[Picture] {
        self.common_slide_data.pictures()
    }

    fn text(&self) -> String {
        self.common_slide_data.text()
    }

    #[cfg(feature = "pml-transitions")]
    fn transition(&self) -> Option<&SlideTransition> {
        self.transition.as_deref()
    }

    #[cfg(feature = "pml-styling")]
    fn background(&self) -> Option<&CTBackground> {
        self.common_slide_data.bg.as_deref()
    }

    fn is_hidden(&self) -> bool {
        !self.show.unwrap_or(true)
    }
}

/// Extension trait for SlideLayout (p:sldLayout element).
pub trait SlideLayoutExt {
    /// Get the common slide data.
    fn common_slide_data(&self) -> &CommonSlideData;

    /// Get the layout type (if specified).
    fn layout_type(&self) -> Option<&STSlideLayoutType>;

    /// Check if this layout should show master shapes.
    fn show_master_shapes(&self) -> bool;
}

impl SlideLayoutExt for SlideLayout {
    fn common_slide_data(&self) -> &CommonSlideData {
        &self.common_slide_data
    }

    #[cfg(feature = "pml-masters")]
    fn layout_type(&self) -> Option<&STSlideLayoutType> {
        self.r#type.as_ref()
    }

    #[cfg(not(feature = "pml-masters"))]
    fn layout_type(&self) -> Option<&STSlideLayoutType> {
        None
    }

    #[cfg(feature = "pml-masters")]
    fn show_master_shapes(&self) -> bool {
        self.show_master_sp.unwrap_or(true)
    }

    #[cfg(not(feature = "pml-masters"))]
    fn show_master_shapes(&self) -> bool {
        true
    }
}

/// Extension trait for SlideMaster (p:sldMaster element).
pub trait SlideMasterExt {
    /// Get the common slide data.
    fn common_slide_data(&self) -> &CommonSlideData;

    /// Check if this master should preserve content.
    fn preserve(&self) -> bool;
}

impl SlideMasterExt for SlideMaster {
    fn common_slide_data(&self) -> &CommonSlideData {
        &self.common_slide_data
    }

    #[cfg(feature = "pml-masters")]
    fn preserve(&self) -> bool {
        self.preserve.unwrap_or(false)
    }

    #[cfg(not(feature = "pml-masters"))]
    fn preserve(&self) -> bool {
        false
    }
}

/// Extension trait for NotesSlide (p:notes element).
#[cfg(feature = "pml-notes")]
pub trait NotesSlideExt {
    /// Get the common slide data containing the notes content.
    fn common_slide_data(&self) -> &CommonSlideData;

    /// Get all text from the notes.
    fn text(&self) -> String;
}

#[cfg(feature = "pml-notes")]
impl NotesSlideExt for NotesSlide {
    fn common_slide_data(&self) -> &CommonSlideData {
        &self.common_slide_data
    }

    fn text(&self) -> String {
        self.common_slide_data.text()
    }
}
