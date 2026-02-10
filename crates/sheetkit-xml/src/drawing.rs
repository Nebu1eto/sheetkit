//! DrawingML Spreadsheet Drawing XML schema structures.
//!
//! Represents `xl/drawings/drawing{N}.xml` in the OOXML package.

use serde::{Deserialize, Serialize};

use crate::namespaces;

/// Root element for a spreadsheet drawing part.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "wsDr")]
pub struct WsDr {
    #[serde(rename = "@xmlns:xdr")]
    pub xmlns_xdr: String,

    #[serde(rename = "@xmlns:a")]
    pub xmlns_a: String,

    #[serde(rename = "@xmlns:r")]
    pub xmlns_r: String,

    #[serde(rename = "xdr:twoCellAnchor", default)]
    pub two_cell_anchors: Vec<TwoCellAnchor>,

    #[serde(rename = "xdr:oneCellAnchor", default)]
    pub one_cell_anchors: Vec<OneCellAnchor>,
}

impl Default for WsDr {
    fn default() -> Self {
        Self {
            xmlns_xdr: namespaces::DRAWING_ML_SPREADSHEET.to_string(),
            xmlns_a: namespaces::DRAWING_ML.to_string(),
            xmlns_r: namespaces::RELATIONSHIPS.to_string(),
            two_cell_anchors: vec![],
            one_cell_anchors: vec![],
        }
    }
}

/// An anchor defined by two cell markers (from/to).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TwoCellAnchor {
    #[serde(rename = "xdr:from")]
    pub from: MarkerType,

    #[serde(rename = "xdr:to")]
    pub to: MarkerType,

    #[serde(rename = "xdr:graphicFrame", skip_serializing_if = "Option::is_none")]
    pub graphic_frame: Option<GraphicFrame>,

    #[serde(rename = "xdr:pic", skip_serializing_if = "Option::is_none")]
    pub pic: Option<Picture>,

    #[serde(rename = "xdr:sp", skip_serializing_if = "Option::is_none")]
    pub shape: Option<Shape>,

    #[serde(rename = "xdr:clientData")]
    pub client_data: ClientData,
}

/// An anchor defined by one cell marker and an extent.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct OneCellAnchor {
    #[serde(rename = "xdr:from")]
    pub from: MarkerType,

    #[serde(rename = "xdr:ext")]
    pub ext: Extent,

    #[serde(rename = "xdr:pic", skip_serializing_if = "Option::is_none")]
    pub pic: Option<Picture>,

    #[serde(rename = "xdr:clientData")]
    pub client_data: ClientData,
}

/// A cell marker indicating column, column offset, row, and row offset.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct MarkerType {
    #[serde(rename = "xdr:col")]
    pub col: u32,

    #[serde(rename = "xdr:colOff")]
    pub col_off: u64,

    #[serde(rename = "xdr:row")]
    pub row: u32,

    #[serde(rename = "xdr:rowOff")]
    pub row_off: u64,
}

/// Extent (size) in EMU (English Metric Units).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Extent {
    #[serde(rename = "@cx")]
    pub cx: u64,

    #[serde(rename = "@cy")]
    pub cy: u64,
}

/// Graphic frame containing a chart reference.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct GraphicFrame {
    #[serde(rename = "xdr:nvGraphicFramePr")]
    pub nv_graphic_frame_pr: NvGraphicFramePr,

    #[serde(rename = "xdr:xfrm")]
    pub xfrm: Xfrm,

    #[serde(rename = "a:graphic")]
    pub graphic: Graphic,
}

/// Non-visual graphic frame properties.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct NvGraphicFramePr {
    #[serde(rename = "xdr:cNvPr")]
    pub c_nv_pr: CNvPr,

    #[serde(rename = "xdr:cNvGraphicFramePr")]
    pub c_nv_graphic_frame_pr: CNvGraphicFramePr,
}

/// Common non-visual properties (id and name).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CNvPr {
    #[serde(rename = "@id")]
    pub id: u32,

    #[serde(rename = "@name")]
    pub name: String,
}

/// Non-visual graphic frame properties (empty marker).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CNvGraphicFramePr {}

/// Transform (position and size) for a graphic frame.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Xfrm {
    #[serde(rename = "a:off")]
    pub off: Offset,

    #[serde(rename = "a:ext")]
    pub ext: AExt,
}

/// Offset position.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Offset {
    #[serde(rename = "@x")]
    pub x: i64,

    #[serde(rename = "@y")]
    pub y: i64,
}

/// DrawingML extent (width/height).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct AExt {
    #[serde(rename = "@cx")]
    pub cx: u64,

    #[serde(rename = "@cy")]
    pub cy: u64,
}

/// Graphic element containing chart data.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Graphic {
    #[serde(rename = "a:graphicData")]
    pub graphic_data: GraphicData,
}

/// Graphic data referencing a chart.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct GraphicData {
    #[serde(rename = "@uri")]
    pub uri: String,

    #[serde(rename = "c:chart")]
    pub chart: ChartRef,
}

/// Reference to a chart part via relationship ID.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ChartRef {
    #[serde(rename = "@xmlns:c")]
    pub xmlns_c: String,

    #[serde(rename = "@r:id")]
    pub r_id: String,
}

/// Picture element for images.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Picture {
    #[serde(rename = "xdr:nvPicPr")]
    pub nv_pic_pr: NvPicPr,

    #[serde(rename = "xdr:blipFill")]
    pub blip_fill: BlipFill,

    #[serde(rename = "xdr:spPr")]
    pub sp_pr: SpPr,
}

/// Non-visual picture properties.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct NvPicPr {
    #[serde(rename = "xdr:cNvPr")]
    pub c_nv_pr: CNvPr,

    #[serde(rename = "xdr:cNvPicPr")]
    pub c_nv_pic_pr: CNvPicPr,
}

/// Non-visual picture-specific properties (empty marker).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CNvPicPr {}

/// Blip fill referencing an embedded image.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BlipFill {
    #[serde(rename = "a:blip")]
    pub blip: Blip,

    #[serde(rename = "a:stretch")]
    pub stretch: Stretch,
}

/// Blip (Binary Large Image or Picture) reference.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Blip {
    #[serde(rename = "@r:embed")]
    pub r_embed: String,
}

/// Stretch fill mode.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Stretch {
    #[serde(rename = "a:fillRect")]
    pub fill_rect: FillRect,
}

/// Fill rectangle (empty element indicating full fill).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FillRect {}

/// Shape properties for a picture.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SpPr {
    #[serde(rename = "a:xfrm")]
    pub xfrm: Xfrm,

    #[serde(rename = "a:prstGeom")]
    pub prst_geom: PrstGeom,
}

/// Preset geometry (e.g., rectangle).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PrstGeom {
    #[serde(rename = "@prst")]
    pub prst: String,
}

/// Shape element (`<xdr:sp>`).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Shape {
    #[serde(rename = "xdr:nvSpPr")]
    pub nv_sp_pr: NvSpPr,

    #[serde(rename = "xdr:spPr")]
    pub sp_pr: ShapeSpPr,

    #[serde(rename = "xdr:txBody", skip_serializing_if = "Option::is_none")]
    pub tx_body: Option<TxBody>,
}

/// Non-visual shape properties.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct NvSpPr {
    #[serde(rename = "xdr:cNvPr")]
    pub c_nv_pr: CNvPr,

    #[serde(rename = "xdr:cNvSpPr")]
    pub c_nv_sp_pr: CNvSpPr,
}

/// Non-visual shape-specific properties (empty marker).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CNvSpPr {}

/// Shape properties with optional fill and line.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapeSpPr {
    #[serde(rename = "a:xfrm")]
    pub xfrm: Xfrm,

    #[serde(rename = "a:prstGeom")]
    pub prst_geom: PrstGeom,

    #[serde(rename = "a:solidFill", skip_serializing_if = "Option::is_none")]
    pub solid_fill: Option<SolidFill>,

    #[serde(rename = "a:ln", skip_serializing_if = "Option::is_none")]
    pub ln: Option<Ln>,
}

/// Solid fill with an sRGB color.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SolidFill {
    #[serde(rename = "a:srgbClr")]
    pub srgb_clr: SrgbClr,
}

/// sRGB color value.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SrgbClr {
    #[serde(rename = "@val")]
    pub val: String,
}

/// Line properties.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Ln {
    #[serde(rename = "@w", skip_serializing_if = "Option::is_none")]
    pub w: Option<u64>,

    #[serde(rename = "a:solidFill", skip_serializing_if = "Option::is_none")]
    pub solid_fill: Option<SolidFill>,
}

/// Text body element.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TxBody {
    #[serde(rename = "a:bodyPr")]
    pub body_pr: BodyPr,

    #[serde(rename = "a:lstStyle")]
    pub lst_style: LstStyle,

    #[serde(rename = "a:p")]
    pub paragraphs: Vec<Paragraph>,
}

/// Body properties for text (empty marker with optional attributes).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BodyPr {}

/// List style for text (empty marker).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct LstStyle {}

/// A text paragraph.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Paragraph {
    #[serde(rename = "a:r", default, skip_serializing_if = "Vec::is_empty")]
    pub runs: Vec<TextRun>,
}

/// A text run within a paragraph.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TextRun {
    #[serde(rename = "a:rPr", skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<RunProperties>,

    #[serde(rename = "a:t")]
    pub t: String,
}

/// Run-level text properties.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct RunProperties {
    #[serde(rename = "@lang", skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,

    #[serde(rename = "@sz", skip_serializing_if = "Option::is_none")]
    pub sz: Option<u32>,
}

/// Client data (empty element required by spec).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ClientData {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_ws_dr_default() {
        let dr = WsDr::default();
        assert_eq!(dr.xmlns_xdr, namespaces::DRAWING_ML_SPREADSHEET);
        assert_eq!(dr.xmlns_a, namespaces::DRAWING_ML);
        assert_eq!(dr.xmlns_r, namespaces::RELATIONSHIPS);
        assert!(dr.two_cell_anchors.is_empty());
        assert!(dr.one_cell_anchors.is_empty());
    }

    #[test]
    fn test_marker_type_serialize() {
        let marker = MarkerType {
            col: 1,
            col_off: 0,
            row: 2,
            row_off: 0,
        };
        let xml = quick_xml::se::to_string(&marker).unwrap();
        assert!(xml.contains("<xdr:col>1</xdr:col>"));
        assert!(xml.contains("<xdr:row>2</xdr:row>"));
    }

    #[test]
    fn test_extent_serialize() {
        let ext = Extent {
            cx: 9525000,
            cy: 4762500,
        };
        let xml = quick_xml::se::to_string(&ext).unwrap();
        assert!(xml.contains("cx=\"9525000\""));
        assert!(xml.contains("cy=\"4762500\""));
    }

    #[test]
    fn test_chart_ref_serialize() {
        let chart_ref = ChartRef {
            xmlns_c: namespaces::DRAWING_ML_CHART.to_string(),
            r_id: "rId1".to_string(),
        };
        let xml = quick_xml::se::to_string(&chart_ref).unwrap();
        assert!(xml.contains("r:id=\"rId1\""));
    }

    #[test]
    fn test_blip_serialize() {
        let blip = Blip {
            r_embed: "rId2".to_string(),
        };
        let xml = quick_xml::se::to_string(&blip).unwrap();
        assert!(xml.contains("r:embed=\"rId2\""));
    }

    #[test]
    fn test_prst_geom_serialize() {
        let geom = PrstGeom {
            prst: "rect".to_string(),
        };
        let xml = quick_xml::se::to_string(&geom).unwrap();
        assert!(xml.contains("prst=\"rect\""));
    }
}
