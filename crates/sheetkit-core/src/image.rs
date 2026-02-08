//! Image insertion and management.
//!
//! Provides types for configuring image placement in worksheets and
//! helpers for building the corresponding drawing XML structures.

use sheetkit_xml::drawing::{
    AExt, Blip, BlipFill, CNvPicPr, CNvPr, ClientData, Extent, FillRect, MarkerType, NvPicPr,
    Offset, OneCellAnchor, Picture, PrstGeom, SpPr, Stretch, WsDr, Xfrm,
};

use crate::error::{Error, Result};
use crate::utils::cell_ref::cell_name_to_coordinates;

/// EMU (English Metric Units) per pixel at 96 DPI.
/// 1 inch = 914400 EMU, 1 inch = 96 pixels => 1 pixel = 9525 EMU.
pub const EMU_PER_PIXEL: u64 = 9525;

/// Supported image formats.
#[derive(Debug, Clone, PartialEq)]
pub enum ImageFormat {
    /// PNG image.
    Png,
    /// JPEG image.
    Jpeg,
    /// GIF image.
    Gif,
}

impl ImageFormat {
    /// Return the MIME content type string for this image format.
    pub fn content_type(&self) -> &str {
        match self {
            ImageFormat::Png => "image/png",
            ImageFormat::Jpeg => "image/jpeg",
            ImageFormat::Gif => "image/gif",
        }
    }

    /// Return the file extension for this image format.
    pub fn extension(&self) -> &str {
        match self {
            ImageFormat::Png => "png",
            ImageFormat::Jpeg => "jpeg",
            ImageFormat::Gif => "gif",
        }
    }
}

/// Configuration for inserting an image into a worksheet.
#[derive(Debug, Clone)]
pub struct ImageConfig {
    /// Raw image bytes.
    pub data: Vec<u8>,
    /// Image format.
    pub format: ImageFormat,
    /// Anchor cell reference (e.g., `"B2"`).
    pub from_cell: String,
    /// Image width in pixels.
    pub width_px: u32,
    /// Image height in pixels.
    pub height_px: u32,
}

/// Convert pixel dimensions to EMU.
pub fn pixels_to_emu(px: u32) -> u64 {
    px as u64 * EMU_PER_PIXEL
}

/// Build a drawing XML structure containing a single image.
///
/// The image is anchored to a single cell (one-cell anchor) with explicit
/// width and height in EMU derived from pixel dimensions.
pub fn build_drawing_with_image(image_ref_id: &str, config: &ImageConfig) -> Result<WsDr> {
    let (col, row) = cell_name_to_coordinates(&config.from_cell)?;
    // MarkerType uses 0-based column and row indices
    let from = MarkerType {
        col: col - 1,
        col_off: 0,
        row: row - 1,
        row_off: 0,
    };

    let cx = pixels_to_emu(config.width_px);
    let cy = pixels_to_emu(config.height_px);

    let pic = Picture {
        nv_pic_pr: NvPicPr {
            c_nv_pr: CNvPr {
                id: 2,
                name: "Picture 1".to_string(),
            },
            c_nv_pic_pr: CNvPicPr {},
        },
        blip_fill: BlipFill {
            blip: Blip {
                r_embed: image_ref_id.to_string(),
            },
            stretch: Stretch {
                fill_rect: FillRect {},
            },
        },
        sp_pr: SpPr {
            xfrm: Xfrm {
                off: Offset { x: 0, y: 0 },
                ext: AExt { cx, cy },
            },
            prst_geom: PrstGeom {
                prst: "rect".to_string(),
            },
        },
    };

    let anchor = OneCellAnchor {
        from,
        ext: Extent { cx, cy },
        pic: Some(pic),
        client_data: ClientData {},
    };

    Ok(WsDr {
        one_cell_anchors: vec![anchor],
        ..WsDr::default()
    })
}

/// Add an image anchor to an existing drawing.
///
/// If a drawing already exists for a sheet (e.g., it already has a chart),
/// this function adds the image anchor to it.
pub fn add_image_to_drawing(
    drawing: &mut WsDr,
    image_ref_id: &str,
    config: &ImageConfig,
    pic_id: u32,
) -> Result<()> {
    let (col, row) = cell_name_to_coordinates(&config.from_cell)?;
    let from = MarkerType {
        col: col - 1,
        col_off: 0,
        row: row - 1,
        row_off: 0,
    };

    let cx = pixels_to_emu(config.width_px);
    let cy = pixels_to_emu(config.height_px);

    let pic = Picture {
        nv_pic_pr: NvPicPr {
            c_nv_pr: CNvPr {
                id: pic_id,
                name: format!("Picture {}", pic_id - 1),
            },
            c_nv_pic_pr: CNvPicPr {},
        },
        blip_fill: BlipFill {
            blip: Blip {
                r_embed: image_ref_id.to_string(),
            },
            stretch: Stretch {
                fill_rect: FillRect {},
            },
        },
        sp_pr: SpPr {
            xfrm: Xfrm {
                off: Offset { x: 0, y: 0 },
                ext: AExt { cx, cy },
            },
            prst_geom: PrstGeom {
                prst: "rect".to_string(),
            },
        },
    };

    drawing.one_cell_anchors.push(OneCellAnchor {
        from,
        ext: Extent { cx, cy },
        pic: Some(pic),
        client_data: ClientData {},
    });

    Ok(())
}

/// Validate an ImageConfig.
pub fn validate_image_config(config: &ImageConfig) -> Result<()> {
    if config.data.is_empty() {
        return Err(Error::Internal("image data is empty".to_string()));
    }
    if config.width_px == 0 || config.height_px == 0 {
        return Err(Error::Internal(
            "image dimensions must be non-zero".to_string(),
        ));
    }
    // Validate the cell reference
    cell_name_to_coordinates(&config.from_cell)?;
    Ok(())
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_emu_per_pixel_constant() {
        assert_eq!(EMU_PER_PIXEL, 9525);
    }

    #[test]
    fn test_pixels_to_emu() {
        assert_eq!(pixels_to_emu(1), 9525);
        assert_eq!(pixels_to_emu(100), 952500);
        assert_eq!(pixels_to_emu(1000), 9525000);
        assert_eq!(pixels_to_emu(0), 0);
    }

    #[test]
    fn test_image_format_content_type() {
        assert_eq!(ImageFormat::Png.content_type(), "image/png");
        assert_eq!(ImageFormat::Jpeg.content_type(), "image/jpeg");
        assert_eq!(ImageFormat::Gif.content_type(), "image/gif");
    }

    #[test]
    fn test_image_format_extension() {
        assert_eq!(ImageFormat::Png.extension(), "png");
        assert_eq!(ImageFormat::Jpeg.extension(), "jpeg");
        assert_eq!(ImageFormat::Gif.extension(), "gif");
    }

    #[test]
    fn test_build_drawing_with_image() {
        let config = ImageConfig {
            data: vec![0x89, 0x50, 0x4E, 0x47], // PNG header bytes
            format: ImageFormat::Png,
            from_cell: "B2".to_string(),
            width_px: 400,
            height_px: 300,
        };

        let dr = build_drawing_with_image("rId1", &config).unwrap();

        assert!(dr.two_cell_anchors.is_empty());
        assert_eq!(dr.one_cell_anchors.len(), 1);

        let anchor = &dr.one_cell_anchors[0];
        // B2 -> col=2, row=2 -> 0-based: col=1, row=1
        assert_eq!(anchor.from.col, 1);
        assert_eq!(anchor.from.row, 1);
        assert_eq!(anchor.ext.cx, 400 * 9525);
        assert_eq!(anchor.ext.cy, 300 * 9525);

        let pic = anchor.pic.as_ref().unwrap();
        assert_eq!(pic.blip_fill.blip.r_embed, "rId1");
        assert_eq!(pic.sp_pr.prst_geom.prst, "rect");
    }

    #[test]
    fn test_build_drawing_with_image_a1() {
        let config = ImageConfig {
            data: vec![0xFF, 0xD8], // JPEG header
            format: ImageFormat::Jpeg,
            from_cell: "A1".to_string(),
            width_px: 200,
            height_px: 100,
        };

        let dr = build_drawing_with_image("rId2", &config).unwrap();
        let anchor = &dr.one_cell_anchors[0];
        // A1 -> col=1, row=1 -> 0-based: col=0, row=0
        assert_eq!(anchor.from.col, 0);
        assert_eq!(anchor.from.row, 0);
    }

    #[test]
    fn test_build_drawing_with_image_invalid_cell() {
        let config = ImageConfig {
            data: vec![0x89],
            format: ImageFormat::Png,
            from_cell: "INVALID".to_string(),
            width_px: 100,
            height_px: 100,
        };

        let result = build_drawing_with_image("rId1", &config);
        assert!(result.is_err());
    }

    #[test]
    fn test_validate_image_config_ok() {
        let config = ImageConfig {
            data: vec![1, 2, 3],
            format: ImageFormat::Png,
            from_cell: "A1".to_string(),
            width_px: 100,
            height_px: 100,
        };
        assert!(validate_image_config(&config).is_ok());
    }

    #[test]
    fn test_validate_image_config_empty_data() {
        let config = ImageConfig {
            data: vec![],
            format: ImageFormat::Png,
            from_cell: "A1".to_string(),
            width_px: 100,
            height_px: 100,
        };
        assert!(validate_image_config(&config).is_err());
    }

    #[test]
    fn test_validate_image_config_zero_width() {
        let config = ImageConfig {
            data: vec![1],
            format: ImageFormat::Png,
            from_cell: "A1".to_string(),
            width_px: 0,
            height_px: 100,
        };
        assert!(validate_image_config(&config).is_err());
    }

    #[test]
    fn test_validate_image_config_zero_height() {
        let config = ImageConfig {
            data: vec![1],
            format: ImageFormat::Png,
            from_cell: "A1".to_string(),
            width_px: 100,
            height_px: 0,
        };
        assert!(validate_image_config(&config).is_err());
    }

    #[test]
    fn test_validate_image_config_invalid_cell() {
        let config = ImageConfig {
            data: vec![1],
            format: ImageFormat::Png,
            from_cell: "ZZZZZ0".to_string(),
            width_px: 100,
            height_px: 100,
        };
        assert!(validate_image_config(&config).is_err());
    }

    #[test]
    fn test_add_image_to_existing_drawing() {
        let mut dr = WsDr::default();

        let config = ImageConfig {
            data: vec![1, 2, 3],
            format: ImageFormat::Png,
            from_cell: "C5".to_string(),
            width_px: 200,
            height_px: 150,
        };

        add_image_to_drawing(&mut dr, "rId3", &config, 3).unwrap();

        assert_eq!(dr.one_cell_anchors.len(), 1);
        let anchor = &dr.one_cell_anchors[0];
        // C5 -> col=3, row=5 -> 0-based: col=2, row=4
        assert_eq!(anchor.from.col, 2);
        assert_eq!(anchor.from.row, 4);
        assert_eq!(
            anchor.pic.as_ref().unwrap().nv_pic_pr.c_nv_pr.name,
            "Picture 2"
        );
    }

    #[test]
    fn test_emu_calculation_accuracy() {
        // 1 inch = 96 pixels at 96 DPI, and 1 inch = 914400 EMU
        // So 96 pixels * 9525 EMU/pixel = 914400 EMU = 1 inch
        assert_eq!(pixels_to_emu(96), 914400);
    }
}
