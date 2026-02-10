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
    /// BMP image.
    Bmp,
    /// ICO image.
    Ico,
    /// TIFF image.
    Tiff,
    /// SVG image.
    Svg,
    /// EMF (Enhanced Metafile) image.
    Emf,
    /// EMZ (compressed EMF) image.
    Emz,
    /// WMF (Windows Metafile) image.
    Wmf,
    /// WMZ (compressed WMF) image.
    Wmz,
}

impl ImageFormat {
    /// Parse an extension string into an `ImageFormat`.
    ///
    /// Accepts common aliases such as `"jpg"` for JPEG and `"tif"` for TIFF.
    /// Returns `Error::UnsupportedImageFormat` for unrecognised strings.
    pub fn from_extension(ext: &str) -> Result<Self> {
        match ext.to_ascii_lowercase().as_str() {
            "png" => Ok(ImageFormat::Png),
            "jpeg" | "jpg" => Ok(ImageFormat::Jpeg),
            "gif" => Ok(ImageFormat::Gif),
            "bmp" => Ok(ImageFormat::Bmp),
            "ico" => Ok(ImageFormat::Ico),
            "tiff" | "tif" => Ok(ImageFormat::Tiff),
            "svg" => Ok(ImageFormat::Svg),
            "emf" => Ok(ImageFormat::Emf),
            "emz" => Ok(ImageFormat::Emz),
            "wmf" => Ok(ImageFormat::Wmf),
            "wmz" => Ok(ImageFormat::Wmz),
            _ => Err(Error::UnsupportedImageFormat {
                format: ext.to_string(),
            }),
        }
    }

    /// Return the MIME content type string for this image format.
    pub fn content_type(&self) -> &str {
        match self {
            ImageFormat::Png => "image/png",
            ImageFormat::Jpeg => "image/jpeg",
            ImageFormat::Gif => "image/gif",
            ImageFormat::Bmp => "image/bmp",
            ImageFormat::Ico => "image/x-icon",
            ImageFormat::Tiff => "image/tiff",
            ImageFormat::Svg => "image/svg+xml",
            ImageFormat::Emf => "image/x-emf",
            ImageFormat::Emz => "image/x-emz",
            ImageFormat::Wmf => "image/x-wmf",
            ImageFormat::Wmz => "image/x-wmz",
        }
    }

    /// Return the file extension for this image format.
    pub fn extension(&self) -> &str {
        match self {
            ImageFormat::Png => "png",
            ImageFormat::Jpeg => "jpeg",
            ImageFormat::Gif => "gif",
            ImageFormat::Bmp => "bmp",
            ImageFormat::Ico => "ico",
            ImageFormat::Tiff => "tiff",
            ImageFormat::Svg => "svg",
            ImageFormat::Emf => "emf",
            ImageFormat::Emz => "emz",
            ImageFormat::Wmf => "wmf",
            ImageFormat::Wmz => "wmz",
        }
    }
}

/// Information about a picture retrieved from a worksheet.
#[derive(Debug, Clone)]
pub struct PictureInfo {
    /// Raw image bytes.
    pub data: Vec<u8>,
    /// Image format.
    pub format: ImageFormat,
    /// Anchor cell reference (e.g., `"B2"`).
    pub cell: String,
    /// Image width in pixels.
    pub width_px: u32,
    /// Image height in pixels.
    pub height_px: u32,
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
    fn test_image_format_content_type_original() {
        assert_eq!(ImageFormat::Png.content_type(), "image/png");
        assert_eq!(ImageFormat::Jpeg.content_type(), "image/jpeg");
        assert_eq!(ImageFormat::Gif.content_type(), "image/gif");
    }

    #[test]
    fn test_image_format_content_type_new_formats() {
        assert_eq!(ImageFormat::Bmp.content_type(), "image/bmp");
        assert_eq!(ImageFormat::Ico.content_type(), "image/x-icon");
        assert_eq!(ImageFormat::Tiff.content_type(), "image/tiff");
        assert_eq!(ImageFormat::Svg.content_type(), "image/svg+xml");
        assert_eq!(ImageFormat::Emf.content_type(), "image/x-emf");
        assert_eq!(ImageFormat::Emz.content_type(), "image/x-emz");
        assert_eq!(ImageFormat::Wmf.content_type(), "image/x-wmf");
        assert_eq!(ImageFormat::Wmz.content_type(), "image/x-wmz");
    }

    #[test]
    fn test_image_format_extension_original() {
        assert_eq!(ImageFormat::Png.extension(), "png");
        assert_eq!(ImageFormat::Jpeg.extension(), "jpeg");
        assert_eq!(ImageFormat::Gif.extension(), "gif");
    }

    #[test]
    fn test_image_format_extension_new_formats() {
        assert_eq!(ImageFormat::Bmp.extension(), "bmp");
        assert_eq!(ImageFormat::Ico.extension(), "ico");
        assert_eq!(ImageFormat::Tiff.extension(), "tiff");
        assert_eq!(ImageFormat::Svg.extension(), "svg");
        assert_eq!(ImageFormat::Emf.extension(), "emf");
        assert_eq!(ImageFormat::Emz.extension(), "emz");
        assert_eq!(ImageFormat::Wmf.extension(), "wmf");
        assert_eq!(ImageFormat::Wmz.extension(), "wmz");
    }

    #[test]
    fn test_from_extension_original_formats() {
        assert_eq!(
            ImageFormat::from_extension("png").unwrap(),
            ImageFormat::Png
        );
        assert_eq!(
            ImageFormat::from_extension("jpeg").unwrap(),
            ImageFormat::Jpeg
        );
        assert_eq!(
            ImageFormat::from_extension("jpg").unwrap(),
            ImageFormat::Jpeg
        );
        assert_eq!(
            ImageFormat::from_extension("gif").unwrap(),
            ImageFormat::Gif
        );
    }

    #[test]
    fn test_from_extension_new_formats() {
        assert_eq!(
            ImageFormat::from_extension("bmp").unwrap(),
            ImageFormat::Bmp
        );
        assert_eq!(
            ImageFormat::from_extension("ico").unwrap(),
            ImageFormat::Ico
        );
        assert_eq!(
            ImageFormat::from_extension("tiff").unwrap(),
            ImageFormat::Tiff
        );
        assert_eq!(
            ImageFormat::from_extension("tif").unwrap(),
            ImageFormat::Tiff
        );
        assert_eq!(
            ImageFormat::from_extension("svg").unwrap(),
            ImageFormat::Svg
        );
        assert_eq!(
            ImageFormat::from_extension("emf").unwrap(),
            ImageFormat::Emf
        );
        assert_eq!(
            ImageFormat::from_extension("emz").unwrap(),
            ImageFormat::Emz
        );
        assert_eq!(
            ImageFormat::from_extension("wmf").unwrap(),
            ImageFormat::Wmf
        );
        assert_eq!(
            ImageFormat::from_extension("wmz").unwrap(),
            ImageFormat::Wmz
        );
    }

    #[test]
    fn test_from_extension_case_insensitive() {
        assert_eq!(
            ImageFormat::from_extension("PNG").unwrap(),
            ImageFormat::Png
        );
        assert_eq!(
            ImageFormat::from_extension("Jpeg").unwrap(),
            ImageFormat::Jpeg
        );
        assert_eq!(
            ImageFormat::from_extension("TIFF").unwrap(),
            ImageFormat::Tiff
        );
        assert_eq!(
            ImageFormat::from_extension("SVG").unwrap(),
            ImageFormat::Svg
        );
        assert_eq!(
            ImageFormat::from_extension("Emf").unwrap(),
            ImageFormat::Emf
        );
    }

    #[test]
    fn test_from_extension_unknown_returns_error() {
        let result = ImageFormat::from_extension("webp");
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(matches!(err, Error::UnsupportedImageFormat { .. }));
        assert!(err.to_string().contains("webp"));
    }

    #[test]
    fn test_from_extension_empty_returns_error() {
        let result = ImageFormat::from_extension("");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::UnsupportedImageFormat { .. }
        ));
    }

    #[test]
    fn test_from_extension_roundtrip() {
        let formats = [
            ImageFormat::Png,
            ImageFormat::Jpeg,
            ImageFormat::Gif,
            ImageFormat::Bmp,
            ImageFormat::Ico,
            ImageFormat::Tiff,
            ImageFormat::Svg,
            ImageFormat::Emf,
            ImageFormat::Emz,
            ImageFormat::Wmf,
            ImageFormat::Wmz,
        ];
        for fmt in &formats {
            let ext = fmt.extension();
            let parsed = ImageFormat::from_extension(ext).unwrap();
            assert_eq!(&parsed, fmt);
        }
    }

    #[test]
    fn test_build_drawing_with_image() {
        let config = ImageConfig {
            data: vec![0x89, 0x50, 0x4E, 0x47],
            format: ImageFormat::Png,
            from_cell: "B2".to_string(),
            width_px: 400,
            height_px: 300,
        };

        let dr = build_drawing_with_image("rId1", &config).unwrap();

        assert!(dr.two_cell_anchors.is_empty());
        assert_eq!(dr.one_cell_anchors.len(), 1);

        let anchor = &dr.one_cell_anchors[0];
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
            data: vec![0xFF, 0xD8],
            format: ImageFormat::Jpeg,
            from_cell: "A1".to_string(),
            width_px: 200,
            height_px: 100,
        };

        let dr = build_drawing_with_image("rId2", &config).unwrap();
        let anchor = &dr.one_cell_anchors[0];
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
    fn test_build_drawing_with_new_format() {
        let config = ImageConfig {
            data: vec![0x42, 0x4D],
            format: ImageFormat::Bmp,
            from_cell: "D4".to_string(),
            width_px: 320,
            height_px: 240,
        };

        let dr = build_drawing_with_image("rId1", &config).unwrap();
        assert_eq!(dr.one_cell_anchors.len(), 1);
        let anchor = &dr.one_cell_anchors[0];
        assert_eq!(anchor.from.col, 3);
        assert_eq!(anchor.from.row, 3);
        assert_eq!(anchor.ext.cx, 320 * 9525);
        assert_eq!(anchor.ext.cy, 240 * 9525);
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
    fn test_validate_image_config_new_format_ok() {
        let config = ImageConfig {
            data: vec![1, 2, 3],
            format: ImageFormat::Svg,
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
        assert_eq!(anchor.from.col, 2);
        assert_eq!(anchor.from.row, 4);
        assert_eq!(
            anchor.pic.as_ref().unwrap().nv_pic_pr.c_nv_pr.name,
            "Picture 2"
        );
    }

    #[test]
    fn test_emu_calculation_accuracy() {
        assert_eq!(pixels_to_emu(96), 914400);
    }
}
