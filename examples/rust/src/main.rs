use sheetkit::{
    AlignmentStyle, AppProperties, BorderLineStyle, BorderSideStyle, BorderStyle, CellValue,
    ChartConfig, ChartSeries, ChartType, CommentConfig, CustomPropertyValue, DataValidationConfig,
    DocProperties, ErrorStyle, FillStyle, FontStyle, HorizontalAlign, ImageConfig, ImageFormat,
    PatternType, Style, StyleColor, ValidationType, VerticalAlign, Workbook,
    WorkbookProtectionConfig,
};

fn main() -> sheetkit::Result<()> {
    println!("=== SheetKit Rust Example ===\n");

    // ── Phase 1: Create and save workbook ──
    let mut wb = Workbook::new();
    println!(
        "[Phase 1] New workbook created. Sheets: {:?}",
        wb.sheet_names()
    );

    // ── Phase 2: Read/write cell values ──
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Name".into()))?;
    wb.set_cell_value("Sheet1", "B1", CellValue::String("Age".into()))?;
    wb.set_cell_value("Sheet1", "C1", CellValue::String("Active".into()))?;
    wb.set_cell_value("Sheet1", "A2", CellValue::String("John Doe".into()))?;
    wb.set_cell_value("Sheet1", "B2", CellValue::Number(30.0))?;
    wb.set_cell_value("Sheet1", "C2", CellValue::Bool(true))?;
    wb.set_cell_value("Sheet1", "A3", CellValue::String("Jane Smith".into()))?;
    wb.set_cell_value("Sheet1", "B3", CellValue::Number(25.0))?;
    wb.set_cell_value("Sheet1", "C3", CellValue::Bool(false))?;

    let val = wb.get_cell_value("Sheet1", "A1")?;
    println!("[Phase 2] A1 cell value: {:?}", val);

    // ── Phase 5: Sheet management ──
    let idx = wb.new_sheet("SalesData")?;
    println!("[Phase 5] 'SalesData' sheet added (index: {})", idx);
    wb.set_sheet_name("SalesData", "Sales")?;
    wb.copy_sheet("Sheet1", "Sheet1_Copy")?;
    wb.set_active_sheet("Sheet1")?;
    println!("[Phase 5] Sheet list: {:?}", wb.sheet_names());

    // ── Phase 3: Row/column operations ──
    wb.set_row_height("Sheet1", 1, 25.0)?;
    wb.set_col_width("Sheet1", "A", 20.0)?;
    wb.set_col_width("Sheet1", "B", 15.0)?;
    wb.set_col_width("Sheet1", "C", 12.0)?;
    wb.insert_rows("Sheet1", 1, 1)?; // Insert title row above header
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Employee List".into()))?;
    println!("[Phase 3] Row/column sizing and row insertion complete");

    // ── Phase 4: Styles ──
    // Title style
    let title_style = wb.add_style(&Style {
        font: Some(FontStyle {
            name: Some("Arial".into()),
            size: Some(16.0),
            bold: true,
            color: Some(StyleColor::Rgb("#FFFFFF".into())),
            ..Default::default()
        }),
        fill: Some(FillStyle {
            pattern: PatternType::Solid,
            fg_color: Some(StyleColor::Rgb("#4472C4".into())),
            bg_color: None,
        }),
        alignment: Some(AlignmentStyle {
            horizontal: Some(HorizontalAlign::Center),
            vertical: Some(VerticalAlign::Center),
            ..Default::default()
        }),
        ..Default::default()
    })?;
    wb.set_cell_style("Sheet1", "A1", title_style)?;

    // Header style
    let header_style = wb.add_style(&Style {
        font: Some(FontStyle {
            bold: true,
            size: Some(11.0),
            color: Some(StyleColor::Rgb("#FFFFFF".into())),
            ..Default::default()
        }),
        fill: Some(FillStyle {
            pattern: PatternType::Solid,
            fg_color: Some(StyleColor::Rgb("#5B9BD5".into())),
            bg_color: None,
        }),
        border: Some(BorderStyle {
            bottom: Some(BorderSideStyle {
                style: BorderLineStyle::Thin,
                color: Some(StyleColor::Rgb("#000000".into())),
            }),
            ..Default::default()
        }),
        alignment: Some(AlignmentStyle {
            horizontal: Some(HorizontalAlign::Center),
            ..Default::default()
        }),
        ..Default::default()
    })?;
    // Row 2 is now the header (after inserting row at 1)
    wb.set_cell_style("Sheet1", "A2", header_style)?;
    wb.set_cell_style("Sheet1", "B2", header_style)?;
    wb.set_cell_style("Sheet1", "C2", header_style)?;
    println!("[Phase 4] Styles applied (title + header)");

    // ── Phase 7: Chart ──
    // Write chart data on Sales sheet
    wb.set_cell_value("Sales", "A1", CellValue::String("Quarter".into()))?;
    wb.set_cell_value("Sales", "B1", CellValue::String("Revenue".into()))?;
    wb.set_cell_value("Sales", "A2", CellValue::String("Q1".into()))?;
    wb.set_cell_value("Sales", "B2", CellValue::Number(1500.0))?;
    wb.set_cell_value("Sales", "A3", CellValue::String("Q2".into()))?;
    wb.set_cell_value("Sales", "B3", CellValue::Number(2300.0))?;
    wb.set_cell_value("Sales", "A4", CellValue::String("Q3".into()))?;
    wb.set_cell_value("Sales", "B4", CellValue::Number(1800.0))?;
    wb.set_cell_value("Sales", "A5", CellValue::String("Q4".into()))?;
    wb.set_cell_value("Sales", "B5", CellValue::Number(2700.0))?;

    wb.add_chart(
        "Sales",
        "D1",
        "K15",
        &ChartConfig {
            chart_type: ChartType::Col,
            title: Some("Quarterly Revenue".into()),
            series: vec![ChartSeries {
                name: "Revenue".into(),
                categories: "Sales!$A$2:$A$5".into(),
                values: "Sales!$B$2:$B$5".into(),
            }],
            show_legend: true,
        },
    )?;
    println!("[Phase 7] Chart added (Sales sheet)");

    // ── Phase 7: Image (1x1 PNG placeholder) ──
    let png_1x1: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00,
        0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];
    wb.add_image(
        "Sheet1",
        &ImageConfig {
            data: png_1x1,
            format: ImageFormat::Png,
            from_cell: "E1".into(),
            width_px: 64,
            height_px: 64,
        },
    )?;
    println!("[Phase 7] Image added");

    // ── Phase 8: Data validation ──
    wb.add_data_validation(
        "Sales",
        &DataValidationConfig {
            sqref: "C2:C5".into(),
            validation_type: ValidationType::List,
            operator: None,
            formula1: Some("\"Achieved,Not Achieved,In Progress\"".into()),
            formula2: None,
            allow_blank: true,
            show_input_message: true,
            prompt_title: Some("Select Status".into()),
            prompt_message: Some("Select a status from the dropdown".into()),
            show_error_message: true,
            error_style: Some(ErrorStyle::Stop),
            error_title: Some("Error".into()),
            error_message: Some("Please select from the list".into()),
        },
    )?;
    println!("[Phase 8] Data validation added");

    // ── Phase 8: Comment ──
    wb.add_comment(
        "Sheet1",
        &CommentConfig {
            cell: "A1".into(),
            author: "Admin".into(),
            text: "This sheet contains the employee list.".into(),
        },
    )?;
    println!("[Phase 8] Comment added");

    // ── Phase 8: Auto filter ──
    wb.set_auto_filter("Sheet1", "A2:C4")?;
    println!("[Phase 8] Auto filter set");

    // ── Phase 9: StreamWriter ──
    let mut sw = wb.new_stream_writer("LargeSheet")?;
    sw.set_col_width(1, 15.0)?;
    sw.set_col_width(2, 10.0)?;
    sw.write_row(
        1,
        &[
            CellValue::String("Item".into()),
            CellValue::String("Value".into()),
        ],
    )?;
    for i in 2..=100 {
        sw.write_row(
            i,
            &[
                CellValue::String(format!("Item_{}", i - 1)),
                CellValue::Number(i as f64 * 10.0),
            ],
        )?;
    }
    sw.add_merge_cell("A1:B1")?;
    wb.apply_stream_writer(sw)?;
    println!("[Phase 9] StreamWriter wrote 100 rows");

    // ── Phase 10: Document properties ──
    wb.set_doc_props(DocProperties {
        title: Some("SheetKit Example Document".into()),
        creator: Some("SheetKit Rust Example".into()),
        description: Some("An example file demonstrating all SheetKit features".into()),
        ..Default::default()
    });
    wb.set_app_props(AppProperties {
        application: Some("SheetKit".into()),
        company: Some("SheetKit Project".into()),
        ..Default::default()
    });
    wb.set_custom_property("Project", CustomPropertyValue::String("SheetKit".into()));
    wb.set_custom_property("Version", CustomPropertyValue::Int(1));
    wb.set_custom_property("Release", CustomPropertyValue::Bool(false));
    println!("[Phase 10] Document properties set");

    // ── Phase 10: Workbook protection ──
    wb.protect_workbook(WorkbookProtectionConfig {
        password: Some("demo".into()),
        lock_structure: true,
        lock_windows: false,
        lock_revision: false,
    });
    println!("[Phase 10] Workbook protection set");

    // ── Save ──
    wb.save("output.xlsx")?;
    println!("\noutput.xlsx has been created!");
    println!("Sheet list: {:?}", wb.sheet_names());

    Ok(())
}
