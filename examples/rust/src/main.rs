use sheetkit::{
    AlignmentStyle, AppProperties, BorderLineStyle, BorderSideStyle, BorderStyle, CellValue,
    ChartConfig, ChartSeries, ChartType, CommentConfig, CustomPropertyValue, DataValidationConfig,
    DocProperties, ErrorStyle, FillStyle, FontStyle, HorizontalAlign, ImageConfig, ImageFormat,
    PatternType, Style, StyleColor, ValidationType, VerticalAlign, Workbook,
    WorkbookProtectionConfig,
};

fn main() -> sheetkit::Result<()> {
    println!("=== SheetKit Rust 예제 ===\n");

    // ── Phase 1: 워크북 생성 및 저장 ──
    let mut wb = Workbook::new();
    println!(
        "[Phase 1] 새 워크북 생성 완료. 시트: {:?}",
        wb.sheet_names()
    );

    // ── Phase 2: 셀 값 읽기/쓰기 ──
    wb.set_cell_value("Sheet1", "A1", CellValue::String("이름".into()))?;
    wb.set_cell_value("Sheet1", "B1", CellValue::String("나이".into()))?;
    wb.set_cell_value("Sheet1", "C1", CellValue::String("활성".into()))?;
    wb.set_cell_value("Sheet1", "A2", CellValue::String("홍길동".into()))?;
    wb.set_cell_value("Sheet1", "B2", CellValue::Number(30.0))?;
    wb.set_cell_value("Sheet1", "C2", CellValue::Bool(true))?;
    wb.set_cell_value("Sheet1", "A3", CellValue::String("김철수".into()))?;
    wb.set_cell_value("Sheet1", "B3", CellValue::Number(25.0))?;
    wb.set_cell_value("Sheet1", "C3", CellValue::Bool(false))?;

    let val = wb.get_cell_value("Sheet1", "A1")?;
    println!("[Phase 2] A1 셀 값: {:?}", val);

    // ── Phase 5: 시트 관리 ──
    let idx = wb.new_sheet("매출데이터")?;
    println!("[Phase 5] '매출데이터' 시트 추가 (인덱스: {})", idx);
    wb.set_sheet_name("매출데이터", "Sales")?;
    wb.copy_sheet("Sheet1", "Sheet1_복사")?;
    wb.set_active_sheet("Sheet1")?;
    println!("[Phase 5] 시트 목록: {:?}", wb.sheet_names());

    // ── Phase 3: 행/열 조작 ──
    wb.set_row_height("Sheet1", 1, 25.0)?;
    wb.set_col_width("Sheet1", "A", 20.0)?;
    wb.set_col_width("Sheet1", "B", 15.0)?;
    wb.set_col_width("Sheet1", "C", 12.0)?;
    wb.insert_rows("Sheet1", 1, 1)?; // 헤더 위에 제목 행 삽입
    wb.set_cell_value("Sheet1", "A1", CellValue::String("직원 목록".into()))?;
    println!("[Phase 3] 행/열 크기 조정 및 행 삽입 완료");

    // ── Phase 4: 스타일 ──
    // 제목 스타일
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

    // 헤더 스타일
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
    println!("[Phase 4] 스타일 적용 완료 (제목 + 헤더)");

    // ── Phase 7: 차트 ──
    // Sales 시트에 차트 데이터 작성
    wb.set_cell_value("Sales", "A1", CellValue::String("분기".into()))?;
    wb.set_cell_value("Sales", "B1", CellValue::String("매출".into()))?;
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
            title: Some("분기별 매출".into()),
            series: vec![ChartSeries {
                name: "매출".into(),
                categories: "Sales!$A$2:$A$5".into(),
                values: "Sales!$B$2:$B$5".into(),
            }],
            show_legend: true,
        },
    )?;
    println!("[Phase 7] 차트 추가 완료 (Sales 시트)");

    // ── Phase 7: 이미지 (1x1 PNG placeholder) ──
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
    println!("[Phase 7] 이미지 추가 완료");

    // ── Phase 8: 데이터 유효성 검사 ──
    wb.add_data_validation(
        "Sales",
        &DataValidationConfig {
            sqref: "C2:C5".into(),
            validation_type: ValidationType::List,
            operator: None,
            formula1: Some("\"달성,미달성,진행중\"".into()),
            formula2: None,
            allow_blank: true,
            show_input_message: true,
            prompt_title: Some("상태 선택".into()),
            prompt_message: Some("드롭다운에서 상태를 선택하세요".into()),
            show_error_message: true,
            error_style: Some(ErrorStyle::Stop),
            error_title: Some("오류".into()),
            error_message: Some("목록에서 선택해주세요".into()),
        },
    )?;
    println!("[Phase 8] 데이터 유효성 검사 추가 완료");

    // ── Phase 8: 코멘트 ──
    wb.add_comment(
        "Sheet1",
        &CommentConfig {
            cell: "A1".into(),
            author: "관리자".into(),
            text: "이 시트는 직원 목록을 포함합니다.".into(),
        },
    )?;
    println!("[Phase 8] 코멘트 추가 완료");

    // ── Phase 8: 자동 필터 ──
    wb.set_auto_filter("Sheet1", "A2:C4")?;
    println!("[Phase 8] 자동 필터 설정 완료");

    // ── Phase 9: StreamWriter ──
    let mut sw = wb.new_stream_writer("대용량시트")?;
    sw.set_col_width(1, 15.0)?;
    sw.set_col_width(2, 10.0)?;
    sw.write_row(
        1,
        &[
            CellValue::String("항목".into()),
            CellValue::String("값".into()),
        ],
    )?;
    for i in 2..=100 {
        sw.write_row(
            i,
            &[
                CellValue::String(format!("항목_{}", i - 1)),
                CellValue::Number(i as f64 * 10.0),
            ],
        )?;
    }
    sw.add_merge_cell("A1:B1")?;
    wb.apply_stream_writer(sw)?;
    println!("[Phase 9] StreamWriter로 100행 작성 완료");

    // ── Phase 10: 문서 속성 ──
    wb.set_doc_props(DocProperties {
        title: Some("SheetKit 예제 문서".into()),
        creator: Some("SheetKit Rust Example".into()),
        description: Some("SheetKit의 모든 기능을 보여주는 예제 파일".into()),
        ..Default::default()
    });
    wb.set_app_props(AppProperties {
        application: Some("SheetKit".into()),
        company: Some("SheetKit Project".into()),
        ..Default::default()
    });
    wb.set_custom_property("프로젝트", CustomPropertyValue::String("SheetKit".into()));
    wb.set_custom_property("버전", CustomPropertyValue::Int(1));
    wb.set_custom_property("릴리즈", CustomPropertyValue::Bool(false));
    println!("[Phase 10] 문서 속성 설정 완료");

    // ── Phase 10: 워크북 보호 ──
    wb.protect_workbook(WorkbookProtectionConfig {
        password: Some("demo".into()),
        lock_structure: true,
        lock_windows: false,
        lock_revision: false,
    });
    println!("[Phase 10] 워크북 보호 설정 완료");

    // ── 최종 저장 ──
    wb.save("output.xlsx")?;
    println!("\noutput.xlsx 파일이 생성되었습니다!");
    println!("시트 목록: {:?}", wb.sheet_names());

    Ok(())
}
