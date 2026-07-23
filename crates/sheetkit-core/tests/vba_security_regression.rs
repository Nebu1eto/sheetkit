use std::io::{Cursor, Write as _};

use sheetkit_core::error::Error;
use sheetkit_core::resource::ResourceDimension;
use sheetkit_core::vba::{extract_vba_modules, extract_vba_modules_with_limits, VbaLimits};

fn cfb_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let mut output = Vec::new();
    let cursor = Cursor::new(&mut output);
    let mut compound = cfb::CompoundFile::create(cursor).expect("create CFB");
    compound.create_storage("/VBA").expect("create storage");
    for (name, bytes) in streams {
        let mut stream = compound
            .create_stream(format!("/VBA/{name}"))
            .expect("create stream");
        stream.write_all(bytes).expect("write stream");
    }
    compound.flush().expect("flush CFB");
    output
}

fn uncompressed_vba(data: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(data.len() + 3);
    output.push(0x01);
    let header = u16::try_from(data.len().saturating_sub(1)).expect("small fixture");
    output.extend_from_slice(&header.to_le_bytes());
    output.extend_from_slice(data);
    output
}

fn write_record(output: &mut Vec<u8>, id: u16, data: &[u8]) {
    output.extend_from_slice(&id.to_le_bytes());
    output.extend_from_slice(
        &u32::try_from(data.len())
            .expect("small fixture")
            .to_le_bytes(),
    );
    output.extend_from_slice(data);
}

fn one_module_project() -> Vec<u8> {
    let mut dir = Vec::new();
    write_record(&mut dir, 0x0003, &1252_u16.to_le_bytes());
    write_record(&mut dir, 0x0019, b"Module1");
    write_record(&mut dir, 0x001A, b"Module1");
    write_record(&mut dir, 0x0031, &0_u32.to_le_bytes());
    write_record(&mut dir, 0x0021, &[]);
    write_record(&mut dir, 0x002B, &[]);

    let compressed_dir = uncompressed_vba(&dir);
    let compressed_source = uncompressed_vba(b"Sub Hello()\r\nEnd Sub\r\n");
    cfb_with_streams(&[
        ("dir", compressed_dir.as_slice()),
        ("Module1", compressed_source.as_slice()),
        ("_VBA_PROJECT", &[0xCC, 0x61]),
    ])
}

fn missing_module_project(module_count: usize) -> Vec<u8> {
    let mut dir = Vec::new();
    write_record(&mut dir, 0x0003, &1252_u16.to_le_bytes());
    for index in 0..module_count {
        let name = format!("attacker-module-{index}");
        write_record(&mut dir, 0x0019, name.as_bytes());
        write_record(&mut dir, 0x001A, name.as_bytes());
        write_record(&mut dir, 0x0031, &0_u32.to_le_bytes());
        write_record(&mut dir, 0x0021, &[]);
    }
    let compressed_dir = uncompressed_vba(&dir);
    cfb_with_streams(&[("dir", compressed_dir.as_slice())])
}

fn assert_dimension(error: Error, expected: ResourceDimension) {
    match error {
        Error::ResourceLimitExceeded(metadata) => assert_eq!(metadata.dimension(), expected),
        other => panic!("expected resource error, got {other}"),
    }
}

#[test]
fn input_limit_accepts_exact_size_and_rejects_plus_one() {
    let project = one_module_project();
    let mut limits = VbaLimits {
        input_bytes: u64::try_from(project.len()).expect("fixture size"),
        ..VbaLimits::default()
    };
    extract_vba_modules_with_limits(&project, limits).expect("exact input limit");

    limits.input_bytes -= 1;
    let error = extract_vba_modules_with_limits(&project, limits).expect_err("over input limit");
    assert_dimension(error, ResourceDimension::InputBytes);
}

#[test]
fn stream_count_and_declared_size_are_preflighted() {
    let project = one_module_project();

    let count_limits = VbaLimits {
        cfb_stream_count: 2,
        ..VbaLimits::default()
    };
    let error = extract_vba_modules_with_limits(&project, count_limits)
        .expect_err("third stream exceeds count");
    assert_dimension(error, ResourceDimension::StreamCount);

    let size_limits = VbaLimits {
        cfb_stream_declared_bytes: 1,
        ..VbaLimits::default()
    };
    let error = extract_vba_modules_with_limits(&project, size_limits)
        .expect_err("declared stream size exceeds limit");
    assert_dimension(error, ResourceDimension::EntryDeclaredBytes);
}

#[test]
fn aggregate_declared_and_read_limits_accept_exact_and_reject_plus_one() {
    let project = one_module_project();
    let compound = cfb::CompoundFile::open(Cursor::new(project.as_slice())).expect("open fixture");
    let aggregate_declared = compound
        .walk()
        .filter(|entry| entry.is_stream())
        .fold(0_u64, |total, entry| total + entry.len());
    let read_streams: Vec<u64> = compound
        .walk()
        .filter(|entry| matches!(entry.name(), "dir" | "Module1"))
        .map(|entry| entry.len())
        .collect();
    let aggregate_read: u64 = read_streams.iter().sum();
    let largest_read = read_streams.into_iter().max().expect("read streams");

    let mut limits = VbaLimits {
        cfb_aggregate_declared_bytes: aggregate_declared,
        cfb_stream_read_bytes: largest_read,
        cfb_aggregate_read_bytes: aggregate_read,
        compressed_bytes: aggregate_read,
        ..VbaLimits::default()
    };
    extract_vba_modules_with_limits(&project, limits).expect("exact CFB limits");

    limits.cfb_aggregate_declared_bytes = aggregate_declared - 1;
    let error = extract_vba_modules_with_limits(&project, limits)
        .expect_err("declared aggregate exceeds limit by one");
    assert_dimension(error, ResourceDimension::AggregateDeclaredBytes);

    limits.cfb_aggregate_declared_bytes = aggregate_declared;
    limits.cfb_aggregate_read_bytes = aggregate_read - 1;
    let error = extract_vba_modules_with_limits(&project, limits)
        .expect_err("read aggregate exceeds limit by one");
    assert_dimension(error, ResourceDimension::AggregateActualBytes);
}

#[test]
fn directory_and_modules_share_one_aggregate_decode_budget() {
    let project = one_module_project();
    let mut limits = VbaLimits {
        aggregate_decoded_bytes: 78,
        ..VbaLimits::default()
    };
    extract_vba_modules_with_limits(&project, limits).expect("calibrated aggregate limit");

    limits.aggregate_decoded_bytes = 77;
    let error = extract_vba_modules_with_limits(&project, limits)
        .expect_err("dir and module output share one budget");
    assert_dimension(error, ResourceDimension::OutputBytes);
}

#[test]
fn compressed_retained_and_output_budgets_are_mandatory() {
    let project = one_module_project();

    let mut limits = VbaLimits {
        compressed_bytes: 1,
        ..VbaLimits::default()
    };
    let error = extract_vba_modules_with_limits(&project, limits)
        .expect_err("compressed input exceeds limit");
    assert_dimension(error, ResourceDimension::InputBytes);

    limits = VbaLimits::default();
    limits.retained_bytes = 1;
    let error = extract_vba_modules_with_limits(&project, limits)
        .expect_err("retained output exceeds limit");
    assert_dimension(error, ResourceDimension::RetainedBytes);

    limits = VbaLimits::default();
    limits.output_bytes = 1;
    let error = extract_vba_modules_with_limits(&project, limits)
        .expect_err("returned output exceeds limit");
    assert_dimension(error, ResourceDimension::OutputBytes);
}

#[test]
fn invalid_limits_and_malformed_input_return_sanitized_errors() {
    let limits = VbaLimits {
        module_count: VbaLimits::hard_ceiling().module_count + 1,
        ..VbaLimits::default()
    };
    let error = extract_vba_modules_with_limits(b"attacker-controlled-source", limits)
        .expect_err("hard ceiling is mandatory");
    let message = error.to_string();
    assert!(!message.contains("attacker-controlled-source"));

    let error =
        extract_vba_modules(b"attacker-controlled-source").expect_err("malformed CFB must fail");
    let message = error.to_string();
    assert!(!message.contains("attacker-controlled-source"));

    let project = one_module_project();
    let truncated = &project[..project.len() / 2];
    let error = extract_vba_modules(truncated).expect_err("truncated CFB must fail");
    assert!(!error.to_string().contains("Module1"));
}

#[test]
fn warnings_are_bounded_and_do_not_echo_module_names() {
    let project = missing_module_project(40);
    let extracted = extract_vba_modules(&project).expect("missing streams are non-fatal");

    assert_eq!(extracted.warnings.len(), 33);
    assert_eq!(
        extracted.warnings.last().map(String::as_str),
        Some("additional VBA warnings omitted")
    );
    assert!(extracted
        .warnings
        .iter()
        .all(|warning| !warning.contains("attacker-module")));
}

#[test]
fn accepted_project_matches_default_explicit_limits() {
    let project = one_module_project();
    let default_project = extract_vba_modules(&project).expect("default extraction");
    let explicit_project = extract_vba_modules_with_limits(&project, VbaLimits::default())
        .expect("explicit extraction");

    assert_eq!(default_project.modules.len(), 1);
    assert_eq!(explicit_project.modules.len(), 1);
    assert_eq!(
        default_project.modules[0].name,
        explicit_project.modules[0].name
    );
    assert_eq!(
        default_project.modules[0].source_code,
        explicit_project.modules[0].source_code
    );
    assert_eq!(default_project.warnings, explicit_project.warnings);
}
