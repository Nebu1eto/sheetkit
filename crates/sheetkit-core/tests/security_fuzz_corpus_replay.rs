#[path = "../../../security/fuzz_support.rs"]
mod fuzz_support;

use std::fs;
use std::path::{Path, PathBuf};

use fuzz_support::{
    exercise_formula_input, exercise_transfer_input, exercise_xml_input, exercise_zip_input,
    materialize_seed, Outcome,
};
use sheetkit_core::cell::CellValue;
use sheetkit_core::raw_transfer_write::{cells_to_raw_buffer, raw_buffer_to_cells};
use sheetkit_core::workbook::Workbook;

#[test]
fn canonical_security_corpus_replays_deterministically_without_panics() {
    for (target, path) in corpus_files() {
        let recipe = fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()));
        let input = materialize_seed(&recipe);
        let first = exercise(target, &input);
        let second = exercise(target, &input);
        assert_eq!(
            first,
            second,
            "non-deterministic replay for {}",
            path.display()
        );
        assert!(
            !first.contains(&Outcome::HarnessLimit),
            "corpus recipe exceeded its target harness limit: {}",
            path.display()
        );
        assert_canonical_outcome(target, &path, &first);
    }
}

#[test]
fn deterministic_structured_mutations_are_panic_free() {
    let zip = Workbook::new().save_to_buffer().unwrap();
    let xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
    let formula = b"SUM(A1:A10)+IF(B1>0,1,0)";
    let transfer =
        cells_to_raw_buffer(&[(1, vec![(1, CellValue::String("value".into()))])]).unwrap();

    let mut state = 0x6a09_e667_f3bc_c909_u64;
    for index in 0..2_048 {
        state = state
            .wrapping_mul(6_364_136_223_846_793_005)
            .wrapping_add(1_442_695_040_888_963_407);
        let target = index % 4;
        let seed = match target {
            0 => zip.as_slice(),
            1 => xml.as_slice(),
            2 => formula.as_slice(),
            _ => transfer.as_slice(),
        };
        let mutated = structured_mutation(seed, state, index);
        let _ = exercise(target, &mutated);
    }
}

#[test]
fn transfer_extremes_reject_and_exact_boundaries_round_trip() {
    for recipe in [
        b"recipe:transfer:v1-max-cell-count".as_slice(),
        b"recipe:transfer:v2-max-cell-count".as_slice(),
        b"hex:44524b530100ffffffff010000000000".as_slice(),
        b"hex:32524b530200ffffffff010000000000".as_slice(),
        b"hex:32524b530200010000000100000000000100000000000000010000ffffffff".as_slice(),
    ] {
        assert_eq!(
            exercise_transfer_input(&materialize_seed(recipe)),
            Outcome::Rejected
        );
    }

    let rows = [(1_048_576, vec![(16_384, CellValue::Number(1.0))])];
    let encoded = cells_to_raw_buffer(&rows).expect("exact worksheet boundary must encode");
    assert_eq!(raw_buffer_to_cells(&encoded).unwrap(), rows);

    for rows in [
        vec![(0, vec![(1, CellValue::Number(1.0))])],
        vec![(1_048_577, vec![(1, CellValue::Number(1.0))])],
        vec![(1, vec![(0, CellValue::Number(1.0))])],
        vec![(1, vec![(16_385, CellValue::Number(1.0))])],
        vec![(1, vec![(u32::MAX, CellValue::Number(1.0))])],
    ] {
        assert!(
            cells_to_raw_buffer(&rows).is_err(),
            "invalid coordinates encoded: {rows:?}"
        );
    }
}

fn corpus_files() -> Vec<(usize, PathBuf)> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join("security/fuzz-corpus");
    let mut files = Vec::new();
    for (target, directory, expected_names) in [
        (
            0,
            "workbook_open",
            &[
                "entry-count-exact.seed",
                "entry-count-over.seed",
                "minimal-workbook.seed",
                "misleading-declared-size.seed",
                "zip64-sector-exact.seed",
                "zip64-sector-over.seed",
            ][..],
        ),
        (
            1,
            "worksheet_xml",
            &[
                "canonical-max.seed",
                "canonical-min.seed",
                "column-over.seed",
                "direct-canonical-max.seed",
                "direct-canonical-min.seed",
                "direct-over.seed",
                "noncanonical-alias.seed",
                "oversized-digits.seed",
                "row-over.seed",
            ],
        ),
        (
            2,
            "formula",
            &[
                "binary-chain-exact.seed",
                "binary-chain-over.seed",
                "depth-exact.seed",
                "depth-over.seed",
                "normal.seed",
                "wildcard-near-limit.seed",
            ],
        ),
        (
            3,
            "transfer",
            &[
                "v1-empty.seed",
                "v1-max-cell-count.seed",
                "v1-max-row-count-tiny.seed",
                "v1-sparse-duplicate-columns.seed",
                "v2-descending-rows.seed",
                "v2-duplicate-columns.seed",
                "v2-duplicate-rows.seed",
                "v2-empty.seed",
                "v2-huge-string-tiny.seed",
                "v2-max-cell-count.seed",
                "v2-max-row-count-tiny.seed",
                "v2-valid-number.seed",
            ],
        ),
    ] {
        let mut target_files: Vec<_> = fs::read_dir(root.join(directory))
            .unwrap()
            .map(|entry| entry.unwrap().path())
            .filter(|path| path.is_file())
            .collect();
        target_files.sort();
        let actual_names: Vec<_> = target_files
            .iter()
            .map(|path| path.file_name().unwrap().to_string_lossy().into_owned())
            .collect();
        assert_eq!(
            actual_names, expected_names,
            "canonical corpus inventory changed for {directory}"
        );
        files.extend(target_files.into_iter().map(|path| (target, path)));
    }
    files
}

fn exercise(target: usize, input: &[u8]) -> Vec<Outcome> {
    match target {
        0 => {
            let (metadata, eager, lazy) = exercise_zip_input(input);
            vec![metadata, eager, lazy]
        }
        1 => {
            let (a1, compact, worksheet) = exercise_xml_input(input);
            vec![a1, compact, worksheet]
        }
        2 => {
            let (parse, evaluate) = exercise_formula_input(input);
            vec![parse, evaluate]
        }
        3 => vec![exercise_transfer_input(input)],
        _ => unreachable!(),
    }
}

fn assert_canonical_outcome(target: usize, path: &Path, outcomes: &[Outcome]) {
    let name = path.file_name().unwrap().to_string_lossy();
    match target {
        0 => match name.as_ref() {
            "minimal-workbook.seed" => assert_eq!(
                outcomes,
                [Outcome::Accepted, Outcome::Accepted, Outcome::Accepted]
            ),
            "entry-count-exact.seed" | "zip64-sector-exact.seed" => {
                assert_eq!(outcomes[0], Outcome::Accepted, "{}", path.display());
            }
            "entry-count-over.seed" | "zip64-sector-over.seed" => {
                assert_eq!(outcomes[0], Outcome::Rejected, "{}", path.display());
            }
            "misleading-declared-size.seed" => assert_eq!(
                outcomes,
                [Outcome::Accepted, Outcome::Rejected, Outcome::Rejected]
            ),
            _ => {}
        },
        1 => match name.as_ref() {
            "canonical-min.seed" | "canonical-max.seed" => {
                assert_eq!(outcomes[2], Outcome::Accepted, "{}", path.display());
            }
            "direct-canonical-min.seed" | "direct-canonical-max.seed" => assert_eq!(
                outcomes,
                [Outcome::Accepted, Outcome::Accepted, Outcome::Rejected]
            ),
            "direct-over.seed" => assert_eq!(
                outcomes,
                [Outcome::Rejected, Outcome::Rejected, Outcome::Rejected]
            ),
            _ => assert_eq!(outcomes[2], Outcome::Rejected, "{}", path.display()),
        },
        2 => match name.as_ref() {
            "depth-over.seed" | "binary-chain-over.seed" => {
                assert_eq!(outcomes[0], Outcome::Rejected, "{}", path.display());
            }
            _ => assert_eq!(outcomes[0], Outcome::Accepted, "{}", path.display()),
        },
        3 => match name.as_ref() {
            "v1-empty.seed" | "v2-empty.seed" | "v2-valid-number.seed" => {
                assert_eq!(outcomes[0], Outcome::Accepted, "{}", path.display());
            }
            _ => assert_eq!(outcomes[0], Outcome::Rejected, "{}", path.display()),
        },
        _ => unreachable!(),
    }
}

fn structured_mutation(seed: &[u8], state: u64, iteration: usize) -> Vec<u8> {
    let mut bytes = seed.to_vec();
    if bytes.is_empty() {
        bytes.push(state as u8);
        return bytes;
    }
    match iteration % 4 {
        0 => bytes.truncate(state as usize % (bytes.len() + 1)),
        1 => {
            let offset = state as usize % bytes.len();
            bytes[offset] ^= 1 << ((state >> 8) & 7);
        }
        2 => {
            let offset = state as usize % bytes.len();
            for (index, byte) in state.to_le_bytes().iter().take(4).enumerate() {
                if let Some(slot) = bytes.get_mut(offset + index) {
                    *slot = *byte;
                }
            }
        }
        _ => bytes.extend_from_slice(&state.to_le_bytes()[..(state as usize % 8) + 1]),
    }
    bytes
}
