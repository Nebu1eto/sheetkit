#![no_main]

use libfuzzer_sys::fuzz_target;

#[path = "../../security/fuzz_support.rs"]
mod fuzz_support;

fuzz_target!(|seed: &[u8]| {
    let input = fuzz_support::materialize_seed(seed);
    let _ = fuzz_support::exercise_transfer_input(&input);
});
