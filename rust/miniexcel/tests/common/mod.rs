use std::path::PathBuf;

pub fn fixture(relative_path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("tests")
        .join("data")
        .join("xlsx")
        .join(relative_path)
}
