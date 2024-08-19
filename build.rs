use std::env;
use std::path::{Path, PathBuf};
use std::process::Command;

// perform make with argument
fn make(path: &str) {
    let out_dir = env::var("OUT_DIR").unwrap();

    let status = Command::new("gcc")
        .args(["src/wrapper.c", "-c", "-fPIC", &format!("-I{path}"), "-o"])
        .arg(format!("{out_dir}/wrapper.o"))
        .status()
        .unwrap();
    if !status.success() {
        panic!(
            "make wrapper returns {:?}, maybe LO_INCLUDE_PATH is empty",
            status.code().unwrap()
        );
    }
    Command::new("ar")
        .args(["crus", "libwrapper.a", "wrapper.o"])
        .current_dir(Path::new(&out_dir))
        .status()
        .unwrap();
    println!("cargo:rustc-link-search=native={out_dir}");
}

fn generate_binding(path: &str) {
    let bindings = bindgen::Builder::default()
        .header("src/wrapper.h")
        .layout_tests(false)
        .clang_arg(format!("-I{path}"))
        .generate()
        .expect("Unable to generate bindings");
    let out_path = PathBuf::from(env::var("OUT_DIR").unwrap());

    bindings
        .write_to_file(out_path.join("bindings.rs"))
        .expect("Couldn't write bindings!");
}

fn main() {
    let lo_include_path =
        std::env::var("LO_INCLUDE_PATH").unwrap_or_else(|_| "/usr/include/LibreOfficeKit".into());
    println!("cargo:rerun-if-changed={lo_include_path}");
    make(&lo_include_path);
    generate_binding(&lo_include_path);
    println!("cargo:rustc-link-lib=static=wrapper");
}
