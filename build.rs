use std::env;
use std::path::PathBuf;

fn main() {
    // Determine compile output dir
    let out_dir = env::var("OUT_DIR").expect("missing OUT_DIR");

    // Determine libreoffice include path
    let lo_include_path =
        std::env::var("LO_INCLUDE_PATH").unwrap_or_else(|_| "/usr/include/LibreOfficeKit".into());

    // Rebuild if the libreoffice source files change
    println!("cargo:rerun-if-changed={lo_include_path}");

    // Build the wrapper library
    // Compile the wrapper
    cc::Build::new()
        .cpp(true)
        .file("src/wrapper.cpp")
        .flag("-fPIC")
        // Include the libreoffice headers
        .include(&lo_include_path)
        // Suppress warnings for unsafe functions
        .define("_CRT_SECURE_NO_WARNINGS", None)
        .compile("wrapper");

    // Add the out dir to the link search path
    println!("cargo:rustc-link-search=native={}", out_dir);

    // Re-run build if the wrapper changes
    println!("cargo:rerun-if-changed=src/wrapper.c");

    // Generate bindings to the library
    let bindings = bindgen::Builder::default()
        .header("src/wrapper.hpp")
        .layout_tests(false)
        .clang_arg(format!("-I{lo_include_path}"))
        .clang_arg("-std=c++14")
        .allowlist_type("LibreOfficeKit")
        .allowlist_type("LibreOfficeKitDocument")
        .allowlist_function("lok_init_wrapper")
        .generate()
        .expect("Unable to generate bindings");
    let out_path = PathBuf::from(out_dir);

    bindings
        .write_to_file(out_path.join("bindings.rs"))
        .expect("Couldn't write bindings!");

    // Tell the linker to statically link to the wrapper
    println!("cargo:rustc-link-lib=static=wrapper");
}
