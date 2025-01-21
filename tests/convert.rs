use std::{
    path::PathBuf,
    rc::Rc,
    sync::{
        atomic::{AtomicBool, Ordering},
        Mutex,
    },
};

use libreofficekit::{CallbackType, DocUrl, DocumentType, Office, OfficeOptionalFeatures};
use tempfile::{tempdir, TempDir};

fn temp_file(name: &str) -> (PathBuf, TempDir) {
    let temp_dir = tempdir().unwrap();
    let output_path = temp_dir.path().join(name);

    (output_path, temp_dir)
}

pub static OFFICE_TEST_LOCK: Mutex<()> = Mutex::new(());

#[test]
fn test_sample_docx() {
    let _lock = OFFICE_TEST_LOCK.lock();

    let office = Office::new(Office::find_install_path().unwrap()).unwrap();

    let (output_path, _temp_dir) = temp_file("test-sample.pdf");
    let input_url = DocUrl::from_relative_path("./tests/samples/sample-docx.docx").unwrap();
    let output_url = DocUrl::from_path(output_path).unwrap();

    let mut document = office.document_load(&input_url).unwrap();

    let document_type = document.get_document_type().unwrap();

    assert_eq!(document_type, DocumentType::Text);

    let _doc = document.save_as(&output_url, "pdf", None).unwrap();
}

#[test]
fn test_sample_docx_many() {
    let _lock = OFFICE_TEST_LOCK.lock();

    let office = Office::new(Office::find_install_path().unwrap()).unwrap();

    for _ in 0..5 {
        let (output_path, _temp_dir) = temp_file("test-sample.pdf");
        let input_url = DocUrl::from_relative_path("./tests/samples/sample-docx.docx").unwrap();
        let output_url = DocUrl::from_path(output_path).unwrap();

        let mut document = office.document_load(&input_url).unwrap();

        let document_type = document.get_document_type().unwrap();

        assert_eq!(document_type, DocumentType::Text);

        let _doc = document.save_as(&output_url, "pdf", None).unwrap();
    }
}

#[test]
fn test_sample_docx_encrypted() {
    let _lock = OFFICE_TEST_LOCK.lock();

    let office = Office::new(Office::find_install_path().unwrap()).unwrap();

    let input_url =
        DocUrl::from_relative_path("./tests/samples/sample-docx-encrypted.docx").unwrap();
    let needs_password = Rc::new(AtomicBool::new(false));

    // Allow password requests
    office
        .set_optional_features(OfficeOptionalFeatures::DOCUMENT_PASSWORD)
        .unwrap();

    office
        .register_callback({
            // Copies of local variables to include in the callback
            let needs_password = needs_password.clone();
            let input_url = input_url.clone();

            // Callback itself
            move |office, ty, _| {
                if let CallbackType::DocumentPassword = ty {
                    // Password was requested
                    needs_password.store(true, Ordering::SeqCst);

                    // Provide "I don't have the password"
                    office.set_document_password(&input_url, None).unwrap();
                }
            }
        })
        .unwrap();

    // Document fails to load
    assert!(office.document_load(&input_url).is_err());

    // Password was requested
    assert!(needs_password.load(Ordering::SeqCst));
}

#[test]
fn test_sample_docx_encrypted_then_normal() {
    let _lock = OFFICE_TEST_LOCK.lock();

    let office = Office::new(Office::find_install_path().unwrap()).unwrap();

    {
        let input_url =
            DocUrl::from_relative_path("./tests/samples/sample-docx-encrypted.docx").unwrap();
        let needs_password = Rc::new(AtomicBool::new(false));

        // Allow password requests
        office
            .set_optional_features(OfficeOptionalFeatures::DOCUMENT_PASSWORD)
            .unwrap();

        office
            .register_callback({
                // Copies of local variables to include in the callback
                let needs_password = needs_password.clone();
                let input_url = input_url.clone();

                // Callback itself
                move |office, ty, _| {
                    if let CallbackType::DocumentPassword = ty {
                        // Password was requested
                        needs_password.store(true, Ordering::SeqCst);

                        // Provide "I don't have the password"
                        office.set_document_password(&input_url, None).unwrap();
                    }
                }
            })
            .unwrap();

        // Document fails to load
        assert!(office.document_load(&input_url).is_err());

        // Password was requested
        assert!(needs_password.load(Ordering::SeqCst));
    }

    {
        let (output_path, _temp_dir) = temp_file("test-sample.pdf");
        let input_url = DocUrl::from_relative_path("./tests/samples/sample-docx.docx").unwrap();
        let output_url = DocUrl::from_path(output_path).unwrap();

        let mut document = office.document_load(&input_url).unwrap();

        let document_type = document.get_document_type().unwrap();

        assert_eq!(document_type, DocumentType::Text);

        let _doc = document.save_as(&output_url, "pdf", None).unwrap();
    }
}

#[test]
fn test_sample_docx_encrypted_known_password() {
    let _lock = OFFICE_TEST_LOCK.lock();

    let office = Office::new(Office::find_install_path().unwrap()).unwrap();

    let input_url =
        DocUrl::from_relative_path("./tests/samples/sample-docx-encrypted.docx").unwrap();
    let needs_password = Rc::new(AtomicBool::new(false));

    // Allow password requests
    office
        .set_optional_features(OfficeOptionalFeatures::DOCUMENT_PASSWORD)
        .unwrap();

    office
        .register_callback({
            // Copies of local variables to include in the callback
            let needs_password = needs_password.clone();
            let input_url = input_url.clone();

            // Callback itself
            move |office, ty, _| {
                if let CallbackType::DocumentPassword = ty {
                    // Password was requested
                    if needs_password.swap(true, Ordering::SeqCst) {
                        // Password we provided was incorrect, clear password to prevent infinite callback loop
                        office.set_document_password(&input_url, None).unwrap();
                        return;
                    }

                    // Provide the password
                    office
                        .set_document_password(&input_url, Some("password"))
                        .unwrap();
                }
            }
        })
        .unwrap();

    // Document loads
    let _document = office.document_load(&input_url).unwrap();

    // Password was requested
    assert!(needs_password.load(Ordering::SeqCst));
}

#[test]
fn test_sample_xlsx() {
    let _lock = OFFICE_TEST_LOCK.lock();

    let office = Office::new(Office::find_install_path().unwrap()).unwrap();

    let (output_path, _temp_dir) = temp_file("sample-xlsx.pdf");
    let input_url = DocUrl::from_relative_path("./tests/samples/sample-xlsx.xlsx").unwrap();
    let output_url = DocUrl::from_path(output_path).unwrap();

    let mut document = office.document_load(&input_url).unwrap();

    let document_type = document.get_document_type().unwrap();

    assert_eq!(document_type, DocumentType::Spreadsheet);

    let _doc = document.save_as(&output_url, "pdf", None).unwrap();
}

#[test]
fn test_sample_txt() {
    let _lock = OFFICE_TEST_LOCK.lock();

    let office = Office::new(Office::find_install_path().unwrap()).unwrap();

    let (output_path, _temp_dir) = temp_file("sample-txt.pdf");
    let input_url = DocUrl::from_relative_path("./tests/samples/sample-text.txt").unwrap();
    let output_url = DocUrl::from_path(output_path).unwrap();

    let mut document = office.document_load(&input_url).unwrap();

    let document_type = document.get_document_type().unwrap();

    assert_eq!(document_type, DocumentType::Text);

    let _doc = document.save_as(&output_url, "pdf", None).unwrap();
}
