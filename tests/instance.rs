use libreofficekit::Office;
use parking_lot::Mutex;

/// Mutex to prevent parallel test runs
static TEST_MUTEX: Mutex<()> = Mutex::new(());

/// Tests that an office instance can be found
#[test]
fn test_find_office_instance() {
    let _lock = TEST_MUTEX.lock();

    let office_path = Office::find_install_path();
    office_path.expect("missing office install path");
}

/// Tests that an office instance can be created
#[test]
fn test_create_office_instance() {
    let _lock = TEST_MUTEX.lock();

    let office_path = Office::find_install_path().expect("missing office install path");
    let _office = Office::new(office_path).expect("failed to create office instance");
}

/// Tests obtaining the LibreOffice version information
#[test]
fn test_version_info() {
    let _lock = TEST_MUTEX.lock();

    let office_path = Office::find_install_path().expect("missing office install path");
    let office = Office::new(office_path).expect("failed to create office instance");
    let version_info = office.get_version_info().unwrap();

    dbg!(version_info);
}

/// Tests obtaining the available filter types
#[test]
fn test_filter_types() {
    let _lock = TEST_MUTEX.lock();

    let office_path = Office::find_install_path().expect("missing office install path");
    let office = Office::new(office_path).expect("failed to create office instance");
    let filter_types = office.get_filter_types().unwrap();

    dbg!(filter_types);
}
