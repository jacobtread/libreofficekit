# LibreOfficeKit

Rust library providing safe access to the LibreOfficeSDK (LOK)

This library provides functionality for:
- Converting documents between various office and non-office formats (docx, xlsx, odt, ...etc into PDF and other supported formats see [File Conversion Filter Names](https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html))
- Cryptographically signing documents
- Obtaining available file filters from LibreOffice
- Obtaining LibreOffice version information
- Executing document macros
- Determine document type

This library does not link to the LibreOfficeKit C++ headers like other implementations, I have move that code into the Rust implementation so a C++ build toolchain is not required to build this library

For examples of how to use the library in a real life setting check out the [Office convert server](https://github.com/jacobtread/office-convert-server) which is a real instance of this library being used in production.

## LibreOffice Support 

Tested against Libreoffice versions 6.4.7.2 and 25.2.6.2 should be compatible with versions supported by the standard LOK C++ library. 

Certain functions are only available in certain LibreOffice versions, calling these on a unsupported install will return `OfficeError::MissingFunction` as the error.

You can also use `Office::get_version_info` which will provide a `OfficeVersionInfo` structure which contains a `product_version` field with helper functions such as `is_free_error_available` which tells you whether a specific function should be available for that version

> [!IMPORTANT]
>
> LibreOffice has some broken behavior in the newer versions where some process end cleanup logic causes a segmentation fault when the program exists.
>
> Versions that are not affected by this LibreOffice bug are versions 6.x with 
> the latest being 6.4.7.2, any versions newer than this seem to be affected by 
> this bug. This bug is present in all versions newer than 6.x including latest (25.2.0.0.alpha0+ as at 1 Sep 2024)
>
> You can find downloads to 6.4.7.2 on the download archives https://downloadarchive.documentfoundation.org/libreoffice/old/6.4.7.2/ but its recommended you use the latest version instead in most cases.



## Windows Support

This library can be run and compiled on Windows. However, the `Office::find_install_path()` will only find valid paths on Linux, for Windows you will need to manually specify the path to your LibreOffice installation

## Converting a file

To convert an office file format (docx, xlsx, odt, ...etc) into a PDF you can use the following code:

```rust
let office = Office::new(Office::find_install_path().unwrap()).unwrap();

let input_url = DocUrl::from_relative_path("./tests/samples/sample-docx.docx").unwrap();
let output_url = DocUrl::from_absolute_path("/tmp/test.pdf").unwrap();

let mut document = office.document_load(&input_url).unwrap();

let success = document.save_as(&output_url, "pdf", None).unwrap();

if !success {
    // ...Document conversion failed
}

// ...Do something with the file at output_url

```

> [!NOTE]
>
> You can find the full supported list of conversion formats on the LibreOffice website
> [Here](https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html)


## Loading a password protected file

You can load password protected office documents using the code below:

```rust
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
                    // Password we provided was incorrect, you must clear the password to prevent infinite callback loop
                    // the callback will be called until the correct password (Or None) is provided
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
let document = office.document_load(&input_url).unwrap();

// Check Password was requested
assert!(needs_password.load(Ordering::SeqCst));

// ...Do something with document
```

> [!IMPORTANT]
>
> Ensure you always specify a `None` password on failure or if you don't have a password (and have specified the `OfficeOptionalFeatures::DOCUMENT_PASSWORD` optional feature) LibreOffice will continue to block the `document_load` call and repeatedly invoke the callback until either the correct password is given or `None` is provided

## Freeing memory

LibreOffice will accumulate buffers over time as you convert/load documents, if you are using LOK in a long running process you will want to use the `Office::trim_memory` function to free some of that memory:

```rust
let office = Office::new(Office::find_install_path().unwrap()).unwrap();

// ... Do some document loading and conversion 

office.trim_memory(2000).unwrap();
```

> [!NOTE]
> Negative number provided to `trim_memory` tells LibreOffice to re-fill its memory caches
>
> Large positive number (>=1000) encourages immediate maximum memory saving.

## Credits

The original implementation of this library was based upon https://github.com/undeflife/libreoffice-rs aiming to be more complex and cover more
use cases, to use as a backend for an Office file format conversion server. 