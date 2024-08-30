# LibreOfficeKit

Wrapper around the C++ LibreOfficeKit library.

This is a modified and improved version of https://github.com/undeflife/libreoffice-rs aiming to be more complex and cover more
use cases, to use as a backend for an Office file format conversion server.


> [!IMPORTANT]
>
> LibreOffice has some broken behavior in the newer versions where the cleanup
> functions of Office cause segmentation faults, you must keep around a single 
> instance of Office around for the duration of your program

## Converting a file


```rust
let office = Office::new(Office::find_install_path().unwrap()).unwrap();

let input_url = DocUrl::from_relative_path("./tests/samples/sample-docx.docx").unwrap();
let output_url = DocUrl::from_absolute_path("/tmp/test.pdf").unwrap();

let mut document = office.document_load(&input_url).unwrap();

let _doc = document.save_as(&output_url, "pdf", None).unwrap();
```

