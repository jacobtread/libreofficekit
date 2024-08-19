#![allow(
    dead_code,
    non_snake_case,
    non_camel_case_types,
    non_upper_case_globals
)]
#![allow(clippy::all)]
include!(concat!(env!("OUT_DIR"), "/bindings.rs"));

mod error;
pub mod urls;

use error::Error;
use num_enum::FromPrimitive;
use serde::Deserialize;
use urls::DocUrl;

use std::{
    ffi::{c_char, c_int, c_void, CStr, CString},
    os::raw::c_ulonglong,
    ptr::null_mut,
};

/// A Wrapper for the `LibreOfficeKit` C API.
#[derive(Clone)]
pub struct Office {
    lok: *mut LibreOfficeKit,
    lok_class: *mut LibreOfficeKitClass,
}

/// A Wrapper for the `LibreOfficeKitDocument` C API.
pub struct Document {
    doc: *mut LibreOfficeKitDocument,
}

/// Optional features of LibreOfficeKit, in particular callbacks that block
///  LibreOfficeKit until the corresponding reply is received, which would
///  deadlock if the client does not support the feature.
///
///  @see [Office::set_optional_features]
#[derive(Copy, Clone)]
pub enum LibreOfficeKitOptionalFeatures {
    /// Handle `LOK_CALLBACK_DOCUMENT_PASSWORD` by prompting the user for a password.
    ///
    /// @see [Office::set_document_password]
    LOK_FEATURE_DOCUMENT_PASSWORD = (1 << 0),

    /// Handle `LOK_CALLBACK_DOCUMENT_PASSWORD_TO_MODIFY` by prompting the user for a password.
    ///
    /// @see [Office::set_document_password]
    LOK_FEATURE_DOCUMENT_PASSWORD_TO_MODIFY = (1 << 1),

    /// Request to have the part number as an 5th value in the `LOK_CALLBACK_INVALIDATE_TILES` payload.
    LOK_FEATURE_PART_IN_INVALIDATION_CALLBACK = (1 << 2),

    /// Turn off tile rendering for annotations
    LOK_FEATURE_NO_TILED_ANNOTATIONS = (1 << 3),

    /// Enable range based header data
    LOK_FEATURE_RANGE_HEADERS = (1 << 4),

    /// Request to have the active view's Id as the 1st value in the `LOK_CALLBACK_INVALIDATE_VISIBLE_CURSOR` payload.
    LOK_FEATURE_VIEWID_IN_VISCURSOR_INVALIDATION_CALLBACK = (1 << 5),
}

/// Obtains an error message string from the provided LOK instance
fn get_error(lok: *mut LibreOfficeKit, lok_class: *mut LibreOfficeKitClass) -> Option<String> {
    unsafe {
        let get_error = (*lok_class).getError.expect("missing getError function");
        let raw_error = get_error(lok);

        // Empty error is considered to be no error
        if *raw_error == 0 {
            return None;
        }

        // Create rust copy of the error message
        let value = CStr::from_ptr(raw_error).to_string_lossy().into_owned();

        // Free error memory
        free_error(lok_class, raw_error);

        Some(value)
    }
}

/// Frees the error memory
fn free_error(lok_class: *mut LibreOfficeKitClass, error: *mut i8) {
    unsafe {
        // Only available LibreOffice >=5.2
        if let Some(free_error) = (*lok_class).freeError {
            free_error(error);
        }
    }
}

impl Office {
    pub fn new(install_path: &str) -> Result<Office, Error> {
        let install_path =
            CString::new(install_path).expect("install path should not contain null byte");

        // Initialize lok
        let (lok, lok_class) = unsafe {
            let lok = lok_init_wrapper(install_path.as_ptr());
            let lok_class = (*lok).pClass;
            (lok, lok_class)
        };

        // Check initialization errors
        if let Some(err) = get_error(lok, lok_class) {
            return Err(Error::new(err));
        }

        Ok(Office { lok, lok_class })
    }

    fn destroy(&mut self) {
        unsafe {
            let destroy = (*self.lok_class).destroy.expect("missing destroy function");
            destroy(self.lok);
        }
    }

    /// Returns the last error as a string
    pub fn get_error(&mut self) -> Option<String> {
        get_error(self.lok, self.lok_class)
    }

    pub fn set_option(&mut self, option: &str, value: &str) {
        unsafe {
            let option = CString::new(option).unwrap();
            let value = CString::new(value).unwrap();

            let set_option = (*self.lok_class)
                .setOption
                .expect("missing setOption function");
            set_option(self.lok, option.as_ptr(), value.as_ptr());
        }
    }
    pub fn dump_state(&mut self) -> Result<String, Error> {
        unsafe {
            let mut state: *mut c_char = null_mut();
            let dump_state = (*self.lok_class)
                .dumpState
                .expect("missing dumpState function");
            dump_state(self.lok, std::ptr::null(), &mut state);
            if let Some(error) = self.get_error() {
                return Err(Error::new(error));
            }
            let value = CString::from_raw(state);

            Ok(value.to_string_lossy().to_string())
        }
    }

    pub fn register_callback<F: FnMut(CallbackType, *const std::os::raw::c_char)>(
        &mut self,
        callback: F,
    ) -> Result<(), Error> {
        /// Create a shim to wrap the callback function so it can be invoked
        unsafe extern "C" fn callback_shim(
            ty: std::os::raw::c_int,
            payload: *const std::os::raw::c_char,
            data: *mut std::os::raw::c_void,
        ) {
            // Get the callback function from the data argument
            let callback: *mut Box<dyn FnMut(CallbackType, *const std::os::raw::c_char)> =
                data.cast();

            let ty = CallbackType::from_primitive(ty);

            // Catch panics from calling the callback
            _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(move || {
                // Invoke the callback
                (**callback)(ty, payload);
            }));
        }

        // Callback is double boxed then leaked
        let callback_ptr: *mut Box<dyn FnMut(CallbackType, *const std::os::raw::c_char)> =
            Box::into_raw(Box::new(Box::new(callback)));

        unsafe {
            let register_callback = (*self.lok_class)
                .registerCallback
                .expect("missing registerCallback function");

            register_callback(self.lok, Some(callback_shim), callback_ptr.cast());
        }

        if let Some(error) = self.get_error() {
            return Err(Error::new(error));
        }

        Ok(())
    }

    pub fn document_load(&mut self, url: DocUrl) -> Result<Document, Error> {
        // Load the document
        let document = unsafe {
            let document_load = (*self.lok_class)
                .documentLoad
                .expect("missing documentLoad function");
            document_load(self.lok, url.as_ptr())
        };

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(Error::new(error));
        }

        Ok(Document { doc: document })
    }

    pub fn document_load_with_options(
        &mut self,
        url: DocUrl,
        options: &str,
    ) -> Result<Document, Error> {
        let options = CString::new(options).unwrap();
        // Load the document
        let document = unsafe {
            let document_load_with_options = (*self.lok_class)
                .documentLoadWithOptions
                .expect("missing documentLoad function");
            document_load_with_options(self.lok, url.as_ptr(), options.as_ptr())
        };

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(Error::new(error));
        }

        Ok(Document { doc: document })
    }

    pub fn trim_memory(&mut self, target: c_int) -> Result<(), Error> {
        unsafe {
            let trim_memory = (*self.lok_class)
                .trimMemory
                .expect("missing trimMemory function");
            trim_memory(self.lok, target)
        };

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(Error::new(error));
        }

        Ok(())
    }

    pub fn set_optional_features<T>(&mut self, optional_features: T) -> Result<u64, Error>
    where
        T: IntoIterator<Item = LibreOfficeKitOptionalFeatures>,
    {
        let feature_flags: u64 = optional_features
            .into_iter()
            .map(|i| i as u64)
            .fold(0, |acc, item| acc | item);

        unsafe {
            let set_optional_features = (*self.lok_class)
                .setOptionalFeatures
                .expect("missing setOptionalFeatures function");
            set_optional_features(self.lok, feature_flags);
        }

        if let Some(error) = self.get_error() {
            return Err(Error::new(error));
        }

        Ok(feature_flags)
    }

    pub fn set_document_password(
        &mut self,
        url: DocUrl,
        password: Option<&str>,
    ) -> Result<(), Error> {
        // Create a C compatible string
        let password =
            password.map(|value| CString::new(value).expect("password cannot contain null"));

        // Get the password ptr if one is set
        let password: *const c_char = match password {
            Some(value) => value.as_ptr(),
            None => std::ptr::null(),
        };

        unsafe {
            let set_document_password = (*self.lok_class)
                .setDocumentPassword
                .expect("missing setDocumentPassword function");

            set_document_password(self.lok, url.as_ptr(), password);
        }

        if let Some(error) = self.get_error() {
            return Err(Error::new(error));
        }

        Ok(())
    }

    pub fn document_load_with(&mut self, url: DocUrl, options: &str) -> Result<Document, Error> {
        let c_options = CString::new(options).unwrap();

        // Load the document
        let doc = unsafe {
            let document_load_with_options = (*self.lok_class)
                .documentLoadWithOptions
                .expect("missing documentLoadWithOptions function");

            document_load_with_options(self.lok, url.as_ptr(), c_options.as_ptr())
        };

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(Error::new(error));
        }

        debug_assert!(!doc.is_null());

        Ok(Document { doc })
    }

    pub fn send_dialog_event(
        &mut self,
        window_id: c_ulonglong,
        arguments: *const c_char,
    ) -> Result<(), Error> {
        unsafe {
            let send_dialog_event = (*self.lok_class)
                .sendDialogEvent
                .expect("missing sendDialogEvent function");

            send_dialog_event(self.lok, window_id, arguments);
        }

        if let Some(error) = self.get_error() {
            return Err(Error::new(error));
        }

        Ok(())
    }
    pub fn run_macro(&mut self, url: DocUrl) -> Result<(), Error> {
        let result = unsafe {
            let run_macro = (*self.lok_class)
                .runMacro
                .expect("missing runMacro function");

            run_macro(self.lok, url.as_ptr())
        };

        if result == 0 {
            if let Some(error) = self.get_error() {
                return Err(Error::new(error));
            }
        }

        Ok(())
    }
}

impl Drop for Office {
    fn drop(&mut self) {
        self.destroy()
    }
}

impl Document {
    /// Saves the document as another format
    pub fn save_as(&mut self, url: DocUrl, format: &str, filter: Option<&str>) -> bool {
        let format: CString = CString::new(format).expect("format cannot contain null byte");
        let filter: CString =
            CString::new(filter.unwrap_or_default()).expect("filter cannot contain null byte");
        let ret = unsafe { self.save_as_internal(url.as_ptr(), format.as_ptr(), filter.as_ptr()) };
        ret != 0
    }

    /// Handles the internal call to [LibreOfficeKitDocumentClass::saveAs]
    unsafe fn save_as_internal(
        &mut self,
        url: *const c_char,
        format: *const c_char,
        filter: *const c_char,
    ) -> i32 {
        let class = (*self.doc).pClass;
        let save_as = (*class).saveAs.expect("missing saveAs function");

        save_as(self.doc, url, format, filter)
    }

    fn destroy(&mut self) {
        unsafe {
            let class = (*self.doc).pClass;
            let destroy = (*class).destroy.expect("missing destroy function");
            destroy(self.doc);
        }
    }
}

impl Drop for Document {
    fn drop(&mut self) {
        self.destroy()
    }
}

#[derive(Debug, FromPrimitive, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub enum CallbackType {
    InvalidateTiles = 0,
    InvalidateVisibleCursor = 1,
    TextSelection = 2,
    TextSelectionStart = 3,
    TextSelectionEnd = 4,
    CursorVisible = 5,
    GraphicSelection = 6,
    HyperlinkClicked = 7,
    StateChanged = 8,
    StatusIndicatorStart = 9,
    StatusIndicatorSetValue = 10,
    StatusIndicatorFinish = 11,
    SearchNotFound = 12,
    DocumentSizeChanged = 13,
    SetPart = 14,
    SearchResultSelection = 15,
    UnoCommandResult = 16,
    CellCursor = 17,
    MousePointer = 18,
    CellFormula = 19,
    DocumentPassword = 20,
    DocumentPasswordModify = 21,
    Error = 22,
    ContextMenu = 23,
    InvalidateViewCursor = 24,
    TextViewSelection = 25,
    CellViewCursor = 26,
    GraphicViewSelection = 27,
    ViewCursorVisible = 28,
    ViewLock = 29,
    RedlineTableSizeChanged = 30,
    RedlineTableEntryModified = 31,
    Comment = 32,
    InvalidateHeader = 33,
    CellAddress = 34,
    RulerUpdate = 35,
    Window = 36,
    ValidityListButton = 37,
    ClipboardChanged = 38,
    ContextChanged = 39,
    SignatureStatus = 40,
    ProfileFrame = 41,
    CellSelectionArea = 42,
    CellAutoFillArea = 43,
    TableSelected = 44,
    ReferenceMarks = 45,
    JSDialog = 46,
    CalcFunctionList = 47,
    TabStopList = 48,
    FormFieldButton = 49,
    InvalidateSheetGeometry = 50,
    ValidityInputHelp = 51,
    DocumentBackgroundColor = 52,
    CommandedBlocked = 53,
    CellCursorFollowJump = 54,
    ContentControl = 55,
    PrintRanges = 56,
    FontsMissing = 57,
    MediaShape = 58,
    ExportFile = 59,
    ViewRenderState = 60,
    ApplicationBackgroundColor = 61,
    A11YFocusChanged = 62,
    A11YCaretChanged = 63,
    A11YTextSelectionChanged = 64,
    ColorPalettes = 65,
    DocumentPasswordReset = 66,
    A11YFocusedCellChanged = 67,
    A11YEditingInSelectionState = 68,
    A11YSelectionChanged = 69,
    CoreLog = 70,

    #[num_enum(catch_all)]
    Unknown(i32),
}

#[derive(Debug)]
pub struct JSDialog(pub serde_json::Value);

impl JSDialog {
    /// Get the ID field for the dialog.
    pub fn get_id(&self) -> Option<c_ulonglong> {
        let obj = self.0.as_object()?;
        obj.iter().find_map(|value| {
            if value.0.ne("id") {
                return None;
            }

            let value = value.1.as_u64()?;
            Some(value)
        })
    }
}

#[derive(Debug, Deserialize)]
pub struct JSDialogResponse {
    id: String,
    response: u64,
}
