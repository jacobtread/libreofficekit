#![allow(
    dead_code,
    non_snake_case,
    non_camel_case_types,
    non_upper_case_globals
)]
#![allow(clippy::all)]
include!(concat!(env!("OUT_DIR"), "/bindings.rs"));

pub mod urls;

use num_enum::FromPrimitive;
use serde::Deserialize;
use thiserror::Error;
use urls::DocUrl;

use std::{
    collections::HashMap,
    ffi::{c_char, c_int, CStr, CString},
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

#[derive(Debug, Deserialize)]
pub struct FilterTypes {
    pub values: HashMap<String, FilterType>,
}

#[derive(Debug, Deserialize)]
pub struct FilterType {
    #[serde(rename = "MediaType")]
    pub media_type: String,
}

#[derive(Debug, Deserialize)]
pub struct OfficeVersionInfo {
    #[serde(rename = "ProductName")]
    pub product_name: String,
    #[serde(rename = "ProductVersion")]
    pub product_version: String,
    #[serde(rename = "ProductExtension")]
    pub product_extension: String,
    #[serde(rename = "BuildId")]
    pub build_id: String,
}

#[derive(Debug, Error)]
pub enum OfficeError {
    /// Error message produced by office
    #[error("{0}")]
    OfficeError(String),

    /// Function is not available in the current office install
    #[error("missing '{0}' function")]
    MissingFunction(&'static str),

    /// Filter value was an invalid string
    #[error("invalid filter types str: {0}")]
    InvalidFilterValue(std::str::Utf8Error),

    /// Filter types could not be parsed
    #[error("failed to parse filters: {0}")]
    InvalidFilterTypes(serde_json::Error),

    /// Version info value was an invalid string
    #[error("invalid version info str: {0}")]
    InvalidVersionInfoValue(std::str::Utf8Error),

    /// Version info could not be parsed
    #[error("failed to parse version info: {0}")]
    InvalidVersionInfo(serde_json::Error),
}

impl Office {
    pub fn new(install_path: &str) -> Result<Office, OfficeError> {
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
            return Err(OfficeError::OfficeError(err));
        }

        Ok(Office { lok, lok_class })
    }

    /// Gets the available filter types for the office instance
    pub fn get_filter_types(&self) -> Result<FilterTypes, OfficeError> {
        let value = unsafe {
            let get_filter_types = (*self.lok_class)
                .getFilterTypes
                .ok_or(OfficeError::MissingFunction("getFilterTypes"))?;

            let value = get_filter_types(self.lok);
            CString::from_raw(value)
        };

        let value = value.to_str().map_err(OfficeError::InvalidFilterValue)?;

        let value: FilterTypes =
            serde_json::from_str(value).map_err(OfficeError::InvalidFilterTypes)?;

        Ok(value)
    }

    /// Gets the version details from office
    pub fn get_version_info(&self) -> Result<OfficeVersionInfo, OfficeError> {
        let value = unsafe {
            let get_version_info = (*self.lok_class)
                .getVersionInfo
                .ok_or(OfficeError::MissingFunction("getVersionInfo"))?;

            let value = get_version_info(self.lok);
            CString::from_raw(value)
        };

        let value = value
            .to_str()
            .map_err(OfficeError::InvalidVersionInfoValue)?;

        let value: OfficeVersionInfo =
            serde_json::from_str(value).map_err(OfficeError::InvalidVersionInfo)?;

        Ok(value)
    }

    /// Returns the last error from office as a string
    pub fn get_error(&mut self) -> Option<String> {
        get_error(self.lok, self.lok_class)
    }

    /// Exports the provided document and signs the content
    pub fn sign_document(
        &mut self,
        url: DocUrl,
        certificate: &[u8],
        private_key: &[u8],
    ) -> Result<bool, OfficeError> {
        debug_assert!(certificate.len() <= i32::MAX as usize);
        debug_assert!(private_key.len() <= i32::MAX as usize);

        let result = unsafe {
            let sign_document = (*self.lok_class)
                .signDocument
                .ok_or(OfficeError::MissingFunction("signDocument"))?;
            sign_document(
                self.lok,
                url.as_ptr(),
                certificate.as_ptr(),
                certificate.len() as i32,
                private_key.as_ptr(),
                private_key.len() as i32,
            )
        };

        Ok(result)
    }

    pub fn set_option(&mut self, option: &str, value: &str) -> Result<(), OfficeError> {
        let option = CString::new(option).expect("option cannot contain null");
        let value = CString::new(value).expect("value cannot contain null");

        unsafe {
            let set_option = (*self.lok_class)
                .setOption
                .ok_or(OfficeError::MissingFunction("setOption"))?;
            set_option(self.lok, option.as_ptr(), value.as_ptr());
        }

        Ok(())
    }

    /// Dumps the state from office as a string
    pub fn dump_state(&mut self) -> Result<String, OfficeError> {
        let value = unsafe {
            let mut state: *mut c_char = null_mut();
            let dump_state = (*self.lok_class)
                .dumpState
                .ok_or(OfficeError::MissingFunction("dumpState"))?;
            dump_state(self.lok, std::ptr::null(), &mut state);

            if let Some(error) = self.get_error() {
                return Err(OfficeError::OfficeError(error));
            }

            CString::from_raw(state)
        };

        Ok(value.to_string_lossy().to_string())
    }

    /// Registers a callback for office
    pub fn register_callback<F: FnMut(CallbackType, *const std::os::raw::c_char)>(
        &mut self,
        callback: F,
    ) -> Result<(), OfficeError> {
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
                .ok_or(OfficeError::MissingFunction("registerCallback"))?;

            register_callback(self.lok, Some(callback_shim), callback_ptr.cast());
        }

        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(())
    }

    pub fn document_load(&mut self, url: DocUrl) -> Result<Document, OfficeError> {
        // Load the document
        let document = unsafe {
            let document_load = (*self.lok_class)
                .documentLoad
                .ok_or(OfficeError::MissingFunction("documentLoad"))?;
            document_load(self.lok, url.as_ptr())
        };

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(Document { doc: document })
    }

    pub fn document_load_with_options(
        &mut self,
        url: DocUrl,
        options: &str,
    ) -> Result<Document, OfficeError> {
        let options = CString::new(options).expect("options cannot contain null");
        // Load the document
        let document = unsafe {
            let document_load_with_options = (*self.lok_class)
                .documentLoadWithOptions
                .ok_or(OfficeError::MissingFunction("documentLoadWithOptions"))?;
            document_load_with_options(self.lok, url.as_ptr(), options.as_ptr())
        };

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(Document { doc: document })
    }

    pub fn trim_memory(&mut self, target: c_int) -> Result<(), OfficeError> {
        unsafe {
            let trim_memory = (*self.lok_class)
                .trimMemory
                .ok_or(OfficeError::MissingFunction("trimMemory"))?;
            trim_memory(self.lok, target)
        };

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(())
    }

    pub fn set_optional_features<T>(&mut self, optional_features: T) -> Result<u64, OfficeError>
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
                .ok_or(OfficeError::MissingFunction("setOptionalFeatures"))?;
            set_optional_features(self.lok, feature_flags);
        }

        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(feature_flags)
    }

    pub fn set_document_password(
        &mut self,
        url: DocUrl,
        password: Option<&str>,
    ) -> Result<(), OfficeError> {
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
                .ok_or(OfficeError::MissingFunction("setDocumentPassword"))?;

            set_document_password(self.lok, url.as_ptr(), password);
        }

        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(())
    }

    pub fn document_load_with(
        &mut self,
        url: DocUrl,
        options: &str,
    ) -> Result<Document, OfficeError> {
        let c_options = CString::new(options).unwrap();

        // Load the document
        let doc = unsafe {
            let document_load_with_options = (*self.lok_class)
                .documentLoadWithOptions
                .ok_or(OfficeError::MissingFunction("documentLoadWithOptions"))?;

            document_load_with_options(self.lok, url.as_ptr(), c_options.as_ptr())
        };

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        debug_assert!(!doc.is_null());

        Ok(Document { doc })
    }

    pub fn send_dialog_event(
        &mut self,
        window_id: c_ulonglong,
        arguments: *const c_char,
    ) -> Result<(), OfficeError> {
        unsafe {
            let send_dialog_event = (*self.lok_class)
                .sendDialogEvent
                .ok_or(OfficeError::MissingFunction("sendDialogEvent"))?;

            send_dialog_event(self.lok, window_id, arguments);
        }

        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(())
    }

    pub fn run_macro(&mut self, url: &str) -> Result<(), OfficeError> {
        let url = CString::new(url).expect("macro url cannot include null");

        let result = unsafe {
            let run_macro = (*self.lok_class)
                .runMacro
                .ok_or(OfficeError::MissingFunction("runMacro"))?;

            run_macro(self.lok, url.as_ptr())
        };

        if result == 0 {
            if let Some(error) = self.get_error() {
                return Err(OfficeError::OfficeError(error));
            }
        }

        Ok(())
    }

    fn destroy(&mut self) {
        unsafe {
            // Destroy should be available in all versions
            let destroy = (*self.lok_class).destroy.expect("missing destroy function");
            destroy(self.lok);
        }
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
        let ret = unsafe {
            let class = (*self.doc).pClass;
            let save_as = (*class).saveAs.expect("missing saveAs function");

            save_as(self.doc, url.as_ptr(), format.as_ptr(), filter.as_ptr())
        };
        ret != 0
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
