pub mod error;
mod sys;
pub mod urls;

use std::{
    collections::HashMap,
    ffi::{c_ulonglong, CString},
    os::raw::{c_char, c_int},
    path::Path,
    rc::{Rc, Weak},
    sync::atomic::Ordering,
};

use bitflags::bitflags;
use num_enum::FromPrimitive;
use serde::Deserialize;

pub use error::OfficeError;
use sys::GLOBAL_OFFICE_LOCK;
pub use urls::DocUrl;

/// Instance of office.
///
/// The underlying raw logic is NOT thread safe
///
/// You cannot use more than one instance at a time in a single process
/// across threads or it will cause a segmentation fault so instance
/// creation is restricted with a static global lock
#[derive(Clone)]
pub struct Office {
    raw: Rc<sys::OfficeRaw>,
}

/// Instance of [Office] provided to callbacks
///
/// Only holds a week reference which is passed to callback
/// functions, provides functions that should only be used
/// from within the callback
#[derive(Clone)]
pub struct CallbackOffice {
    raw: Weak<sys::OfficeRaw>,
}

impl CallbackOffice {
    /// Creates a full [Office] reference from the callback value
    pub fn into_office(self) -> Result<Office, OfficeError> {
        // Obtain raw access
        let raw = self.raw.upgrade().ok_or(OfficeError::InstanceDropped)?;
        Ok(Office { raw })
    }

    /// Sets the password office should try to decrypt the document with.
    ///
    /// Passwords will only be requested if the [OfficeOptionalFeatures::DOCUMENT_PASSWORD]
    /// optional feature is enabled, otherwise the callback will not run.
    ///
    /// ## Important
    ///
    /// Set to [None] when the password should not be used / is unknown. The
    /// callback will continue to be invoked until either the correct password
    /// is specified or [None] is provided.
    pub fn set_document_password(
        &self,
        url: &DocUrl,
        password: Option<&str>,
    ) -> Result<(), OfficeError> {
        // Obtain raw access
        let raw = self.raw.upgrade().ok_or(OfficeError::InstanceDropped)?;

        // Unwrapping the default value ensures an empty string for empty passwords
        let value = CString::new(password.unwrap_or_default())?;

        unsafe { raw.set_document_password(url, value.as_ptr())? };

        Ok(())
    }
}

impl Office {
    /// Attempts to find an installation path from one of the common system install
    /// locations
    pub fn find_install_path() -> Option<&'static str> {
        const KNOWN_PATHS: &[&str] = &[
            "/usr/lib64/libreoffice/program",
            "/usr/lib/libreoffice/program",
        ];

        // Check common paths
        KNOWN_PATHS
            .iter()
            .find(|&path| Path::new(path).exists())
            .copied()
    }

    /// Creates a new LOK instance from the provided install path
    pub fn new(install_path: &str) -> Result<Office, OfficeError> {
        // Try lock the global office lock
        if GLOBAL_OFFICE_LOCK.swap(true, Ordering::SeqCst) {
            return Err(OfficeError::InstanceLock);
        }

        let install_path = CString::new(install_path)?;
        let raw = match unsafe { sys::OfficeRaw::init(install_path.as_ptr()) } {
            Ok(value) => value,
            Err(err) => {
                // Unlock the global office lock on init failure
                GLOBAL_OFFICE_LOCK.store(false, Ordering::SeqCst);
                return Err(err);
            }
        };

        // Check initialization errors
        if let Some(err) = unsafe { raw.get_error() } {
            return Err(OfficeError::OfficeError(err));
        }

        Ok(Office { raw: Rc::new(raw) })
    }

    pub fn get_filter_types(&self) -> Result<FilterTypes, OfficeError> {
        let value = unsafe { self.raw.get_filter_types()? };

        let value = value.to_str().map_err(OfficeError::InvalidUtf8String)?;

        let value: FilterTypes =
            serde_json::from_str(value).map_err(OfficeError::InvalidFilterTypes)?;

        Ok(value)
    }

    pub fn get_version_info(&self) -> Result<OfficeVersionInfo, OfficeError> {
        let value = unsafe { self.raw.get_version_info()? };

        let value = value.to_str().map_err(OfficeError::InvalidUtf8String)?;

        let value: OfficeVersionInfo =
            serde_json::from_str(value).map_err(OfficeError::InvalidVersionInfo)?;

        Ok(value)
    }

    pub fn sign_document(
        &self,
        url: &DocUrl,
        certificate: &[u8],
        private_key: &[u8],
    ) -> Result<bool, OfficeError> {
        // Lengths cannot exceed signed 32bit limit
        debug_assert!(certificate.len() <= i32::MAX as usize);
        debug_assert!(private_key.len() <= i32::MAX as usize);

        let result = unsafe {
            self.raw.sign_document(
                url,
                certificate.as_ptr(),
                certificate.len() as i32,
                private_key.as_ptr(),
                private_key.len() as i32,
            )?
        };

        Ok(result)
    }

    pub fn document_load(&self, url: &DocUrl) -> Result<Document, OfficeError> {
        let raw = unsafe { self.raw.document_load(url)? };
        Ok(Document { raw })
    }

    pub fn document_load_with_options(
        &self,
        url: &DocUrl,
        options: &str,
    ) -> Result<Document, OfficeError> {
        let options = CString::new(options)?;
        let raw = unsafe { self.raw.document_load_with_options(url, options.as_ptr())? };
        Ok(Document { raw })
    }

    pub fn send_dialog_event(
        &self,
        window_id: c_ulonglong,
        arguments: &str,
    ) -> Result<(), OfficeError> {
        let arguments = CString::new(arguments)?;

        unsafe { self.raw.send_dialog_event(window_id, arguments.as_ptr())? };

        Ok(())
    }

    pub fn set_optional_features(
        &self,
        features: OfficeOptionalFeatures,
    ) -> Result<(), OfficeError> {
        unsafe { self.raw.set_optional_features(features.bits())? };

        Ok(())
    }

    pub fn register_callback<F>(&self, mut callback: F) -> Result<(), OfficeError>
    where
        F: FnMut(CallbackOffice, CallbackType, *const c_char) + 'static,
    {
        // Create an office instance to use within the callbacks
        let callback_office = CallbackOffice {
            raw: Rc::downgrade(&self.raw),
        };

        // Create callback wrapper that maps the type
        let callback = move |ty, payload| {
            let callback_office = Clone::clone(&callback_office);
            let ty = CallbackType::from_primitive(ty);
            callback(callback_office, ty, payload)
        };

        unsafe {
            self.raw.register_callback(callback)?;
        }

        Ok(())
    }

    pub fn clear_callback(&self) -> Result<(), OfficeError> {
        unsafe {
            self.raw.clear_callback()?;
        }

        Ok(())
    }

    pub fn run_macro(&self, url: &str) -> Result<bool, OfficeError> {
        let url = CString::new(url)?;
        let result = unsafe { self.raw.run_macro(url.as_ptr())? };
        Ok(result)
    }

    pub fn dump_state(&self) -> Result<String, OfficeError> {
        let value = unsafe { self.raw.dump_state()? };
        Ok(value.to_string_lossy().to_string())
    }

    pub fn set_option(&self, option: &str, value: &str) -> Result<(), OfficeError> {
        let option = CString::new(option)?;
        let value = CString::new(value)?;

        unsafe { self.raw.set_option(option.as_ptr(), value.as_ptr())? }

        Ok(())
    }

    /// Negative number tells LibreOffice to re-fill its memory caches
    ///
    /// Large positive number (>=1000) encourages immediate maximum memory saving.
    pub fn trim_memory(&self, target: c_int) -> Result<(), OfficeError> {
        unsafe { self.raw.trim_memory(target)? };

        Ok(())
    }

    /// Only exposed when destroy on drop is not enabled
    #[cfg(not(feature = "destroy_on_drop"))]
    pub fn destroy(&mut self) {
        unsafe { self.raw.destroy() }
    }
}

pub struct Document {
    /// Raw inner document
    raw: sys::DocumentRaw,
}

impl Document {
    /// Saves the document as another format
    pub fn save_as(
        &mut self,
        url: &DocUrl,
        format: &str,
        filter: Option<&str>,
    ) -> Result<bool, OfficeError> {
        let format: CString = CString::new(format)?;
        let filter: CString = CString::new(filter.unwrap_or_default())?;
        let result = unsafe { self.raw.save_as(url, format.as_ptr(), filter.as_ptr())? };

        Ok(result != 0)
    }

    /// Only exposed when destroy on drop is not enabled
    #[cfg(not(feature = "destroy_on_drop"))]
    pub fn destroy(&mut self) {
        unsafe { self.raw.destroy() }
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

bitflags! {
    /// Optional features of LibreOfficeKit, in particular callbacks that block
    /// LibreOfficeKit until the corresponding reply is received, which would
    /// deadlock if the client does not support the feature.
    ///
    /// @see [Office::set_optional_features]
    pub struct OfficeOptionalFeatures: u64 {
        /// Handle `LOK_CALLBACK_DOCUMENT_PASSWORD` by prompting the user for a password.
        ///
        /// @see [Office::set_document_password]
        const DOCUMENT_PASSWORD = 1 << 0;

        /// Handle `LOK_CALLBACK_DOCUMENT_PASSWORD_TO_MODIFY` by prompting the user for a password.
        ///
        /// @see [Office::set_document_password]
        const DOCUMENT_PASSWORD_TO_MODIFY = 1 << 1;

        /// Request to have the part number as a 5th value in the `LOK_CALLBACK_INVALIDATE_TILES` payload.
        const PART_IN_INVALIDATION_CALLBACK = 1 << 2;

        /// Turn off tile rendering for annotations.
        const NO_TILED_ANNOTATIONS = 1 << 3;

        /// Enable range based header data.
        const RANGE_HEADERS = 1 << 4;

        /// Request to have the active view's Id as the 1st value in the `LOK_CALLBACK_INVALIDATE_VISIBLE_CURSOR` payload.
        const VIEWID_IN_VISCURSOR_INVALIDATION_CALLBACK = 1 << 5;
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
