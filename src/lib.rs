pub mod error;
mod sys;
pub mod urls;

use std::{
    collections::HashMap,
    ffi::{c_ulonglong, CString},
    fmt::Display,
    os::raw::{c_char, c_int},
    path::{Path, PathBuf},
    ptr::null,
    rc::{Rc, Weak},
    str::FromStr,
    sync::atomic::Ordering,
};

use bitflags::bitflags;
use num_enum::FromPrimitive;
use serde::{Deserialize, Serialize};

pub use error::OfficeError;
use sys::GLOBAL_OFFICE_LOCK;
use thiserror::Error;
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

        let password = match password {
            Some(value) => CString::new(value)?,
            None => {
                // Password is unset
                unsafe { raw.set_document_password(url, null())? };
                return Ok(());
            }
        };

        unsafe { raw.set_document_password(url, password.as_ptr())? };

        Ok(())
    }
}

impl Office {
    /// Creates a new LOK instance from the provided install path
    pub fn new<P: Into<PathBuf>>(install_path: P) -> Result<Office, OfficeError> {
        // Try lock the global office lock
        if GLOBAL_OFFICE_LOCK.swap(true, Ordering::SeqCst) {
            return Err(OfficeError::InstanceLock);
        }

        let mut install_path: PathBuf = install_path.into();

        // Resolve non absolute paths
        if !install_path.is_absolute() {
            install_path =
                std::fs::canonicalize(install_path).map_err(|_| OfficeError::InvalidPath)?;
        }

        let install_path = install_path.to_str().ok_or(OfficeError::InvalidPath)?;

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

    /// Attempts to find an installation path from one of the common system install
    /// locations
    pub fn find_install_path() -> Option<PathBuf> {
        // Common set of install paths
        const KNOWN_PATHS: &[&str] = &[
            "/usr/lib64/libreoffice/program",
            "/usr/lib/libreoffice/program",
        ];

        // Check common paths
        if let Some(value) = KNOWN_PATHS.iter().find_map(|path| {
            let path = Path::new(path);
            if !path.exists() {
                return None;
            }

            Some(path.to_path_buf())
        }) {
            return Some(value);
        }

        // Search /opt for installs
        if let Ok(Some(latest)) = Self::find_opt_latest() {
            return Some(latest);
        }

        // No install found
        None
    }

    /// Finds all installations of LibreOffice from the `/opt` directory
    /// provides back a list of the paths along with the version extracted
    /// from the directory name
    pub fn find_opt_installs() -> std::io::Result<Vec<(ProductVersion, PathBuf)>> {
        let opt_path = Path::new("/opt");
        if !opt_path.exists() {
            return Ok(Vec::with_capacity(0));
        }

        // Find all libreoffice folders
        let installs: Vec<(ProductVersion, PathBuf)> = std::fs::read_dir(opt_path)?
            .filter_map(|value| value.ok())
            .filter_map(|value| {
                // Get entry file type
                let file_type = value.file_type().ok()?;

                // Ignore non directories
                if !file_type.is_dir() {
                    return None;
                }

                let dir_name = value.file_name();
                let dir_name = dir_name.to_str()?;

                // Only use dirs prefixed with libreoffice
                let version = dir_name.strip_prefix("libreoffice")?;

                // Only use valid product versions
                let product_version: ProductVersion = version.parse().ok()?;

                let path = value.path();
                let path = path.join("program");

                // Not a valid office install s
                if !path.exists() {
                    return None;
                }

                Some((product_version, path))
            })
            .collect();

        Ok(installs)
    }

    /// Finds the latest LibreOffice installation from the `/opt` directory
    pub fn find_opt_latest() -> std::io::Result<Option<PathBuf>> {
        // Find all libreoffice folders
        let mut installs: Vec<(ProductVersion, PathBuf)> = Self::find_opt_installs()?;

        // Sort to find the latest installed version
        installs.sort_by_key(|(key, _)| *key);

        // Last item will be the latest
        let latest = installs
            .pop()
            // Only use the path portion
            .map(|(_, path)| path);

        Ok(latest)
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

    /// Obtain the document type
    pub fn get_document_type(&mut self) -> Result<DocumentType, OfficeError> {
        let result = unsafe { self.raw.get_document_type()? };
        Ok(DocumentType::from_primitive(result))
    }
}

/// Filter types supported by office
#[derive(Debug, Deserialize)]
pub struct FilterTypes {
    /// Mapping between the filter name and details
    #[serde(flatten)]
    pub values: HashMap<String, FilterType>,
}

impl FilterTypes {
    /// Get the filter type name by mime type
    pub fn get_by_mime(&self, mime: &str) -> Option<&str> {
        self.values
            .iter()
            // Find filter with matching media type
            .find(|(_, value)| value.media_type.eq(mime))
            // Map to only include the key
            .map(|(key, _)| key.as_str())
    }

    /// Checks if the provided mime type is supported for a filter type
    pub fn is_mime_supported(&self, mime: &str) -> bool {
        self.get_by_mime(mime).is_some()
    }

    /// Gets a list of the supported filter mime types
    pub fn supported_mime_types(&self) -> Vec<&str> {
        self.values
            .values()
            .map(|value| value.media_type.as_str())
            .collect()
    }
}

#[derive(Debug, Deserialize)]
pub struct FilterType {
    /// Mime type of the filter format (i.e application/pdf)
    #[serde(rename = "MediaType")]
    pub media_type: String,
}

#[derive(Debug, Deserialize)]
pub struct OfficeVersionInfo {
    #[serde(rename = "ProductName")]
    pub product_name: String,
    #[serde(rename = "ProductVersion")]
    pub product_version: ProductVersion,
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

#[derive(Debug, FromPrimitive, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub enum DocumentType {
    Text = 0,
    Spreadsheet = 1,
    Presentation = 2,
    Drawing = 3,
    #[num_enum(catch_all)]
    Other(i32),
}

#[derive(Debug, PartialEq, Eq, Clone, Copy)]
pub struct ProductVersion {
    pub major: u32,
    pub minor: u32,
}

impl ProductVersion {
    const MIN_SUPPORTED_VERSION: ProductVersion = ProductVersion::new(4, 3);
    const VERSION_6_0: ProductVersion = ProductVersion::new(6, 0);

    pub const fn new(major: u32, minor: u32) -> Self {
        Self { major, minor }
    }

    /// documentLoad requires libreoffice >=4.3
    pub fn is_document_load_available(&self) -> bool {
        self.ge(&Self::MIN_SUPPORTED_VERSION)
    }

    /// documentLoad requires libreoffice >=5.0
    pub fn is_document_load_options_available(&self) -> bool {
        self.ge(&ProductVersion::new(5, 0))
    }

    /// freeError requires libreoffice >=5.2
    pub fn is_free_error_available(&self) -> bool {
        self.ge(&ProductVersion::new(5, 2))
    }

    /// registerCallback requires libreoffice >=6.0
    pub fn is_register_callback_available(&self) -> bool {
        self.ge(&Self::VERSION_6_0)
    }

    /// getFilterTypes requires libreoffice >=6.0
    pub fn is_filter_types_available(&self) -> bool {
        self.ge(&Self::VERSION_6_0)
    }

    /// setOptionalFeatures requires libreoffice >=6.0
    pub fn is_optional_features_available(&self) -> bool {
        self.ge(&Self::VERSION_6_0)
    }

    /// setDocumentPassword requires libreoffice >=6.0
    pub fn is_set_document_password_available(&self) -> bool {
        self.ge(&Self::VERSION_6_0)
    }

    /// getVersionInfo requires libreoffice >=6.0
    pub fn is_get_version_info_available(&self) -> bool {
        self.ge(&Self::VERSION_6_0)
    }

    /// runMacro requires libreoffice >=6.0
    pub fn is_run_macro_available(&self) -> bool {
        self.ge(&Self::VERSION_6_0)
    }

    /// trimMemory requires libreoffice >=7.6
    pub fn is_trim_memory_available(&self) -> bool {
        self.ge(&ProductVersion::new(7, 6))
    }
}

impl PartialOrd for ProductVersion {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for ProductVersion {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        match self.major.cmp(&other.major) {
            // Ignore equal major versions
            core::cmp::Ordering::Equal => {}
            ord => return ord,
        }

        // Check minor versions
        self.minor.cmp(&other.minor)
    }
}

impl Display for ProductVersion {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}.{}", self.major, self.minor)
    }
}

#[derive(Debug, Error)]
#[error("product version is invalid or malformed")]
pub struct InvalidProductVersion;

impl FromStr for ProductVersion {
    type Err = InvalidProductVersion;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let (major, minor) = s.split_once('.').ok_or(InvalidProductVersion)?;

        let major = major.parse().map_err(|_| InvalidProductVersion)?;
        let minor = minor.parse().map_err(|_| InvalidProductVersion)?;

        Ok(Self { major, minor })
    }
}

impl<'de> Deserialize<'de> for ProductVersion {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        let value: &str = <&str>::deserialize(deserializer)?;

        value
            .parse::<ProductVersion>()
            .map_err(|err| serde::de::Error::custom(err.to_string()))
    }
}

impl Serialize for ProductVersion {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        self.to_string().serialize(serializer)
    }
}
