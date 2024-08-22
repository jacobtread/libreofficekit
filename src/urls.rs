use std::ffi::{c_char, CString};
use std::fmt;
use std::path::Path;
use url::Url;

use crate::error::OfficeError;

/// Type-safe URL "container" for LibreOffice documents
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocUrl(CString);

impl DocUrl {
    /// Internal use only, obtains a pointer to the string value
    pub(crate) fn as_ptr(&self) -> *const c_char {
        self.0.as_ptr()
    }

    /// Converts a local relative path into an absolute path that LibreOffice can use
    pub fn local_into_abs<S: AsRef<str>>(path: S) -> Result<DocUrl, OfficeError> {
        let path: &str = path.as_ref();
        let abs_path = std::fs::canonicalize(path).map_err(|err| {
            OfficeError::OfficeError(format!("Does the file exist at {}? {}", path, err))
        })?;

        Self::local_as_abs(abs_path.display().to_string())
    }

    /// Converts a local absolute path into a [DocUrl] the path MUST be an absolute path
    /// otherwise you'll get an error from LibreOffice
    pub fn local_as_abs<S: AsRef<str>>(path: S) -> Result<DocUrl, OfficeError> {
        let value = path.as_ref();
        let path = Path::new(value);

        if !path.is_absolute() {
            return Err(OfficeError::OfficeError(format!(
                "The file path {} must be absolute!",
                &value
            )));
        }

        let url_value = Url::from_file_path(value)
            .map_err(|_| OfficeError::OfficeError(format!("failed to parse url {}", value)))?;

        let value_str = CString::new(url_value.as_str())?;
        Ok(DocUrl(value_str))
    }

    /// Converts a remote URL path into a [DocUrl]
    pub fn remote<S: AsRef<str>>(uri: S) -> Result<DocUrl, OfficeError> {
        let uri: &str = uri.as_ref();

        // Ensure the URL is valid
        _ = Url::parse(uri)
            .map_err(|_| OfficeError::OfficeError(format!("failed to parse uri {}", uri)))?;

        let value_str = CString::new(uri)?;

        Ok(DocUrl(value_str))
    }
}

impl fmt::Display for DocUrl {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.0.to_string_lossy())
    }
}
