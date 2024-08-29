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
    pub fn from_relative_path<S: AsRef<str>>(path: S) -> Result<DocUrl, OfficeError> {
        let path: &str = path.as_ref();
        let abs_path = std::fs::canonicalize(path).map_err(|err| {
            OfficeError::OfficeError(format!("Does the file exist at {}? {}", path, err))
        })?;

        Self::from_absolute_path(abs_path.display().to_string())
    }

    /// Converts a local absolute path into a [DocUrl] the path MUST be an absolute path
    /// otherwise you'll get an error from LibreOffice
    ///
    /// Path MUST be an absolute path, you'll receive an error if is not
    pub fn from_absolute_path<S: AsRef<str>>(path: S) -> Result<DocUrl, OfficeError> {
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

    /// Converts a remote URI path into a [DocUrl]
    pub fn from_remote_uri<S: AsRef<str>>(uri: S) -> Result<DocUrl, OfficeError> {
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

#[cfg(test)]
mod test {
    use super::DocUrl;

    /// Tests a valid relative path
    #[test]
    fn test_relative() {
        let path = "./src";
        let _url = DocUrl::from_relative_path(path).unwrap();
    }

    /// Tests an invalid relative path
    #[test]
    fn test_invalid_relative() {
        let path = "/__ABSOLUTE_PATH_THAT_DOES_NOT_EXIST__";
        let _url = DocUrl::from_relative_path(path).unwrap_err();
    }

    /// Tests a valid absolute URL
    #[test]
    fn test_absolute() {
        let path = "/src";
        let _url = DocUrl::from_absolute_path(path).unwrap();
    }

    /// Tests a non absolute URL fails the checks
    #[test]
    fn test_invalid_absolute() {
        let path = "./src";
        let _url = DocUrl::from_absolute_path(path).unwrap_err();
    }

    /// Tests a remote URL path
    #[test]
    fn test_remote_path() {
        let path = "http://localhost:5555/file.docx";
        let _url = DocUrl::from_remote_uri(path).unwrap();

        let path = "file://file.docx";
        let _url = DocUrl::from_remote_uri(path).unwrap();
    }

    /// Tests an invalid remote paths
    #[test]
    fn test_invalid_remote_path() {
        let path = "h/malformed";
        let _url = DocUrl::from_remote_uri(path).unwrap_err();

        let path = "h";
        let _url = DocUrl::from_remote_uri(path).unwrap_err();
    }
}
