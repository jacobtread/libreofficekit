use std::{ffi::NulError, str::Utf8Error};

use thiserror::Error;

#[derive(Debug, Error)]
pub enum OfficeError {
    /// The library files did not exist
    #[error("library not found")]
    MissingLibrary,

    /// Library is missing required hook function
    #[error("library missing hook function")]
    MissingLibraryHook,

    /// Failed to load the underlying library
    #[error(transparent)]
    LoadLibrary(dlopen2::Error),

    /// Error message produced by office
    #[error("{0}")]
    OfficeError(String),

    /// Function is not available in the current office install
    #[error("missing '{0}' function")]
    MissingFunction(&'static str),

    /// Filter types could not be parsed
    #[error("failed to parse filters: {0}")]
    InvalidFilterTypes(serde_json::Error),

    /// Version info could not be parsed
    #[error("failed to parse version info: {0}")]
    InvalidVersionInfo(serde_json::Error),

    /// Version info or filter types contained invalid UTF-8
    #[error("invalid utf8 string: {0}")]
    InvalidUtf8String(#[from] Utf8Error),

    /// Value with null
    #[error("string cannot contain null byte")]
    InvalidString(#[from] NulError),

    /// Provided path was invalid
    #[error("invalid path provided")]
    InvalidPath,

    /// Prevented from creating another office instance
    #[error("already another active instance")]
    InstanceLock,

    /// Office instance was dropped before a callback was invoked
    #[error("callback invoked after instance was dropped")]
    InstanceDropped,

    /// Unknown error happened while initializing LOK
    #[error("unknown initialization error")]
    UnknownInit,
}
