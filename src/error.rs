use std::{ffi::NulError, str::Utf8Error};

use thiserror::Error;

#[derive(Debug, Error)]
pub enum OfficeError {
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

    /// Prevented from creating another office instance
    #[error("already another active instance")]
    InstanceLock,

    /// Unknown error happened while initializing LOK
    #[error("unknown initialization error")]
    UnknownInit,
}
