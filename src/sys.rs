use std::{
    ffi::{CStr, CString},
    os::raw::{c_char, c_int, c_ulonglong, c_void},
    ptr::null_mut,
    sync::atomic::{AtomicBool, Ordering},
};

use ffi::{LibreOfficeKit, LibreOfficeKitClass, LibreOfficeKitDocument};
use parking_lot::Mutex;

use crate::{error::OfficeError, urls::DocUrl};

// Include the bindings
#[allow(
    dead_code,
    non_snake_case,
    non_camel_case_types,
    non_upper_case_globals
)]
#[allow(clippy::all)]
mod ffi {
    include!(concat!(env!("OUT_DIR"), "/bindings.rs"));
}

/// Global lock to prevent creating multiple office instances
/// at one time, all other instances must be dropped before
/// a new one can be created
pub(crate) static GLOBAL_OFFICE_LOCK: AtomicBool = AtomicBool::new(false);

/// Type used for the callback data
pub type CallbackData = *mut Box<dyn FnMut(c_int, *const c_char)>;

/// Raw office pointer access
pub struct OfficeRaw {
    /// This pointer for LOK
    this: *mut LibreOfficeKit,
    /// Class pointer for LOK
    class: *mut LibreOfficeKitClass,
    /// Callback data if specified
    callback_data: Mutex<CallbackData>,
}

impl OfficeRaw {
    /// Initializes a new instance of LOK
    pub unsafe fn init(install_path: *const c_char) -> Result<Self, OfficeError> {
        let lok = ffi::lok_init_wrapper(install_path);

        if lok.is_null() {
            return Err(OfficeError::UnknownInit);
        }

        let lok_class = (*lok).pClass;

        let instance = Self {
            this: lok,
            class: lok_class,
            callback_data: Mutex::new(null_mut()),
        };

        Ok(instance)
    }

    /// Gets a [CString] containing the JSON for the available LibreOffice filter types
    pub unsafe fn get_filter_types(&self) -> Result<CString, OfficeError> {
        let get_filter_types = (*self.class)
            .getFilterTypes
            .ok_or(OfficeError::MissingFunction("getFilterTypes"))?;

        let value = get_filter_types(self.this);

        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(CString::from_raw(value))
    }

    /// Gets a [CString] containing the JSON for the current LibreOffice version details
    pub unsafe fn get_version_info(&self) -> Result<CString, OfficeError> {
        let get_version_info = (*self.class)
            .getVersionInfo
            .ok_or(OfficeError::MissingFunction("getVersionInfo"))?;

        let value = get_version_info(self.this);

        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(CString::from_raw(value))
    }

    /// Gets a [CString] containing a dump of the current LibreOffice state
    pub unsafe fn dump_state(&self) -> Result<CString, OfficeError> {
        let mut state: *mut c_char = null_mut();
        let dump_state = (*self.class)
            .dumpState
            .ok_or(OfficeError::MissingFunction("dumpState"))?;
        dump_state(self.this, std::ptr::null(), &mut state);

        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(CString::from_raw(state))
    }

    /// Trims memory from LibreOffice
    pub unsafe fn trim_memory(&self, target: c_int) -> Result<(), OfficeError> {
        let trim_memory = (*self.class)
            .trimMemory
            .ok_or(OfficeError::MissingFunction("trimMemory"))?;
        trim_memory(self.this, target);

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(())
    }

    /// Sets an office option
    pub unsafe fn set_option(
        &self,
        option: *const c_char,
        value: *const c_char,
    ) -> Result<(), OfficeError> {
        let set_option = (*self.class)
            .setOption
            .ok_or(OfficeError::MissingFunction("setOption"))?;
        set_option(self.this, option, value);

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(())
    }

    /// Exports the provided document and signs the content
    pub unsafe fn sign_document(
        &self,
        url: &DocUrl,
        certificate: *const u8,
        certificate_len: i32,
        private_key: *const u8,
        private_key_len: i32,
    ) -> Result<bool, OfficeError> {
        let sign_document = (*self.class)
            .signDocument
            .ok_or(OfficeError::MissingFunction("signDocument"))?;
        let result = sign_document(
            self.this,
            url.as_ptr(),
            certificate,
            certificate_len,
            private_key,
            private_key_len,
        );

        Ok(result)
    }

    /// Loads a document without any options
    pub unsafe fn document_load(&self, url: &DocUrl) -> Result<DocumentRaw, OfficeError> {
        let document_load = (*self.class)
            .documentLoad
            .ok_or(OfficeError::MissingFunction("documentLoad"))?;
        let this = document_load(self.this, url.as_ptr());

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        debug_assert!(!this.is_null());

        Ok(DocumentRaw { this })
    }

    /// Loads a document with additional options
    pub unsafe fn document_load_with_options(
        &self,
        url: &DocUrl,
        options: *const c_char,
    ) -> Result<DocumentRaw, OfficeError> {
        let document_load_with_options = (*self.class)
            .documentLoadWithOptions
            .ok_or(OfficeError::MissingFunction("documentLoadWithOptions"))?;
        let this = document_load_with_options(self.this, url.as_ptr(), options);

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        debug_assert!(!this.is_null());

        Ok(DocumentRaw { this })
    }

    /// Sets the current document password
    ///
    /// Can ONLY be used in [OfficeRaw::register_callback] when used outside
    /// a callback LOK will throw an error
    pub unsafe fn set_document_password(
        &self,
        url: &DocUrl,
        password: *const c_char,
    ) -> Result<(), OfficeError> {
        let set_document_password = (*self.class)
            .setDocumentPassword
            .ok_or(OfficeError::MissingFunction("setDocumentPassword"))?;

        set_document_password(self.this, url.as_ptr(), password);

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(())
    }

    /// Sets the optional features bitset
    pub unsafe fn set_optional_features(&self, features: u64) -> Result<(), OfficeError> {
        let set_optional_features = (*self.class)
            .setOptionalFeatures
            .ok_or(OfficeError::MissingFunction("setOptionalFeatures"))?;
        set_optional_features(self.this, features);

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(())
    }

    pub unsafe fn send_dialog_event(
        &self,
        window_id: c_ulonglong,
        arguments: *const c_char,
    ) -> Result<(), OfficeError> {
        let send_dialog_event = (*self.class)
            .sendDialogEvent
            .ok_or(OfficeError::MissingFunction("sendDialogEvent"))?;

        send_dialog_event(self.this, window_id, arguments);

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        Ok(())
    }

    pub unsafe fn run_macro(&self, url: *const c_char) -> Result<bool, OfficeError> {
        let run_macro = (*self.class)
            .runMacro
            .ok_or(OfficeError::MissingFunction("runMacro"))?;

        let result = run_macro(self.this, url);

        if result == 0 {
            // Check for errors
            if let Some(error) = self.get_error() {
                return Err(OfficeError::OfficeError(error));
            }
        }

        Ok(result != 0)
    }

    /// Clears the currently registered callback
    pub unsafe fn clear_callback(&self) -> Result<(), OfficeError> {
        let register_callback = (*self.class)
            .registerCallback
            .ok_or(OfficeError::MissingFunction("registerCallback"))?;

        register_callback(self.this, None, null_mut());

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        self.free_callback();

        Ok(())
    }

    pub unsafe fn register_callback<F>(&self, callback: F) -> Result<(), OfficeError>
    where
        F: FnMut(c_int, *const c_char) + 'static,
    {
        /// Create a shim to wrap the callback function so it can be invoked
        unsafe extern "C" fn callback_shim(ty: c_int, payload: *const c_char, data: *mut c_void) {
            // Get the callback function from the data argument
            let callback: *mut Box<dyn FnMut(c_int, *const c_char)> = data.cast();

            // Catch panics from calling the callback
            _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(move || {
                // Invoke the callback
                (**callback)(ty, payload);
            }));
        }

        // Callback is double boxed then leaked
        let callback_ptr: *mut Box<dyn FnMut(c_int, *const c_char)> =
            Box::into_raw(Box::new(Box::new(callback)));

        let register_callback = (*self.class)
            .registerCallback
            .ok_or(OfficeError::MissingFunction("registerCallback"))?;

        register_callback(self.this, Some(callback_shim), callback_ptr.cast());

        // Check for errors
        if let Some(error) = self.get_error() {
            return Err(OfficeError::OfficeError(error));
        }

        // Free any existing callbacks
        self.free_callback();

        // Store the new callback
        *self.callback_data.lock() = callback_ptr;

        Ok(())
    }

    /// Frees the current allocated callback data memory if
    /// a callback has been set
    unsafe fn free_callback(&self) {
        let callback = &mut *self.callback_data.lock();

        // Callback has not been set
        if callback.is_null() {
            return;
        }

        let mut callback_ptr: CallbackData = null_mut();

        // Obtain the callback pointer
        std::mem::swap(callback, &mut callback_ptr);

        // Reclaim the raw memory
        _ = Box::from_raw(callback_ptr);
    }

    /// Requests the latest error from LOK if one is available
    pub unsafe fn get_error(&self) -> Option<String> {
        let get_error = (*self.class).getError.expect("missing getError function");
        let raw_error = get_error(self.this);

        // Empty error is considered to be no error
        if *raw_error == 0 {
            return None;
        }

        // Create rust copy of the error message
        let value = CStr::from_ptr(raw_error).to_string_lossy().into_owned();

        // Free error memory
        self.free_error(raw_error);

        Some(value)
    }

    /// Frees the memory allocated for an error by LOK
    ///
    /// Used when we've obtained the error as we clone
    /// our own copy of the error
    unsafe fn free_error(&self, error: *mut i8) {
        // Only available LibreOffice >=5.2
        if let Some(free_error) = (*self.class).freeError {
            free_error(error);
        }
    }

    /// Destroys the LOK instance and frees any other
    /// allocated memory
    pub unsafe fn destroy(&self) {
        let destroy = (*self.class).destroy.expect("missing destroy function");
        destroy(self.this);

        // Free the callback if allocated
        self.free_callback();
    }
}

impl Drop for OfficeRaw {
    fn drop(&mut self) {
        #[cfg(feature = "destroy_on_drop")]
        unsafe {
            self.destroy()
        }

        // Unlock the global office lock
        GLOBAL_OFFICE_LOCK.store(false, Ordering::SeqCst)
    }
}

pub struct DocumentRaw {
    /// This pointer for the document
    this: *mut LibreOfficeKitDocument,
}

impl DocumentRaw {
    /// Saves the document as another format
    pub unsafe fn save_as(
        &mut self,
        url: &DocUrl,
        format: *const c_char,
        filter: *const c_char,
    ) -> Result<i32, OfficeError> {
        let class = (*self.this).pClass;
        let save_as = (*class)
            .saveAs
            .ok_or(OfficeError::MissingFunction("saveAs"))?;

        Ok(save_as(self.this, url.as_ptr(), format, filter))
    }

    /// Get the type of document
    pub unsafe fn get_document_type(&mut self) -> Result<i32, OfficeError> {
        let class = (*self.this).pClass;
        let get_document_type = (*class)
            .getDocumentType
            .ok_or(OfficeError::MissingFunction("getDocumentType"))?;

        Ok(get_document_type(self.this))
    }

    pub unsafe fn destroy(&mut self) {
        let class = (*self.this).pClass;
        let destroy = (*class).destroy.expect("missing destroy function");
        destroy(self.this);
    }
}

impl Drop for DocumentRaw {
    fn drop(&mut self) {
        #[cfg(feature = "destroy_on_drop")]
        unsafe {
            self.destroy()
        }
    }
}
