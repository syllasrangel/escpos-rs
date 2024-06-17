use std::{cell::RefCell, ffi::c_void, rc::Rc};

pub use self::windows_printer::WindowsPrinter;
use crate::errors::{PrinterError, Result};
use windows::{
    core::{w, PWSTR},
    Win32::{
        Foundation::{BOOL, HANDLE},
        Graphics::Printing::{
            ClosePrinter, EndDocPrinter, EndPagePrinter, OpenPrinterW, StartDocPrinterW, StartPagePrinter,
            WritePrinter, DOC_INFO_1W,
        },
    },
};

use super::Driver;

mod windows_printer;

#[derive(Debug)]
pub struct WindowsDriver {
    printer_name: PWSTR,
    buffer: Rc<RefCell<Vec<u8>>>,
}

impl WindowsDriver {
    pub fn open(printer: &WindowsPrinter) -> Result<WindowsDriver> {
        Ok(Self {
            printer_name: printer.get_raw_name(),
            buffer: Rc::new(RefCell::new(Vec::new())),
        })
    }

    pub fn write_all(&self) -> Result<()> {
        let mut error: Option<PrinterError> = None;
        let mut printer_handle = HANDLE(0);
        let mut is_printer_open = false;
        let mut is_doc_started = false;
        let mut is_page_started = false;

        unsafe {
            // Open the printer
            if OpenPrinterW(self.printer_name, &mut printer_handle, None).is_err() {
                error = Some(PrinterError::Io("Failed to open printer".to_owned()));
                eprintln!("Error: {:?}", error);
            } else {
                is_printer_open = true;
                // Start the document
                let document_info = DOC_INFO_1W {
                    pDocName: PWSTR(w!("Raw Document").as_wide().as_ptr() as *mut _),
                    pOutputFile: PWSTR::null(),
                    pDatatype: PWSTR(w!("Raw").as_wide().as_ptr() as *mut _),
                };

                if StartDocPrinterW(printer_handle, 1, &document_info) == 0 {
                    error = Some(PrinterError::Io("Failed to start doc".to_owned()));
                    eprintln!("Error: {:?}", error);
                } else {
                    is_doc_started = true;
                    // Start the page
                    if !StartPagePrinter(printer_handle).as_bool() {
                        error = Some(PrinterError::Io("Failed to start page".to_owned()));
                        eprintln!("Error: {:?}", error);
                    } else {
                        is_page_started = true;
                        // Write to the printer
                        let buffer = self.buffer.borrow();
                        let mut written: u32 = 0;
                        if !WritePrinter(
                            printer_handle,
                            buffer.as_ptr() as *const c_void,
                            buffer.len() as u32,
                            &mut written,
                        )
                        .as_bool()
                        {
                            error = Some(PrinterError::Io("Failed to write to printer".to_owned()));
                            eprintln!("Error: {:?}", error);
                        } else if written != buffer.len() as u32 {
                            error = Some(PrinterError::Io("Failed to write all bytes to printer".to_owned()));
                            eprintln!("Error: {:?}", error);
                        }
                    }
                }
            }
        }

        // Clean up resources
        unsafe {
            if is_page_started {
                if EndPagePrinter(printer_handle) == BOOL(0) {
                    eprintln!("Warning: Failed to end page");
                }
            }
            if is_doc_started {
                if EndDocPrinter(printer_handle) == BOOL(0) {
                    eprintln!("Warning: Failed to end document");
                }
            }
            if is_printer_open {
                if let Err(e) = ClosePrinter(printer_handle) {
                    eprintln!("Warning: Failed to close printer: {:?}", e);
                }
            }
        }

        // Return result
        if let Some(err) = error {
            Err(err)
        } else {
            Ok(())
        }
    }
}

impl Driver for WindowsDriver {
    fn name(&self) -> String {
        "Windows Driver".to_owned()
    }

    fn write(&self, data: &[u8]) -> Result<()> {
        let mut buffer = self.buffer.borrow_mut();
        buffer.extend_from_slice(data);
        Ok(())
    }

    fn read(&self, _buf: &mut [u8]) -> Result<usize> {
        Ok(0)
    }

    fn flush(&self) -> Result<()> {
        self.write_all()
    }
}
