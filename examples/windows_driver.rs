use escpos::driver::windows_driver::{WindowsDriver, WindowsPrinter};
use escpos::driver::*;
use escpos::errors::Result;
use escpos::printer::Printer;
use escpos::utils::{DebugMode, Protocol};

fn main() -> Result<()> {
    env_logger::init();

    // List of Windows driver printers
    for device in windows_driver::WindowsPrinter::list_printers().unwrap().iter() {
        println!("{}, online: {}", device.get_name(), device.is_online());
    }

    let windows_printer = WindowsPrinter::from_str("POS-58")?;
    if !windows_printer.is_online() {
        println!("Printer is offline");
        return Ok(());
    }
    let driver = WindowsDriver::open(&windows_printer)?;
    Printer::new(driver, Protocol::default(), None)
        .debug_mode(Some(DebugMode::Dec))
        .init()?
        .writeln("Windows Driver test")?
        .print_cut()?;
    Ok(())
}
