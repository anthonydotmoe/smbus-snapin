use intercom::{prelude::*, BString};
use intercom::interfaces::IUnknown;
use winapi::shared::windef::HICON;
use winapi::um::winuser::{LoadIconW, MAKEINTRESOURCEW};
use winapi::um::libloaderapi::GetModuleHandleW;

mod interfaces;
use interfaces::*;

// NOTE: NodeType GUID for "Fan" node: 977821db-5a85-4130-91c8-f0ad8636608e

com_library!(
    on_load=on_load,
    class MMCSnapIn,
    class MMCSnapInAbout,
);

fn on_load() {
    // Set up logging to project directory
    use log::LevelFilter;
    simple_logging::log_to_file(
        &format!("{}\\debug.log", env!("CARGO_MANIFEST_DIR")),
        LevelFilter::Trace,
    )
    .unwrap();
}

#[com_class(clsid = "55ace359-cfbf-475f-a891-d1203894604c", IComponentData)]
#[derive(Default)]
struct MMCSnapIn {
    console: Option<ComRc<dyn IConsole2>>,
    console_namespace: Option<ComRc<dyn IConsoleNamespace2>>,
}

impl IComponentData for MMCSnapIn {
    fn initialize(&mut self, lp_unknown: &ComItf<dyn IUnknown>) -> ComResult<i32> {
        // Use the received IUnknown interface to query for IConsole2 and
        // IConsoleNamespace2
        
        let console: ComResult<ComRc<dyn IConsole2>>  = ComItf::query_interface(lp_unknown);
        let console_namespace: ComResult<ComRc<dyn IConsoleNamespace2>>  = ComItf::query_interface(lp_unknown);
        
        match console {
            Ok(console) => self.console = Some(console.clone()),
            Err(_e) => { /* Do something! */}
        }
        
        match console_namespace {
            Ok(console_namespace) => self.console_namespace = Some(console_namespace.clone()),
            Err(_e) => { /* Do something! */}
        }
        
        Ok(0)
    }
    // ... rest of the methods
}

#[com_class(clsid = "55ace359-cfbf-475f-a891-d1203894604d", ISnapinAbout)]
#[derive(Default)]
struct MMCSnapInAbout {
}

impl ISnapinAbout for MMCSnapInAbout {
    fn get_snapin_description(&self) -> ComResult<BString> {
        let desc = BString::from("This is a snap-in");
        Ok(desc)
    }

    fn get_provider(&self) -> ComResult<BString> {
        let prov = BString::from("Anthony Guerrero");
        Ok(prov)
    }
    
    fn get_snapin_version(&self) -> ComResult<BString> {
        let ver = BString::from("0.6.9");
        Ok(ver)
    }
    
    fn get_snapin_image(&self) -> ComResult<ComHICON> {
        let icon: HICON = unsafe { LoadIconW(GetModuleHandleW(std::ptr::null()), MAKEINTRESOURCEW(1)) };
        Ok(ComHICON(icon))
    }
    
    fn get_static_folder_image(&self) -> ComResult<(ComHBITMAP,ComHBITMAP,ComHBITMAP,ComCOLORREF)> {
        Err(ComError { hresult: intercom::raw::HRESULT { hr: -1 }, error_info: None })
    }
}