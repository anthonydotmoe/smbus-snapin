use intercom::prelude::*;

mod class;
use class::*;

mod interfaces;

com_library!(
    on_load = on_load,
    //on_register = register_snapin,
    //on_unregister = unregister_snapin,
    class MMCSnapInAbout,
    class MMCSnapIn,
    class Node,
);

fn on_load() {
    // Set up logging to project directory
    use log::LevelFilter;
    simple_logging::log_to_file(
        &format!("{}\\debug.log", env!("CARGO_MANIFEST_DIR")),
        LevelFilter::Debug,
    )
    .unwrap();
}

/*
fn register_snapin() -> intercom::raw::HRESULT {
    use std::path::Path;
    use winreg::enums::*;
    use winreg::RegKey;
    
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);
    let path = Path::new("SOFTWARE\\Microsoft\\MMC\\SnapIns").join(guid_to_string(&SNAPIN_CLSID));
    
    match hklm.create_subkey(&path) {
        Ok((key, _)) => {
            key.set_value("NameString", &"MMCSnapIn");
            key.set_value("About", &guid_to_string(&SNAPIN_ABOUT_CLSID));
            key.set_value("Version", &"1.0");
            
            key.create_subkey("StandAlone");
        }
        Err(e) => {
            panic!("Something happened: {}", e.to_string())
        }
    }
    
    intercom::raw::S_OK
}

fn unregister_snapin() -> intercom::raw::HRESULT {
    use std::path::Path;
    use winreg::enums::*;
    use winreg::RegKey;
    
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);
    let path = Path::new("SOFTWARE\\Microsoft\\MMC\\SnapIns").join(guid_to_string(&SNAPIN_CLSID));
    hklm.delete_subkey_all(path);
    
    intercom::raw::S_OK.into()
}

unsafe fn get_dll_hinstance() -> HMODULE {
    let mut hmodule: HMODULE = HMODULE::default();
    let address = get_dll_hinstance as *const _;
    
    let result = GetModuleHandleExW(
        GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT | GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
        PCWSTR::from_raw(address),
        &mut hmodule
    );
    
    match result.as_bool() {
        true => {
            hmodule
        }
        false => {
            log::error!("Wow, getting our HINSTANCE didn't work. Oh well, I guess.");
            hmodule
        }
    }
}


#[com_class(IComponent)]
#[derive(Default)]
struct MMCSnapInComponent {
    x: i32,
}

impl MMCSnapInComponent {
    pub fn new() -> Self {
        MMCSnapInComponent::default()
    }
}

impl IComponent for MMCSnapInComponent {
    fn initialize(&self, _lp_unknown: &ComItf<dyn IUnknown>) -> ComResult<()> {

        // Use IConsole interface to get handles to IResultData, IToolbar, etc.
        Ok(())
    }
    
    fn notify(&self, lp_dataobject: &ComItf<dyn IDataObject>, event:u32, arg:i64, param:i64) -> ComResult<()> {
        let mmc_event: MmcNotifyType = unsafe { std::mem::transmute(event) };
        log::info!("Received event {:?}, lp_dataobject: {:p}, arg: {}, param: {}", mmc_event, lp_dataobject, arg, param);
        Ok(())
        
    }
    
    fn destroy(&self) -> ComResult<()> {
        Ok(())
    }
    
    fn query_data_object(&mut self, _cookie:isize, _object_type:i32) -> ComResult<ComRc<dyn IDataObject>> {
        Err(ComError::E_POINTER)
    }
    
    fn get_result_view_type(&self, _cookie:isize) -> ComResult<(BString,u64)> {
        Err(ComError::E_NOTIMPL)
    }
    
    fn get_display_info(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn compare_objects(&self,) -> ComResult<()> {
        Ok(())
    }
}

#[derive(Debug)]
enum MMCSnapInNodeType {
    Root,
    Fan,
}

impl Default for MMCSnapInNodeType {
    fn default() -> Self {
        MMCSnapInNodeType::Fan
    }
}


*/