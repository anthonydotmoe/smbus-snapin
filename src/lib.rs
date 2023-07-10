use std::collections::HashMap;

#[macro_use]
extern crate guid;

use guid::GUID;

use intercom::{prelude::*, BString};
use intercom::interfaces::IUnknown;

use windows::Win32::Foundation::{MAX_PATH, GetLastError, NO_ERROR, HMODULE, COLORREF, SysAllocString};
use windows::Win32::Graphics::Gdi::{HBITMAP, CreateBitmap};
use windows::Win32::System::Com::{TYMED_HGLOBAL, TYMED_FILE, TYMED_ISTREAM, TYMED_ISTORAGE, TYMED_GDI, TYMED_MFPICT, TYMED_ENHMF, TYMED_NULL};
use windows::Win32::System::Memory::{GlobalLock, GlobalUnlock, GlobalSize};
use windows::core::PCWSTR;
use windows::Win32::UI::WindowsAndMessaging::{HICON, LoadIconW};
use windows::Win32::System::LibraryLoader::{GetModuleHandleW, GetModuleHandleExW, GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS};
use windows::Win32::System::DataExchange::GetClipboardFormatNameW;

mod interfaces;
pub mod mmc;

use interfaces::*;

use crate::mmc::{MmcNotifyType, MmcDataObjectType};

const FAN_NODE_NODE_TYPE:  GUID = guid!{"977821db-5a85-4130-91c8-f0ad8636608e"};
const ROOT_NODE_NODE_TYPE: GUID = guid!{"f4475e15-6df7-4c72-af06-b7ef9004ec1f"};
const SNAPIN_CLSID:        GUID = guid!{"55ace359-cfbf-475f-a891-d1203894604c"};
const SNAPIN_ABOUT_CLSID:  GUID = guid!{"55ace359-cfbf-475f-a891-d1203894604d"};

com_library!(
    on_load = on_load,
    //on_register = register_snapin,
    //on_unregister = unregister_snapin,
    class MMCSnapIn,
    class MMCSnapInComponent,
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

fn guid_to_string(value: &GUID) -> String {
    format!(
        "{{{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
        value.Data1,
        value.Data2,
        value.Data3,
        value.Data4[0],
        value.Data4[1],
        value.Data4[2],
        value.Data4[3],
        value.Data4[4],
        value.Data4[5],
        value.Data4[6],
        value.Data4[7]
    )
}

#[cfg(test)]
mod tests {
    use crate::guid_to_string;

    #[test]
    fn test_guid_to_string() {
        let mut guid: guid::GUID = guid!{"00010203-0405-0607-0809-0A0B0C0D0E0F"};
        assert_eq!(guid_to_string(&guid), "{00010203-0405-0607-0809-0A0B0C0D0E0F}");
        guid = guid!{"00010203-0405-0607-0809-0a0b0c0d0e0f"};
        assert_eq!(guid_to_string(&guid), "{00010203-0405-0607-0809-0A0B0C0D0E0F}");
    }
}

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

#[com_class(clsid = "55ace359-cfbf-475f-a891-d1203894604c", IComponentData)]
#[derive(Default)]
struct MMCSnapIn {
    console: Option<ComRc<dyn IConsole>>,
    console_namespace: Option<ComRc<dyn IConsoleNamespace>>,
    nodes: HashMap<isize, ComRc<dyn IDataObject>>,
    next_cookie: isize,
}

impl IComponentData for MMCSnapIn {
    fn initialize(&mut self, lp_unknown: &ComItf<dyn IUnknown>) -> ComResult<()> {
        log::debug!("IComponentData for MMCSnapIn called");

        // Use the received IUnknown interface to query for IConsole2 and
        // IConsoleNamespace2

        let console: ComResult<ComRc<dyn IConsole>>  = ComItf::query_interface(lp_unknown);
        let console_namespace: ComResult<ComRc<dyn IConsoleNamespace>>  = ComItf::query_interface(lp_unknown);
        
        match console {
            Ok(console) => {
                log::debug!("Got: {:p} for IConsole", console.as_raw_iunknown());
                self.console = Some(console.clone());
            },
            Err(e) => { log::error!("Error {:?}: QI for IConsole", e); }
        }
        
        match console_namespace {
            Ok(console_namespace) => {
                log::debug!("Got {:p} for IConsoleNamespace", console_namespace.as_raw_iunknown());
                self.console_namespace = Some(console_namespace.clone())
            },
            Err(e) => { log::error!("Error {:?}: QI for IConsoleNamespace", e); }
        }
        
        // Create the root node
        let root_node = MMCSnapInNode { node_type: MMCSnapInNodeType::Root };
        self.nodes.insert(0, ComRc::from(ComBox::new(root_node)));
        self.next_cookie = 1;
        
        log::debug!("IComponentData::Initialize done");
        
        Ok(())
    }
    
    fn create_component(&self) -> ComResult<ComRc<dyn IComponent>> {
        let component = MMCSnapInComponent::new();
        Ok(ComRc::from(ComBox::new(component)))
    }
    
    fn notify(&self,lp_dataobject: &ComItf<dyn IDataObject>,event:u32,arg:i64,param:i64) -> ComResult<()> {
        let mmc_event: MmcNotifyType = unsafe { std::mem::transmute(event) };
        log::info!("Received event {:?}", mmc_event);
        Ok(())
    }
    
    fn destroy(&self) -> ComResult<()> {
        Ok(())
    }
    
    fn query_data_object(&mut self, cookie: isize, view_type: i32) -> ComResult<ComRc<dyn IDataObject>> {
        let view_type: MmcDataObjectType = unsafe { std::mem::transmute(view_type) };
    
        log::debug!("QueryDataObject: cookie: {}, ppviewtype: {:?}", cookie, view_type);
            
        match &cookie {
            0 => {
                match view_type {
                    MmcDataObjectType::SnapinManager => {
                        match self.nodes.get(&0) {
                            Some(data_object) => {
                                log::debug!("Found cookie in HashMap: {:?}", data_object);
                                return Ok(data_object.clone());
                            }
                            None => {
                                log::error!("Did not find cookie in HashMap!");
                                let root_node = ComRc::from(ComBox::new(MMCSnapInNode { node_type: MMCSnapInNodeType::Root }));
                                log::debug!("Created snapin manager scope root node: {:?}", root_node);
                                self.nodes.insert(0, root_node);
                                return Ok(self.nodes.get(&cookie).unwrap().clone());
                            }
                        }
                    }
                    MmcDataObjectType::Scope => {
                        match self.nodes.get(&1) {
                            Some(data_object) => {
                                log::debug!("Found cookie in HashMap: {:?}", data_object);
                                return Ok(data_object.clone());
                            }
                            None => {
                                log::error!("Did not find cookie in HashMap!");
                                let root_node = ComRc::from(ComBox::new(MMCSnapInNode { node_type: MMCSnapInNodeType::Root }));
                                log::debug!("Created temporary scope root node: {:?}", root_node);
                                self.nodes.insert(0, root_node);
                                return Ok(self.nodes.get(&cookie).unwrap().clone());
                            }
                        }
                    }
                    MmcDataObjectType::Result => {
                        match self.nodes.get(&2) {
                            Some(data_object) => {
                                log::debug!("Found cookie in HashMap: {:?}", data_object);
                                return Ok(data_object.clone());
                            }
                            None => {
                                log::error!("Did not find cookie in HashMap!");
                                let root_node = ComRc::from(ComBox::new(MMCSnapInNode { node_type: MMCSnapInNodeType::Root }));
                                log::debug!("Created temporary result root node: {:?}", root_node);
                                self.nodes.insert(0, root_node);
                                return Ok(self.nodes.get(&cookie).unwrap().clone());
                            }
                        }
                    }
                    _ => {
                        return Err(ComError::E_POINTER);
                    }
                }
            }
            _ => {
                Err(ComError::E_POINTER)
            }
        }
    }
    
    /*
    fn query_data_object(&mut self, cookie: isize, view_type: i32) -> ComResult<ComRc<dyn IDataObject>> {
        let view_type: MmcDataObjectType = unsafe { std::mem::transmute(view_type) };

        log::debug!("QueryDataObject: cookie: {}, ppviewtype: {:?}", cookie, view_type);
        
        match self.nodes.contains_key(&cookie) {
            true => {
                Ok(ComRc::from(ComBox::new((self.nodes[&cookie]))))
            }
            false => {
                Err(ComError::E_POINTER)
            }
        }
    }
    */
    
    fn get_display_info(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn compare_objects(&self,) -> ComResult<()> {
        Ok(())
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

#[com_class(IDataObject)]
#[derive(Default)]
struct MMCSnapInNode {
    node_type: MMCSnapInNodeType,
}

impl MMCSnapInNode {
    pub fn new() -> Self {
        MMCSnapInNode::default()
    }
}

impl IDataObject for MMCSnapInNode {
    fn get_data(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn get_data_here(&self,pformatetc: *const ComFORMATETC, pmedium: *mut ComSTGMEDIUM) -> ComResult<()> {
        let clipformat: u32 = unsafe { (*pformatetc).0.cfFormat.into() };
        log::debug!("Got Clipformat: {}", clipformat);
        
        let mut clipformat_name: [u16; MAX_PATH as usize] = [0; MAX_PATH as usize];
        
        let clipformat_name_len: i32 = unsafe { GetClipboardFormatNameW(clipformat, &mut clipformat_name) };
        
        if clipformat_name_len > 0 {
            let name = String::from_utf16_lossy(&clipformat_name[0..clipformat_name_len as usize]);
            let name_str = name.as_str();
            log::debug!("Got clipformat name: {}", name);
            
            match name_str {
                "CCF_DISPLAY_NAME" => {
                    // Return node display name as WSTR
                    unsafe {
                        match (*pmedium).0.tymed {
                            TYMED_HGLOBAL => {
                                let ptr = GlobalLock((*pmedium).0.Anonymous.hGlobal);
                                
                                if ptr.is_null() {
                                    log::error!("HGLOBAL returns null pointer");
                                    return Err(ComError::E_FAIL);
                                }
                                
                                let mut node_name_utf16: Vec<u16> = "Root Node for MMCSnapIn".encode_utf16().collect();
                                node_name_utf16.push(0);
                                
                                let node_name_utf16_size = node_name_utf16.len() * std::mem::size_of::<u16>();
                                
                                std::ptr::copy_nonoverlapping(node_name_utf16.as_ptr(), ptr as *mut u16, node_name_utf16_size);
                                
                                GlobalUnlock((*pmedium).0.Anonymous.hGlobal);
                                
                                if GetLastError() != NO_ERROR {
                                    log::error!("Failed to unlock HGLOBAL");
                                    return Err(ComError::E_FAIL);
                                }
                                
                            }
                            _ => {
                                log::error!("Unsupported TYMED: {:?}", (*pmedium).0.tymed);
                                return Err(mmc::DV_E_TYMED);
                            }
                        }
                    }
                }
                "CCF_NODETYPE" => {
                    // Return the GUID of the node type
                    unsafe {
                        match (*pmedium).0.tymed {
                            TYMED_HGLOBAL => {
                                let ptr = GlobalLock((*pmedium).0.Anonymous.hGlobal);
                                
                                if ptr.is_null() {
                                    log::error!("HGLOBAL returns null pointer");
                                    return Err(ComError::E_FAIL);
                                }
                                
                                let hglobal_size = GlobalSize((*pmedium).0.Anonymous.hGlobal);
                                log::debug!("HGLOBAL is {} bytes", hglobal_size);
                                log::debug!("Writing {:?} to HGLOBAL", ROOT_NODE_NODE_TYPE);
                                
                                let guid_size_in_bytes = std::mem::size_of::<GUID>();
                                
                                std::ptr::copy_nonoverlapping(&ROOT_NODE_NODE_TYPE as *const _ as *const u8, ptr as *mut u8, guid_size_in_bytes);

                                match GlobalUnlock((*pmedium).0.Anonymous.hGlobal).0 {
                                    0 => {
                                        let global_error = GetLastError();
                                        if global_error != NO_ERROR {
                                            log::error!("Failed to unlock HGLOBAL: {:?}", global_error);
                                            return Err(ComError::E_FAIL);
                                        }
                                    }
                                    _ => {
                                        log::debug!("Unlocking HGLOBAL seemed to work");
                                    }

                                }
                            }
                            _ => {
                                log::error!("Unsupported TYMED: {:?}", (*pmedium).0.tymed);
                                return Err(mmc::DV_E_TYMED);
                                
                            }
                        }
                    }
                }
                "CCF_SNAPIN_CLSID" => {
                    // Return the CLSID of the snap in
                    unsafe {
                        match (*pmedium).0.tymed {
                            TYMED_HGLOBAL => {
                                let ptr = GlobalLock((*pmedium).0.Anonymous.hGlobal);
                                
                                if ptr.is_null() {
                                    log::error!("HGLOBAL returns null pointer");
                                    return Err(ComError::E_FAIL);
                                }
                                
                                let hglobal_size = GlobalSize((*pmedium).0.Anonymous.hGlobal);
                                log::debug!("HGLOBAL is {} bytes", hglobal_size);
                                log::debug!("Writing {:?} to HGLOBAL", ROOT_NODE_NODE_TYPE);
                                
                                let guid_size_in_bytes = std::mem::size_of::<GUID>();
                                

                                std::ptr::copy_nonoverlapping(&SNAPIN_CLSID as *const _ as *const u8, ptr as *mut u8, guid_size_in_bytes);

                                match GlobalUnlock((*pmedium).0.Anonymous.hGlobal).0 {
                                    0 => {
                                        let global_error = GetLastError();
                                        if global_error != NO_ERROR {
                                            log::error!("Failed to unlock HGLOBAL: {:?}", global_error);
                                            return Err(ComError::E_FAIL);
                                        }
                                    }
                                    _ => {
                                        log::debug!("Unlocking HGLOBAL seemed to work");
                                    }
                                        
                                }
                            }
                            _ => {
                                log::error!("Unsupported TYMED: {:?}", (*pmedium).0.tymed);
                                return Err(mmc::DV_E_TYMED);
                                
                            }
                        }
                    }
                    
                }
                _ => {
                    log::error!("Can't handle this type of clipboard format. tymed: {:?} btw",
                        unsafe { (*pmedium).0.tymed }
                    );
                    return Err(mmc::DV_E_FORMATETC);
                }
            }
        } else {
            log::error!("Couldn't get clipboard format name. Dieing I guess");
            return Err(ComError::E_FAIL);
        }

        Ok(())
    }
    
    fn query_get_data(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn get_canonical_format(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn set_data(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn enum_format_etc(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn d_advise(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn d_unadvise(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn enum_d_advise(&self) -> ComResult<()> {
        Ok(())
    }
}

#[com_class(clsid = "55ace359-cfbf-475f-a891-d1203894604d", ISnapinAbout)]
#[derive(Default)]
struct MMCSnapInAbout {
    /*
    small_bitmap: std::cell::Cell<Option<HBITMAP>>,
    large_bitmap: std::cell::Cell<Option<HBITMAP>>,
    colorref: std::cell::Cell<Option<COLORREF>>,
    */
}

impl ISnapinAbout for MMCSnapInAbout {
    fn get_snapin_description(&self) -> ComResult<BString> {
        Err(ComError{ hresult: intercom::raw::HRESULT{ hr: 1 } , error_info: None })
    }

    fn get_provider(&self) -> ComResult<BString> {
        // let prov = BString::from("Anthony Guerrero");
        // Ok(prov)
        Err(ComError{ hresult: intercom::raw::HRESULT{ hr: 1 } , error_info: None })
    }
    
    fn get_snapin_version(&self) -> ComResult<BString> {
        // let ver = BString::from("0.6.9");
        // Ok(ver)
        Err(ComError{ hresult: intercom::raw::HRESULT{ hr: 1 } , error_info: None })
    }
    
    fn get_snapin_image(&self) -> ComResult<ComHICON> {
        /*
        let icon = unsafe {
            let hinstance = get_dll_hinstance();
        
            log::debug!("Got an HINSTANCE: {:?}", hinstance);

            LoadIconW(
                hinstance,
                make_int_resource(101)
            )
        };
        
        match icon {
            Ok(icon) => {
                log::debug!("Got an HICON: {:?}", icon);
                Ok(ComHICON(icon))
            },
            Err(e) => {
                log::error!("{}", e.to_string());
                Err(ComError::E_FAIL)
            }
        }
        */
        Err(ComError{ hresult: intercom::raw::HRESULT{ hr: 1 } , error_info: None })
    }
    
    fn get_static_folder_image(&self) -> ComResult<(ComHBITMAP,ComHBITMAP,ComHBITMAP,ComCOLORREF)> {
        /*
        self.small_bitmap.set(
            Some(unsafe{CreateBitmap(
                16,
                16,
                1,
                32,
                Some(render_circle_image(16).as_ptr() as *const _)
            )})
        );

        self.large_bitmap.set(
            Some(unsafe{CreateBitmap(
                32,
                32,
                1,
                32,
                Some(render_circle_image(32).as_ptr() as *const _)
            )})
        );
        
        self.colorref.set(Some(COLORREF(0xDEADBEEF)));
        
        Ok((
            ComHBITMAP(self.small_bitmap.get().unwrap()),
            ComHBITMAP(self.small_bitmap.get().unwrap()),
            ComHBITMAP(self.large_bitmap.get().unwrap()),
            ComCOLORREF(self.colorref.get().unwrap())
        ))
            */
        Err(ComError{ hresult: intercom::raw::HRESULT{ hr: 1 } , error_info: None })
    }
}

/*
fn render_circle_image(size: usize) -> Vec<u32> {
    // Render a faded circle as an example thumbnail.
    let mut data = Vec::new();
    data.resize((size * size) as usize, 0);
    let midpoint = size as f64 / 2.0;
    for x in 0..size {
        for y in 0..size {
            let x_coord = x as f64;
            let y_coord = y as f64;
            let dist_squared = (midpoint - x_coord).powf(2.0) + (midpoint - y_coord).powf(2.0);
            let mut value = (dist_squared.sqrt() / midpoint) * 255.0;
            if value > 255.0 {
                value = 255.0
            }
            let value = value as u32;
            data[(x * size + y) as usize] = value + (value << 8) + (value << 16) + (value << 24);
        }
    }
    data
}

fn make_int_resource(resource_id: u16) -> PCWSTR {
    PCWSTR(resource_id as _)
}
*/