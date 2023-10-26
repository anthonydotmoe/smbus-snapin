use guid::GUID;
use intercom::{prelude::*, raw::HRESULT};
use windows::{Win32::{System::{Memory::{ GlobalUnlock, GlobalLock, GlobalSize }, DataExchange::GetClipboardFormatNameW, Com::TYMED_HGLOBAL}, Foundation::{MAX_PATH, GetLastError, NO_ERROR}}, core::PCWSTR, w};

use crate::{interfaces::{IDataObject, ComFORMATETC, ComSTGMEDIUM}, class::snapin::CLSID_MMCSnapIn};

pub const DV_E_FORMATETC: ComError = ComError{
    hresult: HRESULT { hr: windows::Win32::Foundation::DV_E_FORMATETC.0 },
    error_info: None,
};

pub const DV_E_TYMED: ComError = ComError{
    hresult: HRESULT { hr: windows::Win32::Foundation::DV_E_TYMED.0 },
    error_info: None,
};

#[derive(Debug, Default)]
pub enum NodeType {
    #[default]
    Folder,
    Root,
}

#[com_class(IDataObject)]
#[derive(Debug)]
pub struct Node {
    pub node_type: NodeType,
    pub display_name: PCWSTR,
}

impl Default for Node {
    fn default() -> Self {
        Node {
            node_type: NodeType::Folder,
            display_name: w!("Default Node")
        }
    }
}

/*
impl Node {
    pub fn new() -> Self {
        Node {
            node_type: NodeType::Root,
            display_name: w!("Default GPFolders::Node Name"),
        }
    }
}
*/

impl IDataObject for Node {
    fn get_data(&self, _pformatetc: *const ComFORMATETC) -> ComResult<*mut ComSTGMEDIUM> {
        Err(ComError::E_NOTIMPL)
    }
    
    fn get_data_here(&self, pformatetc: *const ComFORMATETC, pmedium: *mut ComSTGMEDIUM) -> ComResult<()> {
        let clipformat: u32 = unsafe { (*pformatetc).0.cfFormat.into() };
        log::debug!("Got Clipformat: {}", clipformat);
        
        let mut clipformat_name: [u16; MAX_PATH as usize] = [0; MAX_PATH as usize];
        
        let clipformat_name_len: i32 = unsafe { GetClipboardFormatNameW(clipformat, &mut clipformat_name) };
        
        if clipformat_name_len > 0 {
            let name = String::from_utf16_lossy(&clipformat_name[0..clipformat_name_len as usize]);
            log::debug!("Got clipformat name: {}", name);
            
            match name.as_str() {
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
                                
                                let mut node_name_utf16: Vec<u16> = self.display_name.as_wide().to_vec();
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
                                return Err(DV_E_TYMED);
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
                                log::debug!("Writing {:?} to HGLOBAL", &CLSID_MMCSnapIn);
                                
                                let guid_size_in_bytes = std::mem::size_of::<GUID>();
                                
                                std::ptr::copy_nonoverlapping(&CLSID_MMCSnapIn as *const _ as *const u8, ptr as *mut u8, guid_size_in_bytes);

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
                                return Err(DV_E_TYMED);
                                
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
                                log::debug!("Writing {:?} to HGLOBAL", &CLSID_MMCSnapIn);
                                
                                let guid_size_in_bytes = std::mem::size_of::<GUID>();
                                

                                std::ptr::copy_nonoverlapping(&CLSID_MMCSnapIn as *const _ as *const u8, ptr as *mut u8, guid_size_in_bytes);

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
                                return Err(DV_E_TYMED);
                                
                            }
                        }
                    }
                    
                }
                _ => {
                    log::error!("Can't handle this type of clipboard format. tymed: {:?} btw",
                        unsafe { (*pmedium).0.tymed }
                    );
                    return Err(DV_E_FORMATETC);
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