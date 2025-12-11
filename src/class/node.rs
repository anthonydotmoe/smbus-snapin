use guid::GUID;
use intercom::{prelude::*, raw::HRESULT};
use windows::{Win32::{System::{Memory::{ GlobalUnlock, GlobalLock, GlobalSize }, DataExchange::GetClipboardFormatNameW, Com::{ CoTaskMemFree, CoTaskMemAlloc }}, Foundation::{MAX_PATH, GetLastError, NO_ERROR}}, core::PCWSTR};

use crate::{interfaces::{IDataObject, ComFORMATETC, ComSTGMEDIUM, HSCOPEITEM, ComPCWSTR}, class::snapin::CLSID_MMCSnapIn};

use super::MMCSnapIn;

pub const DV_E_FORMATETC: ComError = ComError{
    hresult: HRESULT { hr: windows::Win32::Foundation::DV_E_FORMATETC.0 },
    error_info: None,
};

pub const DV_E_TYMED: ComError = ComError{
    hresult: HRESULT { hr: windows::Win32::Foundation::DV_E_TYMED.0 },
    error_info: None,
};

#[derive(Debug, Default, PartialEq)]
pub enum NodeType {
    #[default]
    Folder,
    Root,
}

#[com_class(IDataObject)]
#[derive(Debug)]
pub struct Node {
    _owner: *const MMCSnapIn,
    pub node_type: NodeType,
    pcwstr_name: Option<PCWSTR>,
    pub display_name: String,
    pub hscopeitem: HSCOPEITEM,
}

impl Default for Node {
    fn default() -> Self {
        Self {
            _owner: std::ptr::null(),
            node_type: NodeType::Folder,
            display_name: String::new(),
            pcwstr_name: None,
            hscopeitem: HSCOPEITEM(0),
        }
    }
}

impl Node {
    pub fn new(owner: *const MMCSnapIn, name: String, ntype: NodeType) -> Self {
        Node {
            _owner: owner,
            display_name: name,
            node_type: ntype,
            pcwstr_name: None,
            hscopeitem: HSCOPEITEM(0),
        }
    }
    
    // Calls to this function release the pointer to the PCWSTR and allocate a new one.
    pub fn pcwstr(&mut self) -> ComResult<ComPCWSTR> {
        if let Some(old) = self.pcwstr_name {
            unsafe { CoTaskMemFree(Some(old.0 as *const _)); }
            self.pcwstr_name = None;
        }
        
        log::debug!("Converting string \"{}\" to PCWSTR", self.display_name);
        
        let wide: Vec<u16> = self.display_name
            .to_owned()
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();
        
        let size_bytes = wide.len() * std::mem::size_of::<u16>();
        let olestrbuf = unsafe { CoTaskMemAlloc(size_bytes) };
        if olestrbuf.is_null() {
            log::error!("CoTaskMemAlloc returns null pointer");
            return Err(ComError::E_FAIL);
        }
        
        unsafe {
            std::ptr::copy_nonoverlapping(
                wide.as_ptr(),
                olestrbuf as *mut u16,
                wide.len());
        }

        let ptr = PCWSTR::from_raw(olestrbuf as *const u16);
        self.pcwstr_name = Some(ptr);
        
        Ok(ComPCWSTR(ptr))
    }
}

impl Drop for Node {
    fn drop(&mut self) {
        match self.pcwstr_name {
            Some(str) => unsafe {
                CoTaskMemFree(Some(str.0 as *const _));
            },
            None => {}
        }
    }
}

impl IDataObject for Node {
    fn get_data(&self, _pformatetc: *const ComFORMATETC) -> ComResult<*mut ComSTGMEDIUM> {
        Err(ComError::E_NOTIMPL)
    }
    
    fn get_data_here(&self, pformatetc: *const ComFORMATETC, pmedium: *mut ComSTGMEDIUM) -> ComResult<()> {
        let clipformat: u32 = unsafe { (*pformatetc).0.cfFormat.into() };
        log::debug!("Got Clipformat: {}", clipformat);

        let tymed: TagTYMED = unsafe { std::mem::transmute((*pmedium).0.tymed.0) };
        
        let mut clipformat_name: [u16; MAX_PATH as usize] = [0; MAX_PATH as usize];
        
        let clipformat_name_len: i32 = unsafe { GetClipboardFormatNameW(clipformat, &mut clipformat_name) };
        
        if clipformat_name_len > 0 {
            let name = String::from_utf16_lossy(&clipformat_name[0..clipformat_name_len as usize]);
            log::debug!("Got clipformat name: {}", name);
            
            match name.as_str() {
                "CCF_DISPLAY_NAME" => {
                    // Return node display name as WSTR
                    unsafe {
                        match tymed {
                            TagTYMED::HGlobal => {
                                let ptr = GlobalLock((*pmedium).0.Anonymous.hGlobal);
                                
                                if ptr.is_null() {
                                    log::error!("HGLOBAL returns null pointer");
                                    return Err(ComError::E_FAIL);
                                }
                                
                                let node_name_utf16: Vec<u16> = self.display_name.encode_utf16().chain(std::iter::once::<u16>(0)).collect();
                                
                                std::ptr::copy_nonoverlapping(
                                    node_name_utf16.as_ptr(),
                                    ptr as *mut u16,
                                    node_name_utf16.len(),
                                );
                                
                                global_unlock_checked((*pmedium).0.Anonymous.hGlobal)?;
                                
                            }
                            _ => {
                                log::error!("Unsupported TYMED: {:?}", tymed);
                                return Err(DV_E_TYMED);
                            }
                        }
                    }
                }
                "CCF_NODETYPE" => {
                    // Return the GUID of the node type
                    unsafe {
                        match tymed {
                            TagTYMED::HGlobal => {
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

                                global_unlock_checked((*pmedium).0.Anonymous.hGlobal)?;
                            }
                            _ => {
                                log::error!("Unsupported TYMED: {:?}", tymed);
                                return Err(DV_E_TYMED);
                                
                            }
                        }
                    }
                }
                "CCF_SNAPIN_CLSID" => {
                    // Return the CLSID of the snap in
                    unsafe {
                        match tymed {
                            TagTYMED::HGlobal => {
                                let ptr = GlobalLock((*pmedium).0.Anonymous.hGlobal);
                                
                                if ptr.is_null() {
                                    log::error!("HGLOBAL returns null pointer");
                                    return Err(ComError::E_FAIL);
                                }
                                
                                let guid_size_in_bytes = std::mem::size_of::<GUID>();
                                std::ptr::copy_nonoverlapping(&CLSID_MMCSnapIn as *const _ as *const u8, ptr as *mut u8, guid_size_in_bytes);

                                global_unlock_checked((*pmedium).0.Anonymous.hGlobal)?;
                            }
                            _ => {
                                log::error!("Unsupported TYMED: {:?}", tymed);
                                return Err(DV_E_TYMED);
                                
                            }
                        }
                    }
                    
                }
                /*
                 * Clipboard formats:
                 *   - CCF_NODETYPE
                 *   - CCF_SZNODETYPE
                 *   - CCF_DISPLAY_NAME
                 *   - CCF_SNAPIN_CLASSID
                 *   - CCF_SNAPIN_CLASS
                 *   - CCF_WINDOW_TITLE
                 *   - CCF_MMC_MULTISELECT_DATAOBJECT
                 *   - CCF_MULTI_SELECT_SNAPINS
                 *   - CCF_OBJECT_TYPES_IN_MULTI_SELECT
                 *   - CCF_MMC_DYNAMIC_EXTENSIONS
                 *   - CCF_SNAPIN_PRELOADS
                 *   - CCF_NODEID2
                 *   - CCF_NODEID
                 *   - CCF_COLUMN_SET_ID
                 *   - CCF_DESCRIPTION
                 *   - CCF_HTML_DETAILS
                */
                _ => {
                    log::error!("Can't handle this type of clipboard format. tymed: {:?} btw", tymed);
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

fn global_unlock_checked(hglobal: HGLOBAL) -> ComResult<()> {
    let res = unsafe { GlobalUnlock(hglobal) };
    if res.0 == 0 {
        let err = unsafe { GetLastError() };
        if err != NO_ERROR {
            log::error!("Failed to unlock HGLOBAL: {:?}", err);
            return Err(ComError::E_FAIL);
        }
    }
    Ok(())
}

#[repr(i32)]
#[derive(Debug)]
#[allow(dead_code)]
pub enum TagTYMED {
    HGlobal     = 1,
    File        = 2,
    IStream     = 4,
    IStorage    = 8,
    Gdi         = 16,
    Metafile    = 32,
    EnhMetafile = 64,
    Null        = 0,
}
