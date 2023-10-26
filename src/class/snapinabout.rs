use intercom::prelude::*;
use windows::Win32::UI::WindowsAndMessaging::{HICON, LoadImageW, IMAGE_ICON, LR_DEFAULTSIZE};
use windows::Win32::Foundation::HMODULE;
use windows::{Win32::System::Com::CoTaskMemAlloc, core::PCWSTR};

use crate::__INTERCOM_DLL_INSTANCE;
use crate::interfaces::{ISnapinAbout, ComHICON, ComHBITMAP, ComCOLORREF, ComPCWSTR};

#[com_class(clsid = "d39d9c35-6106-4735-b944-7e929d607001", ISnapinAbout)]
#[derive(Default)]
pub struct MMCSnapInAbout {}

impl ISnapinAbout for MMCSnapInAbout {
    fn get_snapin_description(&self) -> ComResult<ComPCWSTR> {
        str_to_cotaskmem_compcwstr(
            "This is the snap-in description"
        )
    }

    fn get_provider(&self) -> ComResult<ComPCWSTR> {
        str_to_cotaskmem_compcwstr(
            "anthony.moe"
        )
    }
    
    fn get_snapin_version(&self) -> ComResult<ComPCWSTR> {
        str_to_cotaskmem_compcwstr(
            "0.0.69"
        )
    }
    
    // MMC creates a copy of the returned icon. The snap-in can free the icon
    // when the ISnapinAbout interface is released.
    // 
    // We accomplish this by storing the HICON in the struct so it can destroy
    // the data when it's not needed anymore.
    fn get_snapin_image(&self) -> ComResult<ComHICON> {
        let dll = unsafe { HMODULE(__INTERCOM_DLL_INSTANCE as isize) };
        
        let icon = unsafe {
            LoadImageW(
                dll,
                PCWSTR(101 as *const u16), // God awful
                IMAGE_ICON,
                0,
                0,
                LR_DEFAULTSIZE
            )
        };
        
        match icon {
            Ok(icon_h) => {
                return Ok(ComHICON(HICON(icon_h.0)));
            }
            
            Err(e) => {
                return Err(ComError::new_hr(intercom::raw::HRESULT { hr: e.code().0 }));
            }
        }

    }
    
    fn get_static_folder_image(&self) -> ComResult<(ComHBITMAP,ComHBITMAP,ComHBITMAP,ComCOLORREF)> {
        Err(ComError{ hresult: intercom::raw::HRESULT{ hr: 1 } , error_info: None })
    }
}

// Helper function, allocates string space using CoTaskMemAlloc and returns
// ComPCWSTR wrapped in result
fn str_to_cotaskmem_compcwstr(s: &str) -> ComResult<ComPCWSTR> { 
    let str: Vec<u16> = s.encode_utf16().chain(std::iter::once(0)).collect();
    let size = str.len() * std::mem::size_of::<u16>();
    let olestrbuf = unsafe { CoTaskMemAlloc(size) };
    
    if olestrbuf.is_null() {
        log::error!("CoTaskMemAlloc returns null pointer");
        return Err(ComError::E_FAIL);
    }
    
    unsafe {
        std::ptr::copy_nonoverlapping(str.as_ptr(), olestrbuf as *mut u16, str.len());
    }
    
    Ok(ComPCWSTR(PCWSTR::from_raw(olestrbuf as *const u16)))
}

impl Drop for MMCSnapInAbout {
    fn drop(&mut self) {
        // Delete the bitmaps
        /*
        if let Some(icon) = self.icon.get() {
            unsafe { DeleteObject(icon as _) };
        }
        */
    }
}