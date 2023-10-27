use intercom::prelude::*;

use crate::interfaces::{IComponent, IConsole, IConsole2, IDataObject, ComPCWSTR, IResultData, RESULTDATAITEM};

use super::{MmcNotifyType, MMCSnapIn};

#[com_class(IComponent)]
#[derive(Debug)]
pub struct MMCSnapInComponent {
    parent: *mut MMCSnapIn,
    console: Option<ComRc<dyn IConsole2>>,
    resultdata: Option<ComRc<dyn IResultData>>,
}

impl Default for MMCSnapInComponent {
    fn default() -> Self {
        Self {
            parent: std::ptr::null_mut(),
            console: None,
            resultdata: None,
        }
    }
}

impl MMCSnapInComponent {
    pub fn new(parent: *mut MMCSnapIn) -> Self {
        Self {
            parent,
            console: None,
            resultdata: None,
        }
    }
}

impl IComponent for MMCSnapInComponent {
    fn initialize(&mut self, lp_console: &ComItf<dyn IConsole>) -> ComResult<()> {
        
        // Cache the IConsole interface of the MMC
        let console2: ComResult<ComRc<dyn IConsole2>> = ComItf::query_interface(lp_console);
        let resultdata: ComResult<ComRc<dyn IResultData>> = ComItf::query_interface(lp_console);
        
        match console2 {
            Ok(console2) => {
                log::debug!("Go {:p} for IConsole2", console2.as_raw_iunknown());
                self.console = Some(console2.clone());
            },
            Err(e) => { log::error!("Error {:?}: QI for IConsole2", e); }
        }
        
        match resultdata {
            Ok(resultdata) => {
                log::debug!("Got {:p} for IResultData", resultdata.as_raw_iunknown());
                self.resultdata = Some(resultdata.clone());
            },
            Err(e) => { log::error!("Error {:?}: QI for IResultData", e); }
        }
        
        log::debug!("IComponent::Initialize done");

        Ok(())
    }
    
    fn destroy(&self) -> ComResult<()> {
        log::error!("Not implemented");
        Err(ComError::E_NOTIMPL)
    }
    
    fn compare_objects(&self, _obj_a: &ComItf<dyn IDataObject>, _obj_b: &ComItf<dyn IDataObject>) -> ComResult<()> {
        log::error!("Not implemented");
        Err(ComError::E_NOTIMPL)
    }
    
    fn get_display_info(&mut self, resultdataitem: *mut RESULTDATAITEM) -> ComResult<()> {
        let cookie = &(unsafe {*resultdataitem}.lparam.0);
        let node = unsafe { (*self.parent).nodes.get_mut(cookie) };

        match node {
            None => {
                log::error!("Couldn't match cookie: {}", cookie);
                return Err(ComError::E_POINTER);
            }
            Some(node) => {
                let mask = unsafe { (*resultdataitem).mask.clone() };
                if (mask & 0x0002) != 0 { // It wants the display string
                    match node.pcwstr() {
                        Ok(pcwstr) => {
                            unsafe {
                                (*resultdataitem).str.0 = pcwstr.0.0;
                            }
                        }
                        Err(e) => {
                            log::error!("Error setting str");
                            return Err(e);
                        }
                    }
                }
                if (mask & 0x0004) != 0 {
                    log::debug!("MMC wants the image, but I have no image to provide");
                }
                Ok(())
            }
        }
    }
    
    fn get_result_view_type(&self, _cookie:isize) -> ComResult<(ComPCWSTR,u64)> {
        // Return S_FALSE (0x1) to indicate using a List View
        Err(ComError::new_hr(intercom::raw::HRESULT { hr: 1 }))
        /*
        match &cookie {
            &0 => Ok((ComPCWSTR(PCWSTR(std::ptr::null())), 0)),
            _ => Err(ComError::E_NOTIMPL),
        }
        */
    }
    
    fn notify(&self, _lp_dataobject: &ComItf<dyn IDataObject>, event:u32, _arg:i64, _param:i64) -> ComResult<()> {
        let mmc_event: MmcNotifyType = unsafe { std::mem::transmute(event) };
        log::info!("Received event: {:#06X} ({:?})", event, mmc_event);
        
        Err(ComError::new_hr(intercom::raw::HRESULT { hr: 1 }))
        
        /*
        // Test insert_item()
        match mmc_event {
            MmcNotifyType::Show => {

                match &self.resultdata {
                    Some(resultdata) => {
                        let mut resultdataitem = RESULTDATAITEM {
                            mask: 0x0002,
                            scope_item: false,
                            itemid: HRESULTITEM(0),
                            index: 0,
                            col: 0,
                            str: MMC_CALLBACK,
                            image: 0,
                            state: 0,
                            lparam: LPARAM(1),
                            indent: 0,
                        };
                        
                        match resultdata.insert_item((&mut resultdataitem) as *mut _) {
                            Ok(_) => {
                                log::info!("Wow, inserting the item worked!\n{:?}", resultdataitem);
                                return Ok(());
                            }
                            Err(e) => {
                                log::error!("IResultData::InsertItem() error: {}", e);
                                return Ok(());
                            }
                        }
                        
                    }
                    None => Err(ComError::new_hr(intercom::raw::HRESULT { hr: 1 }))
                }
            }
            _ => Err(ComError::new_hr(intercom::raw::HRESULT { hr: 1 }))

        }
        */
    }
    
    fn query_data_object(&mut self, _cookie:isize, _type:i32) -> ComResult<ComRc<dyn IDataObject>> {
        log::error!("Not implemented");
        Err(ComError::E_NOTIMPL)
    }
}