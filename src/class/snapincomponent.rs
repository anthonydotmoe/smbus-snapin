use intercom::prelude::*;
use windows::{core::PCWSTR, w, Win32::Foundation::LPARAM};

use crate::interfaces::{IComponent, IConsole, ComPCWSTR, IResultData, RESULTDATAITEM, HRESULTITEM, MMC_CALLBACK};

use super::MmcNotifyType;

//use super::MMCSnapIn;

#[com_class(IComponent)]
#[derive(Default)]
pub struct MMCSnapInComponent {
    console: Option<ComRc<dyn IConsole>>,
    resultdata: Option<ComRc<dyn IResultData>>,
    // parent: ComBox<MMCSnapIn>,
}

impl IComponent for MMCSnapInComponent {
    fn initialize(&mut self, lp_console: &ComItf<dyn IConsole>) -> ComResult<()> {
        
        // Cache the IConsole interface of the MMC
        self.console = Some(ComRc::from(lp_console));
        let resultdata: ComResult<ComRc<dyn IResultData>> = ComItf::query_interface(lp_console);
        
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
    
    fn compare_objects(&self,) -> ComResult<()> {
        log::error!("Not implemented");
        Err(ComError::E_NOTIMPL)
    }
    
    fn get_display_info(&self,) -> ComResult<()> {
        log::error!("Not implemented");
        Err(ComError::E_NOTIMPL)
    }
    
    fn get_result_view_type(&self, cookie:isize) -> ComResult<(ComPCWSTR,u64)> {
        
        match &cookie {
            &0 => Ok((ComPCWSTR(PCWSTR(std::ptr::null())), 0)),
            _ => Err(ComError::E_NOTIMPL),
        }
    }
    
    fn notify(&self,lp_dataobject: &ComItf<dyn crate::interfaces::IDataObject>,event:u32,arg:i64,param:i64) -> ComResult<()> {
        let mmc_event: MmcNotifyType = unsafe { std::mem::transmute(event) };
        log::info!("Received event: {:#06X} ({:?})", event, mmc_event);
        
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
                    None => { Err(ComError::E_FAIL) }
                }
            }
            _ => Err(ComError::new_hr(intercom::raw::HRESULT { hr: 1 }))

        }
    }
    
    fn query_data_object(&mut self,cookie:isize,r#type:i32) -> ComResult<ComRc<dyn crate::interfaces::IDataObject>> {
        log::error!("Not implemented");
        Err(ComError::E_NOTIMPL)
    }
}