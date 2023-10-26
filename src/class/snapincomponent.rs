use intercom::prelude::*;
use windows::core::PCWSTR;

use crate::interfaces::{IComponent, IConsole, ComPCWSTR};

use super::MmcNotifyType;

//use super::MMCSnapIn;

#[com_class(IComponent)]
#[derive(Default)]
pub struct MMCSnapInComponent {
    console: Option<ComRc<dyn IConsole>>,
    // parent: ComBox<MMCSnapIn>,
}

impl IComponent for MMCSnapInComponent {
    fn initialize(&mut self, lp_console: &ComItf<dyn IConsole>) -> ComResult<()> {
        
        // Cache the IConsole interface of the MMC
        self.console = Some(ComRc::from(lp_console));
        
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
        log::info!("Received event: {:#08X} ({:?})", event, mmc_event);

        Err(ComError::E_NOTIMPL)
    }
    
    fn query_data_object(&mut self,cookie:isize,r#type:i32) -> ComResult<ComRc<dyn crate::interfaces::IDataObject>> {
        log::error!("Not implemented");
        Err(ComError::E_NOTIMPL)
    }
}