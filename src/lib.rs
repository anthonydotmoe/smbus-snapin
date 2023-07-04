use intercom::{prelude::*, BString};
use intercom::interfaces::IUnknown;
use winapi::shared::windef::HICON;
use winapi::um::winuser::{LoadIconW, MAKEINTRESOURCEW};
use winapi::um::libloaderapi::GetModuleHandleW;

mod interfaces;
pub mod mmc;

use interfaces::*;

use crate::mmc::MmcNotifyType;


// NOTE: NodeType GUID for "Fan" node: 977821db-5a85-4130-91c8-f0ad8636608e

com_library!(
    on_load=on_load,
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

#[com_class(clsid = "55ace359-cfbf-475f-a891-d1203894604c", IComponentData)]
#[derive(Default)]
struct MMCSnapIn {
    console: Option<ComRc<dyn IConsole>>,
    console_namespace: Option<ComRc<dyn IConsoleNamespace>>,
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
            Err(_e) => { log::error!("Error: QI for IConsole"); }
        }
        
        match console_namespace {
            Ok(console_namespace) => {
                log::debug!("Got {:p} for IConsoleNamespace", console_namespace.as_raw_iunknown());
                self.console_namespace = Some(console_namespace.clone())
            },
            Err(_e) => { log::error!("Error: QI for IConsoleNamespace"); }
        }
        
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
    
    fn query_data_object(&self,) -> ComResult<()> {
        Ok(())
    }
    
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
    fn initialize(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn notify(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn destroy(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn query_data_object(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn get_result_view_type(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn get_display_info(&self,) -> ComResult<()> {
        Ok(())
    }
    
    fn compare_objects(&self,) -> ComResult<()> {
        Ok(())
    }
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