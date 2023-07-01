use intercom::prelude::*;
use intercom::interfaces::IUnknown;

mod interfaces;
use interfaces::*;

// NOTE: NodeType GUID for "Fan" node: 977821db-5a85-4130-91c8-f0ad8636608e

com_library!(
    class MMCSnapIn
);

#[com_class(clsid = "55ace359-cfbf-475f-a891-d1203894604c", MMCSnapIn, IComponentData)]
#[derive(Default)]
pub struct MMCSnapIn {
    console: Option<ComRc<dyn IConsole2>>,
    console_namespace: Option<ComRc<dyn IConsoleNamespace2>>,
}

#[com_interface]
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