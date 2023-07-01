use intercom::prelude::*;
use intercom::interfaces::IUnknown;

mod iid;
use iid::*;

// NOTE: NodeType GUID for "Fan" node: 977821db-5a85-4130-91c8-f0ad8636608e

com_library!(
    class MMCSnapIn
);

#[com_class(clsid = "55ace359-cfbf-475f-a891-d1203894604c", MMCSnapIn, IComponentData)]
#[derive(Default)]
pub struct MMCSnapIn;

#[com_interface]
impl IComponentData for MMCSnapIn {
    fn initialize(&self, _console: &ComItf<dyn IUnknown>) -> ComResult<i32> {
        // Not doing anything yet
        Ok(0)
    }
    // ... rest of the methods
}