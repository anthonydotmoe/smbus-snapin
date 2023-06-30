use intercom::*;

mod iid;
use iid::*;

// IComponentData GUID from MMC SDK
#[com_interface(com_iid = "{955AB28A-5218-11D0-A985-00C04FD8D565}")]
pub trait IComponentData: IUnknown {
    fn Initialize(&self, pUnknown: dyn IUnknown) -> ComResult<i32>;

    fn CreateComponent(&self, ppComponent: LPCOMPONENT) -> ComResult<i32>;

    fn Notify(&self, lpDataObject: LPDATAOBJECT, event: MMC_NOTIFY_TYPE, arg: LPARAM, param: LPARAM) -> ComResult<i32>;

    fn Destroy(&self) -> ComResult<i32>;

    fn QueryDataObject(&self, cookie: MMC_COOKIE, r#type: DATA_OBJECT_TYPES, ppDataObject: LPDATAOBJECT) -> ComResult<i32>;

    fn GetDisplayInfo(&self, pScopeDataItem: SCOPEDATAITEM) -> ComResult<i32>;

    fn CompareObjects(&self, lpDataObjectA: LPDATAOBJECT, lpDataObjectB: LPDATAOBJECT) -> ComResult<i32>;
}


#[com_class(clsid = "{55ace359-cfbf-475f-a891-d1203894604c}", ComSnapIn, IComponentData)]
#[derive(Default)]
pub struct ComSnapIn {
    // Any necessary state...
}

#[com_interface]
impl IComponentData for ComSnapIn {

}