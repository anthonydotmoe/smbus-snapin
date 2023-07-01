use intercom::{prelude::*, IUnknown};

/*
#[com_interface(com_iid = "0000010e-0000-0000-C000-000000000046")]
pub trait IDataObject: IUnknown {
    fn GetData(&self, pformatetcIn: &FORMATETC, pmedium: &STGMEDIUM) -> ComResult<i32>;
}
*/

// IComponentData GUID from MMC SDK
#[com_interface(com_iid = "{955AB28A-5218-11D0-A985-00C04FD8D565}")]
pub trait IComponentData: IUnknown {
    fn initialize(&self, console: &ComItf<dyn IUnknown>) -> ComResult<i32>;

    /* TO BE IMPLEMENTED
    fn CreateComponent(&self, ppComponent: *mut ComPtr<dyn IComponent>) -> ComResult<i32>;

    fn Notify(&self, lpDataObject: LPDATAOBJECT, event: MMC_NOTIFY_TYPE, arg: LPARAM, param: LPARAM) -> ComResult<i32>;

    fn Destroy(&self) -> ComResult<i32>;

    fn QueryDataObject(&self, cookie: MMC_COOKIE, r#type: DATA_OBJECT_TYPES, ppDataObject: LPDATAOBJECT) -> ComResult<i32>;

    fn GetDisplayInfo(&self, pScopeDataItem: SCOPEDATAITEM) -> ComResult<i32>;

    fn CompareObjects(&self, lpDataObjectA: LPDATAOBJECT, lpDataObjectB: LPDATAOBJECT) -> ComResult<i32>;
    */
}

#[com_interface(com_iid = "43136eb2-d36c-11cf-adbc-00aa00a80033")]
pub trait IComponent: IUnknown {
    // provides an entry point to the console.
    fn initialize(&self, ) -> ComResult<i32>;

    // notifies the snap-in of actions taken by the user.
    fn notify(&self, ) -> ComResult<i32>;

    // releases all references to the console that are held by this component.
    fn destroy(&self, ) -> ComResult<i32>;

    // returns a data object that can be used to retrieve context information
    // for the specified cookie.
    fn query_data_object(&self, ) -> ComResult<i32>;
 
    // determines what the result pane view should be.
    fn get_result_view_type(&self, ) -> ComResult<i32>;

    // retrieves display information for an item in the result pane.
    fn get_display_info(&self, ) -> ComResult<i32>;

    // enables a snap-in to compare two data objects acquired through
    // IComponent::QueryDataObject. Be aware that data objects can be acquired
    // from two different instances of IComponent.
    fn compare_objects(&self, ) -> ComResult<i32>;
}