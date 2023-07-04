use intercom::{prelude::*, IUnknown, BString};
use winapi::shared::windef::{COLORREF, HBITMAP, HICON};

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternOutput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComHBITMAP(pub HBITMAP);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternOutput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComCOLORREF(pub COLORREF);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternOutput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComHICON(pub HICON);

/*
#[com_interface(com_iid = "0000010e-0000-0000-C000-000000000046")]
pub trait IDataObject: IUnknown {
    fn GetData(&self, pformatetcIn: &FORMATETC, pmedium: &STGMEDIUM) -> ComResult<i32>;
}
*/

// IComponentData GUID from MMC SDK
#[com_interface(com_iid = "955AB28A-5218-11D0-A985-00C04FD8D565")]
pub trait IComponentData: IUnknown {
    fn initialize(&mut self, lp_unknown: &ComItf<dyn IUnknown>) -> ComResult<i32>;

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

#[com_interface(com_iid = "43136EB1-D36C-11CF-ADBC-00AA00A80033")]
pub trait IConsole: IUnknown {
    fn set_header(&self, ) -> ComResult<i32>;

    //Sets IConsoles toolbar interface
    fn set_toolbar(&self, ) -> ComResult<i32>;

    //Queries IConsoles user provided IUnknown
    fn query_result_view(&self, ) -> ComResult<i32>;

    //Queries the IConsole provided image list for the scope pane.
    fn query_scope_image_list(&self, ) -> ComResult<i32>;

    //Queries the IConsole provided image list for the result pane.
    fn query_result_image_list(&self, ) -> ComResult<i32>;

    //Generates a notification to update view(s) because of content change
    fn update_all_views(&self, ) -> ComResult<i32>;

    //Displays a message box
    fn message_box(&self, ) -> ComResult<i32>;

    //Query for the IConsoleVerb.
    fn query_console_verb(&self, ) -> ComResult<i32>;

    //Selects the given scope item.
    fn select_scope_item(&self, ) -> ComResult<i32>;

    //Returns handle to the main frame window.
    fn get_main_window(&self, ) -> ComResult<i32>;

    //Create a new window rooted at the scope item specified by hScopeItem.
    fn new_window(&self, ) -> ComResult<i32>;
}

#[com_interface(com_iid = "255F18CC-65DB-11D1-A7DC-00C04FD8D565")]
pub trait IConsole2: IConsole {
    // Allows the snap-in to expand/collapse a scope item in the corresponding
    // view. Should be called only by the IConsole associated with a IComponent.
    fn expand(&self, ) -> ComResult<i32>;

    // Determines if the user prefers taskpad views by default.
    fn is_taskpad_view_preferred(&self, ) -> ComResult<i32>;

    // Allows the snap-in to change the text on the status bar.
    fn set_status_text(&self, ) -> ComResult<i32>;
}

#[com_interface(com_iid = "BEDEB620-F24D-11cf-8AFC-00AA003CA9F6")]
pub trait IConsoleNamespace: IUnknown {
    // Allows the snap-in to insert a single item into the scope view.
    fn insert_item(&self, ) -> ComResult<i32>;

    // Allows the snap-in to delete a single item from the scope view.
    fn delete_item(&self, ) -> ComResult<i32>;

    // Allows the snap-in to set a single scope view item.
    fn set_item(&self, ) -> ComResult<i32>;

    // Allows the snap-in to get a single scope view item.
    fn get_item(&self, ) -> ComResult<i32>;

    // The handle of the child item if successful, otherwise NULL.
    fn get_child_item(&self, ) -> ComResult<i32>;

    // The handle of the next item if successful, otherwise NULL.
    fn get_next_item(&self, ) -> ComResult<i32>;

    // The handle of the parent item if successful, otherwise NULL.
    fn get_parent_item(&self, ) -> ComResult<i32>;
}

#[com_interface(com_iid = "255F18CC-65DB-11D1-A7DC-00C04FD8D565")]
pub trait IConsoleNamespace2: IConsoleNamespace {
    // Allows the snap-in to expand an item in the console namespace.
    fn expand(&self, ) -> ComResult<i32>;

    // Add a dynamic extension to a selected node
    fn add_extension(&self, ) -> ComResult<i32>;
}

#[com_interface(com_iid = "1245208C-A151-11D0-A7D7-00C04FD909DD")]
pub trait ISnapinAbout: IUnknown {
    // Text for the snap-in description box
    fn get_snapin_description(&self) -> ComResult<BString>;
    
    // Provider name
    fn get_provider(&self) -> ComResult<BString>;
    
    // Version number for the snap-in
    fn get_snapin_version(&self) -> ComResult<BString>;
    
    // Main icon for about box
    fn get_snapin_image(&self) -> ComResult<ComHICON>;
    
    // Static folder images for scope and result panes
    fn get_static_folder_image(&self) -> ComResult<(ComHBITMAP, ComHBITMAP, ComHBITMAP, ComCOLORREF)>;
}