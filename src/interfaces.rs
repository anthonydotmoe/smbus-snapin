use intercom::{prelude::*, IUnknown};

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
    fn SetHeader(&self, ) -> ComResult<i32>;

    //Sets IConsoles toolbar interface
    fn SetToolbar(&self, ) -> ComResult<i32>;

    //Queries IConsoles user provided IUnknown
    fn QueryResultView(&self, ) -> ComResult<i32>;

    //Queries the IConsole provided image list for the scope pane.
    fn QueryScopeImageList(&self, ) -> ComResult<i32>;

    //Queries the IConsole provided image list for the result pane.
    fn QueryResultImageList(&self, ) -> ComResult<i32>;

    //Generates a notification to update view(s) because of content change
    fn UpdateAllViews(&self, ) -> ComResult<i32>;

    //Displays a message box
    fn MessageBox(&self, ) -> ComResult<i32>;

    //Query for the IConsoleVerb.
    fn QueryConsoleVerb(&self, ) -> ComResult<i32>;

    //Selects the given scope item.
    fn SelectScopeItem(&self, ) -> ComResult<i32>;

    //Returns handle to the main frame window.
    fn GetMainWindow(&self, ) -> ComResult<i32>;

    //Create a new window rooted at the scope item specified by hScopeItem.
    fn NewWindow(&self, ) -> ComResult<i32>;
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
    fn InsertItem(&self, ) -> ComResult<i32>;

    // Allows the snap-in to delete a single item from the scope view.
    fn DeleteItem(&self, ) -> ComResult<i32>;

    // Allows the snap-in to set a single scope view item.
    fn SetItem(&self, ) -> ComResult<i32>;

    // Allows the snap-in to get a single scope view item.
    fn GetItem(&self, ) -> ComResult<i32>;

    // The handle of the child item if successful, otherwise NULL.
    fn GetChildItem(&self, ) -> ComResult<i32>;

    // The handle of the next item if successful, otherwise NULL.
    fn GetNextItem(&self, ) -> ComResult<i32>;

    // The handle of the parent item if successful, otherwise NULL.
    fn GetParentItem(&self, ) -> ComResult<i32>;
}

#[com_interface(com_iid = "255F18CC-65DB-11D1-A7DC-00C04FD8D565")]
pub trait IConsoleNamespace2: IConsoleNamespace {
    // Allows the snap-in to expand an item in the console namespace.
    fn Expand(&self, ) -> ComResult<i32>;

    // Add a dynamic extension to a selected node
    fn AddExtension(&self, ) -> ComResult<i32>;
}