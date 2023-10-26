use std::isize;

use intercom::{prelude::*, IUnknown};

use windows::Win32::System::Com::{FORMATETC, STGMEDIUM};
use windows::Win32::Foundation::{COLORREF, LPARAM};
use windows::Win32::UI::WindowsAndMessaging::HICON;
use windows::Win32::Graphics::Gdi::HBITMAP;
use windows::core::PCWSTR;


#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternOutput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComPCWSTR(pub PCWSTR);

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

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComFORMATETC(pub FORMATETC);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput, intercom::ExternOutput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComSTGMEDIUM(pub STGMEDIUM);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput, intercom::ExternOutput)]
#[derive(Debug, Clone, Copy)]
#[allow(non_camel_case_types)]
#[repr(C)]
pub struct SCOPEDATAITEM {
    pub mask: u32,
    pub display_name: PCWSTR,
    pub image: i32,
    pub open_image: i32,
    pub state: u32,
    pub children: i32,
    pub lparam: LPARAM,
    pub relative_id: HSCOPEITEM,
    pub id: HSCOPEITEM,
}

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput, intercom::ExternOutput)]
#[derive(Debug, Clone, Copy)]
#[allow(non_camel_case_types)]
#[repr(C)]
pub struct RESULTDATAITEM {
    pub mask: u32,
    pub scope_item: bool,
    pub itemid: HRESULTITEM,
    pub index: i32,
    pub col: i32,
    pub str: PCWSTR,
    pub image: i32,
    pub state: u32,
    pub lparam: LPARAM,
    pub indent: i32,
}

// Should be correct
pub const MMC_CALLBACK: PCWSTR = PCWSTR::from_raw(usize::MAX as *const u16);

#[derive(Debug, Clone, Copy)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct HSCOPEITEM(pub isize);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[derive(Debug, Clone, Copy)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct HRESULTITEM(pub isize);

#[com_interface(com_iid = "0000010e-0000-0000-C000-000000000046")]
pub trait IDataObject: IUnknown {
    fn get_data(&self, pformatetc: *const ComFORMATETC) -> ComResult<*mut ComSTGMEDIUM>;
    fn get_data_here(&self, pformatetc: *const ComFORMATETC, pmedium: *mut ComSTGMEDIUM) -> ComResult<()>;
    fn query_get_data(&self, ) -> ComResult<()>;
    fn get_canonical_format(&self, ) -> ComResult<()>;
    fn set_data(&self, ) -> ComResult<()>;
    fn enum_format_etc(&self, ) -> ComResult<()>;
    fn d_advise(&self, ) -> ComResult<()>;
    fn d_unadvise(&self, ) -> ComResult<()>;
    fn enum_d_advise(&self) -> ComResult<()>;
}

#[com_interface(com_iid = "955AB28A-5218-11D0-A985-00C04FD8D565")]
pub trait IComponentData: IUnknown {
    // Snap-in entry point. Can QI for IConsole & IConsoleNameSpace
    fn initialize(&mut self, lp_unknown: &ComItf<dyn IUnknown>) -> ComResult<()>;

    // Create a Component for this ComponentData
    fn create_component(&mut self) -> ComResult<ComRc<dyn IComponent>>;

    // User actions
    fn notify(&self, lp_dataobject: &ComItf<dyn IDataObject>, event: u32, arg: i64, param: i64) -> ComResult<()>;

    // Release cookies associated with the children of a specific node
    fn destroy(&self) -> ComResult<()>;

    // Returns a data object which may be used to retrieve the context information for the specified cookie
    fn query_data_object(&mut self, cookie: isize, r#type: i32) -> ComResult<ComRc<dyn IDataObject>>;

    // Get display info for the name space item
    fn get_display_info(&self, lpscopedataitem: *mut SCOPEDATAITEM) -> ComResult<()>;

    // The snap-in's compare function for two data objects
    fn compare_objects(&self, ) -> ComResult<()>;
}

#[com_interface(com_iid = "43136eb2-d36c-11cf-adbc-00aa00a80033")]
pub trait IComponent: IUnknown {
    // provides an entry point to the console.
    fn initialize(&mut self, lp_console: &ComItf<dyn IConsole>) -> ComResult<()>;

    // notifies the snap-in of actions taken by the user.
    fn notify(&self, lp_dataobject: &ComItf<dyn IDataObject>, event: u32, arg: i64, param: i64) -> ComResult<()>;

    // releases all references to the console that are held by this component.
    fn destroy(&self) -> ComResult<()>;

    // returns a data object that can be used to retrieve context information
    // for the specified cookie.
    fn query_data_object(&mut self, cookie: isize, r#type: i32) -> ComResult<ComRc<dyn IDataObject>>;
 
    // determines what the result pane view should be.
    fn get_result_view_type(&self, cookie: isize) -> ComResult<(ComPCWSTR, u64)>;

    // retrieves display information for an item in the result pane.
    fn get_display_info(&self, ) -> ComResult<()>;

    // enables a snap-in to compare two data objects acquired through
    // IComponent::QueryDataObject. Be aware that data objects can be acquired
    // from two different instances of IComponent.
    fn compare_objects(&self, ) -> ComResult<()>;
}

#[com_interface(com_iid = "43136EB1-D36C-11CF-ADBC-00AA00A80033")]
pub trait IConsole: IUnknown {
    // Sets IConsoles header interface
    fn set_header(&self, ) -> ComResult<()>;

    // Sets IConsoles toolbar interface
    fn set_toolbar(&self, ) -> ComResult<i32>;

    // Queries IConsoles user provided IUnknown
    fn query_result_view(&self, ) -> ComResult<i32>;

    // Queries the IConsole provided image list for the scope pane.
    fn query_scope_image_list(&self, ) -> ComResult<i32>;

    // Queries the IConsole provided image list for the result pane.
    fn query_result_image_list(&self, ) -> ComResult<i32>;

    // Generates a notification to update view(s) because of content change
    fn update_all_views(&self, ) -> ComResult<i32>;

    // Displays a message box
    fn message_box(&self, ) -> ComResult<i32>;

    // Query for the IConsoleVerb.
    fn query_console_verb(&self, ) -> ComResult<i32>;

    // Selects the given scope item.
    fn select_scope_item(&self, ) -> ComResult<i32>;

    // Returns handle to the main frame window.
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
    fn insert_item(&self, lpscopedataitem: *mut SCOPEDATAITEM) -> ComResult<()>;

    // Allows the snap-in to delete a single item from the scope view.
    fn delete_item(&self, ) -> ComResult<()>;

    // Allows the snap-in to set a single scope view item.
    fn set_item(&self, ) -> ComResult<()>;

    // Allows the snap-in to get a single scope view item.
    fn get_item(&self, ) -> ComResult<()>;

    // The handle of the child item if successful, otherwise NULL.
    fn get_child_item(&self, ) -> ComResult<()>;

    // The handle of the next item if successful, otherwise NULL.
    fn get_next_item(&self, ) -> ComResult<()>;

    // The handle of the parent item if successful, otherwise NULL.
    fn get_parent_item(&self, ) -> ComResult<()>;
}

#[com_interface(com_iid = "255F18CC-65DB-11D1-A7DC-00C04FD8D565")]
pub trait IConsoleNamespace2: IConsoleNamespace {
    // Allows the snap-in to expand an item in the console namespace.
    fn expand(&self, ) -> ComResult<()>;

    // Add a dynamic extension to a selected node
    fn add_extension(&self, ) -> ComResult<()>;
}

/// The functions that return strings--`get_snapin_description`, `get_provider`,
/// and `get_snapin_version`--must allocate memory for out parameters using the
/// COM API function `CoTaskMemAlloc`
#[com_interface(com_iid = "1245208C-A151-11D0-A7D7-00C04FD909DD")]
pub trait ISnapinAbout: IUnknown {
    // Text for the snap-in description box
    fn get_snapin_description(&self) -> ComResult<ComPCWSTR>;
    
    // Provider name
    fn get_provider(&self) -> ComResult<ComPCWSTR>;
    
    // Version number for the snap-in
    fn get_snapin_version(&self) -> ComResult<ComPCWSTR>;
    
    // Main icon for about box
    fn get_snapin_image(&self) -> ComResult<ComHICON>;
    
    // Static folder images for scope and result panes
    fn get_static_folder_image(&self) -> ComResult<(ComHBITMAP, ComHBITMAP, ComHBITMAP, ComCOLORREF)>;
}

#[com_interface(com_iid = "72782D7A-A4A0-11D1-AF0F-00C04FB6DD2C")]
pub trait IRequiredExtensions: IUnknown {
    // Enable all extensions
    fn enable_all_extensions(&self) -> ComResult<()>;
    
    // Get first required extension
    // TODO: Implement ComCLSID fn get_first_extension(&self) -> ComResult<ComCLSID>
    fn get_first_extension(&self) -> ComResult<()>;

    // Get next required extension
    // TODO: Implement ComCLSID fn get_first_extension(&self) -> ComResult<ComCLSID>
    fn get_next_extension(&self) -> ComResult<()>;
}

// The next step is to get IResultData working
// https://learn.microsoft.com/en-us/previous-versions/windows/desktop/mmc/using-list-views-implementation-details
#[com_interface(com_iid = "31DA5FA0-E0EB-11cf-9F21-00AA003CA9F6")]
pub trait IResultData: IUnknown {
    // Allows the snap-in to insert a single item.
    fn insert_item(&self, resultdataitem: *mut RESULTDATAITEM) -> ComResult<()>;

    // Allows the snap-in to delete a single item.
    fn delete_item(&self, itemid: HRESULTITEM, _reserved: std::ffi::c_int) -> ComResult<()>;

    // Allows the snap-in to find an item/subitem based on its user inserted lParam.
    // HRESULT FindItemByLParam([in] LPARAM lParam, [out] HRESULTITEM *pItemID);
    fn find_item_by_lparam(&self, ) -> ComResult<()>;

    // Allows the snap-in to delete all the items.
    // HRESULT DeleteAllRsltItems();
    fn delete_all_rslt_items(&self, ) -> ComResult<()>;

    // Allows the snap-in to set a single item.
    // HRESULT SetItem([in] LPRESULTDATAITEM item);
    fn set_item(&self, ) -> ComResult<()>;

    // Allows the snap-in to get a single item.
    // HRESULT GetItem([in,out] LPRESULTDATAITEM item);
    fn get_item(&self, ) -> ComResult<()>;

    // Returns the lParam of the first item, which matches the given state.
    // HRESULT GetNextItem([in,out] LPRESULTDATAITEM item);
    fn get_next_item(&self, ) -> ComResult<()>;

    // Allows the snap-in to modify the state of an item.
    // HRESULT ModifyItemState([in] int nIndex, [in] HRESULTITEM itemID,
    //                      [in] UINT uAdd, [in] UINT uRemove);
    fn modify_item_state(&self, ) -> ComResult<()>;

    // Allows the snap-in to set the result view style.
    // HRESULT ModifyViewStyle([in] MMC_RESULT_VIEW_STYLE add,
    //                   [in] MMC_RESULT_VIEW_STYLE remove);
    fn modify_view_style(&self, ) -> ComResult<()>;

    // Allows the snap-in to set the result view mode.
    // HRESULT SetViewMode([in] long lViewMode);
    fn set_view_mode(&self, ) -> ComResult<()>;

    // Allows the snap-in to get the result view mode.
    // HRESULT GetViewMode([out] long* lViewMode);
    fn get_view_mode(&self, ) -> ComResult<()>;

    // Allows the snap-in to update a single item.
    // HRESULT UpdateItem([in] HRESULTITEM itemID);
    fn update_item(&self, ) -> ComResult<()>;

    // Sort all items in result pane
    // HRESULT Sort([in] int nColumn, [in] DWORD dwSortOptions, [in] LPARAM lUserParam);
    fn sort(&self, ) -> ComResult<()>;

    // Set the description bar text for the result view
    // HRESULT SetDescBarText([in] LPOLESTR DescText);
    fn set_desc_bar_text(&self, ) -> ComResult<()>;

    // Set number of items in result pane list
    // HRESULT SetItemCount([in] int nItemCount, [in] DWORD dwOptions);
    fn set_item_count(&self, ) -> ComResult<()>;
}