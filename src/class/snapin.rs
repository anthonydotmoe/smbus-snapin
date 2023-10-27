use std::collections::HashMap;

use intercom::{ IUnknown, prelude::* };
use windows::Win32::Foundation::LPARAM;

use crate::class::node::NodeType;
use crate::MMCSnapInComponent;
use crate::interfaces::*;
use crate::Node;

#[com_class(clsid = "d39d9c35-6106-4735-b944-7e929d607000", IComponentData)]
#[derive(Debug)]
pub struct MMCSnapIn {
    console: Option<ComRc<dyn IConsole2>>,
    console_namespace: Option<ComRc<dyn IConsoleNamespace>>,
    pub nodes: HashMap<isize, ComBox<Node>>,
    //_components: Vec<ComBox<MMCSnapInComponent>>,
    //nodes: HashMap<isize, ComRc<dyn IDataObject>>,
    next_cookie: isize,
}

// Impl'd because default Default implementation makes next_cookie = 0.
impl Default for MMCSnapIn {
    fn default() -> Self {
        MMCSnapIn {
            console: None,
            console_namespace: None,
            nodes: HashMap::<isize, ComBox<Node>>::default(),
            //_components: Vec::new(),
            next_cookie: 1,
        }
    }
}

impl MMCSnapIn {
    fn add_node(&mut self, name: &str) {
        let new_node = Node::new(
            self as *const _,
            name.to_owned(),
            NodeType::Folder,
        );
        
        self.nodes.insert(self.next_cookie, ComBox::new(new_node));
        
        self.next_cookie += 1;
    }
    
    fn add_root_node(&mut self) {
        match self.nodes.get(&0) {
            Some(_) => return,
            None => {
                let root_node = Node::new(
                    self as *const _,
                    "SMBus Snap-in".to_owned(),
                    NodeType::Root,
                );
                self.nodes.insert(0, ComBox::new(root_node));
            }
        }
    }
}

impl IComponentData for MMCSnapIn {
    fn initialize(&mut self, lp_unknown: &ComItf<dyn IUnknown>) -> ComResult<()> {
        log::debug!("IComponentData for MMCSnapIn called");

        // Use the received IUnknown interface to query for IConsole2 and
        // IConsoleNamespace2

        let console: ComResult<ComRc<dyn IConsole2>>  = ComItf::query_interface(lp_unknown);
        let console_namespace: ComResult<ComRc<dyn IConsoleNamespace>>  = ComItf::query_interface(lp_unknown);
        
        match console {
            Ok(console) => {
                log::debug!("Got: {:p} for IConsole", console.as_raw_iunknown());
                self.console = Some(console.clone());
            },
            Err(e) => { log::error!("Error {:?}: QI for IConsole", e); }
        }
        
        match console_namespace {
            Ok(console_namespace) => {
                log::debug!("Got {:p} for IConsoleNamespace", console_namespace.as_raw_iunknown());
                self.console_namespace = Some(console_namespace.clone())
            },
            Err(e) => { log::error!("Error {:?}: QI for IConsoleNamespace", e); }
        }
        
        log::debug!("IComponentData::Initialize done");
        
        // Make a few nodes
        self.add_node("Node 1");
        self.add_node("Node 2");
        self.add_node("Node 3");
        
        Ok(())
    }
    
    fn create_component(&mut self) -> ComResult<ComRc<dyn IComponent>> {
        //Err(ComError::E_NOTIMPL)

        
        let component = MMCSnapInComponent::new(self as *mut _);
        Ok(ComRc::from(ComBox::new(component)))
        // or 
        /*
        let component = MMCSnapInComponent::new(self.into());
        
        Ok(component.into())
        */

    }
    
    fn notify(&mut self, _lp_dataobject: &ComItf<dyn IDataObject>, event:u32, arg:i64, param:i64) -> ComResult<()> {
        let mmc_event: MmcNotifyType = unsafe { std::mem::transmute(event) };
        log::info!("Received event: {:#06X} ({:?})", event, mmc_event);
        
        // Try adding the node to the scope pane
        if mmc_event == MmcNotifyType::Expand {
            log::info!("{} {}", param, if arg == 0 { "Collapsed" } else { "Expanded" });
            match &self.console_namespace {
                Some(consolens) => {
                    for (cookie, node) in self.nodes.iter_mut() {
                        
                        if node.node_type == NodeType::Root {
                            continue;
                        }

                        if node.hscopeitem.0 == 0 {

                            let mut scopedataitem = SCOPEDATAITEM {
                                mask: 0x00002 | 0x00020 | 0x00040, // SDI_STR | SDI_PARAM | SDI_CHILDREN
                                display_name: crate::interfaces::MMC_CALLBACK,
                                image: 0,
                                open_image: 0,
                                state: 0,
                                children: 0,
                                lparam: LPARAM(cookie.clone()),
                                relative_id: HSCOPEITEM(param as isize),
                                id: HSCOPEITEM(0),
                            };
            
                            match consolens.insert_item((&mut scopedataitem) as *mut _) {
                                Ok(_) => {
                                    // Store the id back in the Node struct
                                    node.hscopeitem = scopedataitem.id;
                                }
                                Err(e) => {
                                    log::error!("IConsoleNamespace::InsertItem() error: {}", e)
                                }
                            }
                        }
                    }
                }
                None => {}
            }
        }
        Ok(())
    }
    
    fn destroy(&self) -> ComResult<()> {
        Ok(())
    }
    
    fn query_data_object(&mut self, cookie: isize, view_type: i32) -> ComResult<ComRc<dyn IDataObject>> {
        let view_type: MmcDataObjectType = unsafe { std::mem::transmute(view_type) };
    
        log::debug!("QueryDataObject: cookie: {}, ppviewtype: {:?}", cookie, view_type);
        
        match &cookie {
            0 => {
                // View type doesn't matter for root node
                match self.nodes.get(&cookie) {
                    None => {
                        log::debug!("No root node has been made yet.");
                        
                        // Create the root node
                        self.add_root_node();
                        return Ok(ComRc::from(self.nodes.get(&cookie).unwrap()));
                    }
                    Some(root_node) => {
                        log::debug!("Root node IDataObject returned to caller");
                        return Ok(ComRc::from(root_node));
                    }
                }
            }
            // Anything but the root node
            _ => {
                match self.nodes.get(&cookie) {
                    Some(node) => {
                        log::debug!("Node #{} IDataObject returned to caller", &cookie);
                        return Ok(ComRc::from(node));
                    }
                    None => {
                        log::error!("MMC tried to get node with cookie: {}", &cookie);
                        return Err(ComError::E_POINTER);
                    }
                }
            }
        }
    }
    
    fn get_display_info(&mut self,lpscopedataitem: *mut SCOPEDATAITEM) -> ComResult<()> {
        log::debug!("Got {:?}", unsafe { *lpscopedataitem });
        let cookie = &(unsafe {*lpscopedataitem}.lparam.0);
        let nameptr = self.nodes.get_mut(cookie);
        
        match nameptr {
            None => {
                log::error!("Couldn't match cookie: {}", &cookie);
                return Err(ComError::E_POINTER);
            }
            Some(obj) => {
                let mask = unsafe { (*lpscopedataitem).mask.clone() };
                if (mask & 0x0002) != 0 {  // It wants the display string
                    match obj.pcwstr() {
                        Ok(pcwstr) => {
                            unsafe {
                                (*lpscopedataitem).display_name.0 = pcwstr.0.0;
                            }
                        }
                        Err(e) => {
                            log::error!("Error setting string");
                            return Err(e);
                        }
                    }
                }
                Ok(())
            }
        }
    }

    fn compare_objects(&self, _obj_a: &ComItf<dyn IDataObject>, _obj_b: &ComItf<dyn IDataObject>) -> ComResult<()> {
        log::error!("Not implemented");
        Ok(())
    }
}

/*
impl IRequiredExtensions for MMCSnapIn {
    fn enable_all_extensions(&self) -> ComResult<()> {
        Ok(())
    }
    
    // The following never get called (ideally) because MMC only calls them if
    // enable_all_extensions returns a value that is not S_OK
    fn get_first_extension(&self) -> ComResult<()> {
        Ok(())
    }
    
    fn get_next_extension(&self) -> ComResult<()> {
        Ok(())
    }
}
*/

#[repr(i32)]
#[derive(Debug, PartialEq)]
#[allow(dead_code)]
pub enum MmcNotifyType {
    Activate           = 0x8001,
    AddImages          = 0x8002,
    BtnClick           = 0x8003,
    ColumnClick        = 0x8005,
    CutOrMove          = 0x8007,
    DblClick           = 0x8008,
    Delete             = 0x8009,
    DeselectAll        = 0x800A,
    Expand             = 0x800B,
    MenuBtnClick       = 0x800D,
    Minimized          = 0x800E,
    Paste              = 0x800F,
    PropertyChange     = 0x8010,
    QueryPaste         = 0x8011,
    Refresh            = 0x8012,
    RemoveChildren     = 0x8013,
    Rename             = 0x8014,
    Select             = 0x8015,
    Show               = 0x8016,
    ViewChange         = 0x8017,
    SnapinHelp         = 0x8018,
    ContextHelp        = 0x8019,
    InitOcx            = 0x801A,
    FilterChange       = 0x801B,
    FilterBtnClick     = 0x801C,
    RestoreView        = 0x801D,
    Print              = 0x801E,
    Preload            = 0x801F,
    Listpad            = 0x8020,
    ExpandSync         = 0x8021,
    ColumnsChanged     = 0x8022,
    CanPasteOutOfProc  = 0x8023,
    //UnknownNotifyType(i32),
}

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternOutput)]
#[repr(i32)]
#[derive(Debug)]
#[allow(dead_code)]
pub enum MmcDataObjectType {
    Scope         = 0x8000,
    Result        = 0x8001,
    SnapinManager = 0x8002,
    Uninitialized = 0xffff,
}
