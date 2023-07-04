#[repr(i32)]
#[derive(Debug)]
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
}

#[repr(i32)]
#[derive(Debug)]
pub enum MmcDataObjectType {
    Scope         = 0x8000,
    Result        = 0x8001,
    SnapinManager = 0x8002,
    Uninitialized = 0xffff,
}