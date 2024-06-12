use windows::core::GUID;

pub const SNAPIN_NAME: &'static str = "MMCSnapIn";
pub const SNAPIN_CLSID: GUID = GUID { data1: 0xd39d9c35, data2: 0x6106, data3: 0x4735, data4: [0xb9, 0x44, 0x7e, 0x92, 0x9d, 0x60, 0x70, 0x00]};

pub const SNAPINABOUT_NAME: &'static str = "MMCSnapInAbout";
pub const SNAPINABOUT_CLSID: GUID = GUID { data1: 0xd39d9c35, data2: 0x6106, data3: 0x4735, data4: [0xb9, 0x44, 0x7e, 0x92, 0x9d, 0x60, 0x70, 0x01]};

pub const SNAPIN_VERSION: &'static str = env!("CARGO_PKG_VERSION");