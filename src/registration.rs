use std::path::Path;

use winreg::enums::*;
use winreg::RegKey;

use log::{error, trace};

use crate::{CLSID_MMCSnapIn, CLSID_MMCSnapInAbout};
use crate::id::{SNAPIN_NAME, SNAPINABOUT_NAME, SNAPIN_VERSION};

fn register_snapin() -> Result<(), Box<dyn std::error::Error>> {
    trace!("crate::registration::register_snapin()");
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);
    let snapins = hklm.open_subkey("SOFTWARE\\Microsoft\\MMC\\SnapIns")?;

    let (snapin_key, _) = snapins.create_subkey(CLSID_MMCSnapIn.to_string())?;
    
    snapin_key.set_value("NameString", &SNAPIN_NAME)?;
    snapin_key.set_value("About", &CLSID_MMCSnapInAbout.to_string())?;
    snapin_key.set_value("Version", &SNAPIN_VERSION)?;
            
    snapin_key.create_subkey("StandAlone")?;

    Ok(())
}

pub fn register() -> Result<(), intercom::raw::HRESULT> {
    match register_snapin() {
        Ok(_) => {
            return Ok(());
        }
        Err(e) => {
            error!("{}", e);
            return Err(intercom::raw::E_FAIL);
        }
    }
}

fn unregister_snapin() -> Result<(), Box<dyn std::error::Error>> {
    trace!("crate::registration::unregister_snapin()");
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);
    let path = Path::new("SOFTWARE\\Microsoft\\MMC\\SnapIns").join(CLSID_MMCSnapIn.to_string());
    hklm.delete_subkey_all(path)?;

    Ok(())
}

pub fn unregister() -> Result<(), intercom::raw::HRESULT> {
    match unregister_snapin() {
        Ok(_) => {
            return Ok(());
        }
        Err(e) => {
            error!("{}", e);
            return Err(intercom::raw::E_FAIL);
        }
    }
}