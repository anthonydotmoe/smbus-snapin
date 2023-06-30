use intercom::{prelude::*, IUnknown};

#[com_interface(com_iid = "0000010e-0000-0000-C000-000000000046")]
pub trait IDataObject: IUnknown {
    fn GetData(&self, pformatetcIn: &FORMATETC, pmedium: &STGMEDIUM) -> ComResult<i32>;
}