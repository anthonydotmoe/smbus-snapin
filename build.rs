extern crate winres;

fn main() {
    let mut res = winres::WindowsResource::new();
    res.set_icon("snapin.ico");
    res.compile().unwrap();
}