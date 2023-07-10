extern crate winres;

fn main() {
    let mut res = winres::WindowsResource::new();
    res.set_resource_file("resources.rc");
    res.compile().unwrap();
}

/*
fn main() {
    let mut res = winres::WindowsResource::new();
    res.set_icon("snapin.ico");
    res.compile().unwrap();
}
*/