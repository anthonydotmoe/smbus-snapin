use intercom::prelude::*;

mod class;
use class::*;

mod interfaces;
mod registration;
pub mod id;

use registration::{register, unregister};

com_library!(
    on_load = on_load,
    on_register = register,
    on_unregister = unregister,
    class MMCSnapInAbout,
    class MMCSnapIn,
    class Node,
);

fn on_load() {
    // Set up logging to project directory
    use log::LevelFilter;
    simple_logging::log_to_file(
        &format!("{}\\debug.log", env!("CARGO_MANIFEST_DIR")),
        LevelFilter::Trace,
    )
    .unwrap();
}
