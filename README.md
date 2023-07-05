# smbus-snapin

This project aims to create a MMC Snap-in in Rust. The goal is to enumerate a
list of fan controls for fans connected over the SMBus and allow the user to
view and potentially set speeds.

## Status

Currently the snap-in will load to the "Selected snap-ins" pane, but MMC
crashes when you click OK on the "Add or Remove Snap-ins" page. This is because
the MMC requests a data node for the root node of the snap in and will call
`IDataObject::GetDataHere` to set the display name of the snap-in in the scope
pane, but the current implementation returns a null pointer.