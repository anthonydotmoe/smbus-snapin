# smbus-snapin

This project aims to create a MMC Snap-in in Rust. The goal is to enumerate a
list of fan controls for fans connected over the SMBus and allow the user to
view and potentially set speeds.

## Status

Currently the snap-in will load to the "Selected snap-ins" pane, but MMC
crashes when you click OK on the "Add or Remove Snap-ins" page.