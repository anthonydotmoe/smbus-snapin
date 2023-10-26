# smbus-snapin

An MMC Snap-in implementation in Rust using [intercom](https://github.com/Rantanen/intercom)

## Goal

The project aims to create an MMC Snap-in that would enumerate a list of devices
probed on the SMBus. The user could view these devices and set parameters for
known device types, such as fan speeds, RGB, etc.

## Status

The snap-in will load into MMC after adding the required keys in the registry.
It displays the name, provider, and version strings and provides a snap-in icon
when the About page is opened. A root node is created for the snap-in and will
appear in the Snap-in Manager, Scope, and Result views without issue.

---

I also tried making this project extend **Group Policy Management** so I could
group Group Policy Objects in folders, but then I realized that I wouldn't be
able to get the proper GPO nodes for the scope pane and the result pane would
have to be redone. I did find out though that if you want to put an extension
snap-in in under the domain node, the NodeType GUID for that node is:
43e7c72d-54ab-41f6-83dc-7954f586b647