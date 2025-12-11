> [!NOTE]
> I prompted ChatGPT 5.1 to produce a project structure in case I want to try this again later.

If I were doing *your* project today, with what you know now, I’d aim for this shape:

* Core idea: **“MMC is just a weird GUI adapter around my own clean data model.”**
* Win32/MMC weirdness is pushed to the edges, behind types and helpers.
* Nodes are domain objects with nice APIs; they don’t personally fiddle with `HGLOBAL` or clipboard format strings.

Let me sketch what that might look like.

---

## 1. Project structure / layers

Right now everything is sort of together. I’d split conceptually into:

1. **Domain / business logic** (what you actually care about)

   * SMBus scanning, fan sensors, controllers, etc.
   * Types like `Bus`, `Device`, `Sensor`, `FanSpeed`, `Temperature`.
   * No MMC, no COM, no Win32.

2. **MMC abstraction layer**

   * Safe-ish Rusty surface that says:

     * “Here’s what a `Node` is.”
     * “Here’s how you describe what your node supports (`display_name`, `node_type`, clipboard formats).”
   * One place where you implement `IDataObject`, `IComponentData`, `IComponent`, etc.
   * Nodes from (1) are plugged into this via traits.

3. **Raw Win32 / FFI utilities**

   * Things like `HGlobalGuard`, `WideString`, clipboard format wrappers.
   * ActCtx guard you already wrote.
   * Enum conversions for MMC notify / TYMED / etc.

You already have pieces of (2) and (3) but they’re tangled. I’d lean into the separation so node code just looks like:

> “I am a node representing a fan controller; here’s how I populate my name and children.”

and **not**:

> “Let me call `GlobalLock` and manually pack UTF-16 bytes.”

---

## 2. Node model: data vs MMC plumbing

Right now your `Node` is:

```rust
pub struct Node {
    _owner: *const MMCSnapIn,
    pub node_type: NodeType,
    pcwstr_name: Option<PCWSTR>,
    pub display_name: String,
    pub hscopeitem: HSCOPEITEM,
}
```

If I were redoing this, I’d separate **data** from **MMC object**.

Something like:

```rust
/// Your domain-level node description
pub struct NodeModel {
    pub id: NodeId,
    pub kind: NodeKind,       // root, folder, device, sensor, etc.
    pub label: NodeLabel,     // nice wrapper around String + PCWSTR conversion
    pub children: Vec<NodeId>,
    // possibly: optional payload with actual device data
}

/// Thin handle used as MMC cookie/lparam
#[repr(transparent)]
#[derive(Copy, Clone, Eq, PartialEq, Hash, Debug)]
pub struct NodeId(pub isize);
```

And a registry:

```rust
pub struct NodeRegistry {
    // central authority for cookies <-> node
    nodes: HashMap<NodeId, NodeModel>,
    next_cookie: isize,
}

impl NodeRegistry {
    pub fn new() -> Self { ... }
    pub fn alloc_id(&mut self) -> NodeId { ... }
    pub fn insert(&mut self, model: NodeModel) { ... }
    pub fn get(&self, id: NodeId) -> Option<&NodeModel> { ... }
    pub fn get_mut(&mut self, id: NodeId) -> Option<&mut NodeModel> { ... }
}
```

Then your `MMCSnapIn` holds a `NodeRegistry` and the COM-specific wrappers:

```rust
pub struct MMCSnapIn {
    console: Option<ComRc<dyn IConsole2>>,
    console_namespace: Option<ComRc<dyn IConsoleNamespace>>,
    registry: NodeRegistry,
    // maybe map NodeId -> ComRc<NodeDataObject>
}
```

And you have a separate `NodeDataObject` type that implements `IDataObject` and wraps a `NodeId` + pointer back to the registry:

```rust
#[com_class(IDataObject)]
pub struct NodeDataObject {
    snapin: *const MMCSnapIn,  // same raw backpointer idea
    id: NodeId,
}
```

MMC only ever sees `NodeDataObject`. `NodeDataObject` then:

* Looks up `NodeModel` by `id` in the registry.
* Uses clean helper functions to fill `HGLOBAL`s, etc.

Your existing `Node` becomes essentially `NodeDataObject` + `NodeModel`, but with a clear boundary.

---

## 3. Display name: a type that owns the conversion

Your instinct here is exactly right:

> “I want a `String` internally, and some type that handles PCWSTR conversion for MMC.”

Something like:

```rust
pub struct NodeLabel {
    text: String,
    // cache of COM-allocated wide string
    cached: RefCell<Option<WideComString>>,
}

pub struct WideComString {
    ptr: *mut u16,
    len: usize, // including null
}

impl Drop for WideComString {
    fn drop(&mut self) {
        unsafe { CoTaskMemFree(Some(self.ptr as *const _)); }
    }
}
```

With methods:

```rust
impl NodeLabel {
    pub fn new<S: Into<String>>(s: S) -> Self {
        Self { text: s.into(), cached: RefCell::new(None) }
    }

    pub fn set<S: Into<String>>(&mut self, s: S) {
        self.text = s.into();
        *self.cached.get_mut() = None; // invalidate cache
    }

    pub fn as_pcwstr(&self) -> ComResult<PCWSTR> {
        if let Some(ref wide) = *self.cached.borrow() {
            return Ok(PCWSTR::from_raw(wide.ptr));
        }

        // build new wide string, cache it
        let utf16: Vec<u16> = self.text.encode_utf16().chain(std::iter::once(0)).collect();
        let size_bytes = utf16.len() * std::mem::size_of::<u16>();
        let buf = unsafe { CoTaskMemAlloc(size_bytes) };
        if buf.is_null() {
            return Err(ComError::E_FAIL);
        }
        unsafe {
            std::ptr::copy_nonoverlapping(utf16.as_ptr(), buf as *mut u16, utf16.len());
        }

        let wide = WideComString { ptr: buf as *mut u16, len: utf16.len() };
        let pcwstr = PCWSTR::from_raw(wide.ptr);
        *self.cached.borrow_mut() = Some(wide);
        Ok(pcwstr)
    }
}
```

Then:

* `NodeModel` just has a `NodeLabel`.
* `get_display_info` becomes:

  ```rust
  if (mask & SDI_STR) != 0 {
      let node = registry.get(id).ok_or(ComError::E_POINTER)?;
      let pcwstr = node.label.as_pcwstr()?;
      unsafe {
          (*lpscopedataitem).display_name = MMC_CALLBACK_OR_PASSTHROUGH(pcwstr);
      }
  }
  ```

No CoTaskMem calls inside the MMC glue anymore—just `label.as_pcwstr()`.

---

## 4. Centralizing `IDataObject` handling

Today, `Node` is doing all the clipboard format logic itself. I’d lift that into a reusable helper, something like:

```rust
pub trait MmcDataProvider {
    fn display_name(&self) -> &NodeLabel;
    fn node_type_guid(&self) -> &GUID;
    fn snapin_clsid(&self) -> &GUID;
    // later: add more formats as needed
}
```

And a generic “formatter” for `IDataObject` that you call from your concrete implementation:

```rust
pub struct MmcDataObjectFormatter<'a, P: MmcDataProvider> {
    provider: &'a P,
}

impl<'a, P: MmcDataProvider> MmcDataObjectFormatter<'a, P> {
    pub fn get_data_here(
        &self,
        format: ClipboardFormat,
        tymed: TagTYMED,
        medium: &mut STGMEDIUM,
    ) -> ComResult<()> {
        match format {
            ClipboardFormat::DisplayName => self.fill_display_name(tymed, medium),
            ClipboardFormat::NodeType    => self.fill_guid(tymed, medium, self.provider.node_type_guid()),
            ClipboardFormat::SnapinClsid => self.fill_guid(tymed, medium, self.provider.snapin_clsid()),
            _ => Err(DV_E_FORMATETC),
        }
    }
}
```

Where `ClipboardFormat` is a small enum you map from `cfFormat`:

```rust
pub enum ClipboardFormat {
    DisplayName,
    NodeType,
    SnapinClsid,
    Unknown(u16),
}

impl ClipboardFormat {
    pub fn from_cf(cf: u16) -> Option<Self> {
        // either use numeric IDs you cache at startup,
        // or ask GetClipboardFormatName and match on constant names once.
        ...
    }
}
```

Then your `NodeDataObject` implementation of `IDataObject` becomes:

```rust
impl IDataObject for NodeDataObject {
    fn get_data_here(&self, pformatetc: *const ComFORMATETC, pmedium: *mut ComSTGMEDIUM) -> ComResult<()> {
        let format = ClipboardFormat::from_cf(unsafe { (*pformatetc).0.cfFormat.into() })
            .ok_or(DV_E_FORMATETC)?;

        let tymed = TagTYMED::try_from(unsafe { (*pmedium).0.tymed.0 })?;

        let snapin = unsafe { &*self.snapin };
        let node = snapin.registry.get(self.id).ok_or(ComError::E_POINTER)?;
        let provider = NodeModelProvider { node, snapin }; // implements MmcDataProvider

        let formatter = MmcDataObjectFormatter { provider: &provider };
        formatter.get_data_here(format, tymed, unsafe { &mut (*pmedium).0 })
    }

    // other IDataObject methods can just be E_NOTIMPL
}
```

Now **all** the gory `GlobalLock`, `HGLOBAL`, byte-copying, etc. is abstracted away in one place; `NodeModel` and higher-level logic just provide structured values.

---

## 5. FFI helpers and types

You already wrote a nice `actctx` module. I’d build a tiny FFI-utility module too:

* `HGlobalGuard` (lock/unlock, returns `&mut [u8]` or `&mut [u16]`).
* `TagTYMED` as a safe enum with `TryFrom<i32>`.
* `ClipboardFormat` mapping.
* `global_unlock_checked` helper we talked about.
* `GuidWriter` helper that writes a `GUID` into an `HGLOBAL`.

That turns:

```rust
let ptr = GlobalLock((*pmedium).0.Anonymous.hGlobal);
...
std::ptr::copy_nonoverlapping(&CLSID_MMCSnapIn as *const _ as *const u8, ptr as *mut u8, guid_size_in_bytes);
...
GlobalUnlock(...)
```

into something like:

```rust
let mut lock = HGlobalGuard::new((*pmedium).0.Anonymous.hGlobal)?;
lock.write_guid(&CLSID_MMCSnapIn)?;
```

So again, when you go back in here six months later, you’re seeing **business logic**, not raw pointer calisthenics.

---

## 6. Path to “real” business logic (SMBus, fans, controllers)

With the above structure, plugging in actual SMBus logic becomes much easier:

* Your **domain layer** has:

  * `SmbusController`, `SmbusDevice`, `Sensor`, `Fan`, etc.
  * Some kind of discovery: `BusTopology::scan()` that returns a tree of devices/sensors.

* Your **NodeModel** becomes something like:

  ```rust
  pub enum NodeKind {
      Root,
      Controller(SmbusControllerId),
      Device(SmbusDeviceId),
      Sensor(SensorId),
      Folder,
  }
  ```

* When MMC asks for children of a node:

  * You look at the `NodeKind`:

    * If it’s `Root` → enumerate controllers from `SmbusTopology`.
    * If it’s `Controller` → enumerate devices under it.
    * If it’s `Device` → enumerate sensors under it.
  * Insert new `NodeModel`s into `NodeRegistry` as needed, using meaningful labels.

All of that logic is **domain-level**, no `HGLOBAL` needed. MMC glue just becomes the thin adapter that:

* Maps MMC cookies → `NodeId`.
* Maps clip formats → domain properties via `MmcDataProvider`.

---

## 7. What I’d actually do *first* if you revisit this

If you reopen the project and want momentum:

1. Keep your current working COM + MMC plumbing, but:

   * Fix the UB (you did).
   * Factor out a `NodeLabel` type and move the `pcwstr` logic there.
   * Change `Node` to use `NodeLabel` instead of raw `String` + ad-hoc `pcwstr()`.

2. Introduce a `NodeId` newtype for cookies and a `NodeRegistry` struct.

   * Replace `HashMap<isize, ComBox<Node>>` with `HashMap<NodeId, ComBox<Node>>`.
   * Replace uses of raw `isize` cookies with `NodeId`.

3. Make a small FFI helper module:

   * `global_unlock_checked`
   * `TagTYMED::try_from`
   * Maybe `HGlobalGuard` if you feel like it.

That alone will make the codebase feel *way* more friendly, and sets you up nicely for:

* A generic `MmcDataObjectFormatter`.
* Adding actual SMBus enumeration without drowning in COM boilerplate.

From there, you can decide how far you want to push the abstraction (full traits and helper types) versus “just enough to stop losing time to pointer bugs,” but your instincts about “good types so business logic can ignore byte twiddling” are exactly right.

> [!NOTE]
> I then brought in the issue of actions and Property Sheet pages

Yeah, you’re basically describing the natural “next layer up”:

* nodes stop caring about HGLOBAL, TYMED, cfFormats, etc.
* the MMC layer becomes “just another UI front-end” that renders:

  * **Nodes**
  * **Actions**
  * **Property pages**

Let me walk through how I’d structure that *today* so it scales to actions + property sheets without you getting destroyed by lifetimes or fat pointers again.

---

## 1. Core idea: Node + Actions + Pages as domain concepts

At the domain / snap-in layer, pretend MMC doesn’t exist.

You’d have something like:

```rust
/// Logical thing that appears in the tree / result view
pub struct NodeModel {
    pub id: NodeId,          // your cookie
    pub label: NodeLabel,    // we talked about this earlier
    pub kind: NodeKind,      // Root / Controller / Device / Sensor / Folder / ...
}

pub enum NodeKind {
    Root,
    Controller(SmbusControllerId),
    Device(SmbusDeviceId),
    Sensor(SensorId),
}

/// “What can I do with this node when the user right-clicks?”
pub struct NodeAction {
    pub id: ActionId,
    pub label: String,
    pub style: ActionStyle,  // normal / default / separator, etc.
    pub scope: ActionScope,  // scope / result / both
}

pub enum ActionStyle { Normal, Default, Separator }
pub enum ActionScope { Scope, Result, Both }

/// “What property pages does this node expose in its Properties dialog?”
pub struct NodePropertyPage {
    pub id: PageId,
    pub title: String,
    // You can store a callback or some kind of descriptor here:
    pub kind: PageKind,
}

pub enum PageKind {
    BasicInfo,
    AdvancedSettings,
    // whatever makes sense
}
```

Then define a trait that your domain implements:

```rust
pub trait NodeBehavior {
    /// context menu entries
    fn actions(&self) -> Vec<NodeAction>;

    /// property sheet pages
    fn property_pages(&self) -> Vec<NodePropertyPage>;

    /// actually execute an action
    fn invoke_action(&mut self, action_id: ActionId) -> Result<(), ActionError>;

    /// (optionally) load/save for property pages
    fn load_page(&self, page_id: PageId, sink: &mut dyn PageDataSink);
    fn save_page(&mut self, page_id: PageId, source: &dyn PageDataSource)
        -> Result<(), PageValidationError>;
}
```

**Important:** this trait has *no MMC types in it*. It’s all your model.

Your `NodeRegistry` can then give you:

```rust
impl NodeRegistry {
    pub fn get_behavior(&mut self, id: NodeId) -> &mut dyn NodeBehavior { ... }
}
```

Backed by enums or boxed trait objects internally.

---

## 2. MMC layer: “adapter” over NodeBehavior

Now the MMC layer just adapts that trait to the gross COM interfaces.

### IDataObject

You already wired this; now just use your `NodeLabel` + `NodeId`:

* `CCF_DISPLAY_NAME` → `node.label.as_pcwstr()`
* `CCF_NODETYPE`, `CCF_SNAPIN_CLSID` → GUID helper
* same pattern as before, but all the logic is centralized

```rust
impl IDataObject for NodeDataObject {
    fn get_data_here(&self, pformatetc: *const ComFORMATETC,
                     pmedium: *mut ComSTGMEDIUM) -> ComResult<()> {

        let cf = ClipboardFormat::from_formatetc(pformatetc)?;
        let tymed = TagTYMED::try_from(unsafe { (*pmedium).0.tymed.0 })?;

        let snapin = unsafe { &*self.snapin };
        let node = snapin.registry.get(self.id)
            .ok_or(ComError::E_POINTER)?;

        // now delegate to some utility that knows how to fill these formats
        MmcDataObjectFormatter::for_node(node)
            .get_data_here(cf, tymed, unsafe { &mut (*pmedium).0 })
    }
}
```

The node implementation no longer sees `HGLOBAL`, `TYMED`, etc.

---

## 3. Context menu: actions → IExtendContextMenu

You already have the right mental model:

> actions at high level; MMC layer enumerates them via IExtendContextMenu/IContextMenuCallback/IContextMenuProvider.

Rough shape:

1. **When MMC wants a context menu**, it calls your `IExtendContextMenu::AddMenuItems` with a data object and a callback.
2. You:

   * Resolve the `NodeId` from the data object.
   * Ask your `NodeBehavior` for `actions()`.
   * For each `NodeAction`, call `IContextMenuCallback::AddItem` (or `IContextMenuProvider::InsertMenuItem` depending on style).
   * At the same time, record a mapping of *MMC command ID* → `(NodeId, ActionId)` somewhere.

In Rust:

```rust
/// Held on the snap-in, shared by IComponent/IComponentData/etc.
struct CommandMap {
    // mmc_cmd_id -> (node, action)
    map: HashMap<i32, (NodeId, ActionId)>,
    next_cmd_id: i32,
}
```

In `AddMenuItems`:

```rust
fn add_menu_items(&mut self,
    data_obj: &ComItf<dyn IDataObject>,
    callback: &ComItf<dyn IContextMenuCallback>,
    insertion: *mut MENUCALLBACKDATA,
) -> ComResult<()> {

    let node_id = self.node_id_from_data_object(data_obj)?;
    let node = self.registry.get(node_id).ok_or(ComError::E_POINTER)?;
    let actions = node.behavior.actions();

    for action in actions {
        let cmd_id = self.command_map.alloc(node_id, action.id);

        // map NodeAction -> MMC C-style struct
        let item = MMC_CONTEXTMENUITEM {
            strName: to_wide(&action.label),
            nCommandID: cmd_id,
            lCommandID: cmd_id,
            // flags based on style/scope...
            ..Default::default()
        };

        callback.add_item(&item)?;
    }

    Ok(())
}
```

Then in your `IExtendContextMenu::Command` or `ICommandProvider` method:

```rust
fn command(&mut self, command_id: i32, data_obj: &ComItf<dyn IDataObject>) -> ComResult<()> {
    let (node_id, action_id) = self.command_map.resolve(command_id)
        .ok_or(ComError::E_FAIL)?;

    let behavior = self.registry.get_behavior(node_id);
    behavior.invoke_action(action_id)
        .map_err(|_| ComError::E_FAIL)
}
```

So the MMC layer becomes:

* “Menu shown?” ⇒ enumerate `NodeAction` items.
* “Command invoked?” ⇒ ask the `NodeBehavior` to actually do it.

You never stuff fat pointers into cookies; it’s all `NodeId` and `ActionId`.

---

## 4. Property sheets: NodePropertyPage → IExtendPropertySheet

Property sheets in MMC are just Win32 property sheets plus a bit of COM. Conceptually:

1. MMC calls your `IExtendPropertySheet::CreatePropertyPages(callback, handle, data_obj)`.
2. You resolve the `NodeId`, ask `NodeBehavior` for `property_pages()`.

```rust
fn create_property_pages(
    &mut self,
    callback: &ComItf<dyn IPropertySheetCallback>,
    handle: isize,
    data_obj: &ComItf<dyn IDataObject>,
) -> ComResult<()> {

    let node_id = self.node_id_from_data_object(data_obj)?;
    let node = self.registry.get(node_id).ok_or(ComError::E_POINTER)?;
    let pages = node.behavior.property_pages();

    for page in pages {
        let hpage = self.build_property_sheet_page(node_id, page.id)?;
        callback.add_page(hpage)?;
    }

    Ok(())
}
```

`build_property_sheet_page` is where `winsafe` could shine, or you can keep doing raw `PROPSHEETPAGEW` like you did in the wizard code.

The key design trick:

* **Use `lParam` in `PROPSHEETPAGE` to store a tiny context struct**, e.g.:

  ```rust
  #[repr(C)]
  struct PageCtx {
      snapin: *const MMCSnapIn,
      node_id: NodeId,
      page_id: PageId,
  }
  ```

* Allocate `Box<PageCtx>` → leak it into `lParam` so it survives as long as the page.

* In the dialog proc (or `winsafe`-style handler), cast `lParam` back to `*mut PageCtx` and:

  * On `WM_INITDIALOG` → call `NodeBehavior::load_page(page_id, sink)`.
  * On `PSN_APPLY` (OK/Apply) → call `NodeBehavior::save_page(page_id, source)`.

This keeps the per-page logic in your domain `NodeBehavior` land and leaves the dialog proc as a thin adapter that:

* Pulls data from UI → `PageDataSource`.
* Pushes data into UI ← `PageDataSink`.

If you do decide to use `winsafe` for property sheets, it should simplify some of that boilerplate (creating pages, dealing with messages) while still letting you store a `PageCtx` pointer in `lParam`.

---

## 5. Lifetimes / “Rust quirks” mitigation

You’ve already hit the big gotchas once:

* Trait objects are fat pointers → you cannot stash them in MMC cookies.
* COM objects are long-lived and moved only by pointer; Rust lifetimes don’t map directly.

The patterns above intentionally avoid the worst pain:

1. **No trait objects in cookies**
   `NodeId`, `ActionId`, `PageId` are all thin `isize`/`i32` newtypes.

2. **One authoritative registry in the snap-in root**
   `MMCSnapIn` owns a `NodeRegistry`, `CommandMap`, etc.
   Other COM objects (`IComponent`, `IDataObject`, `IExtend*` implementors) hold a raw `*const MMCSnapIn` backpointer (like you already did).

3. **Stateful domain layer behind interior mutability**
   When you need mutation from multiple COM entrypoints, stash your actual mutable domain state in `RefCell`/`RwLock`/`Mutex` inside the snap-in. From the outside, every COM call just borrows mutably from there.

4. **Minimize `unsafe` to well-audited FFI helpers**
   HGLOBAL, `GlobalLock`, `PROPSHEETPAGE`, `CreatePropertySheetPage`—all live in a small number of modules, not scattered.

Compared to “translate C++ snap-in line-by-line,” this is more “treat MMC as a weird RPC/GUI adapter for your own state machine.”

---

## 6. Realistically: how to revive this without getting stuck

If you reopen this repo and want actual progress instead of drowning in re-architecture, I’d probably:

1. Keep your existing snap-in + node + IDataObject implementation, but:

   * Keep the UB fixes you just committed.
   * Add `NodeLabel` for display name caching.
   * Add a tiny `CommandMap` and **one** hard-coded action per node, just to prove the IExtendContextMenu path.

2. After that works:

   * Introduce `NodeId` + `NodeRegistry` (if you haven’t yet).
   * Add `NodeAction` + `NodeBehavior::actions()` only for the root node.
   * Wire `IExtendContextMenu::AddMenuItems` / `Command` to `NodeBehavior::invoke_action` with something trivial like “Refresh” that just logs.

3. Then:

   * Add **one** property page for either the root or a specific node, with one text field bound to your `NodeLabel`.
   * Implement `load_page` / `save_page` for that.

Once that thin vertical slice works—context menu + property page for a single node type—you’re basically done with the hairy MMC plumbing. After that, it’s “just Rust” to bolt on SMBus enumeration and sensor logic.

And you’ll get exactly what you described: your *business logic* lives in good types and methods, and the MMC layer is a kinda ugly but contained adapter that you rarely have to think about.
