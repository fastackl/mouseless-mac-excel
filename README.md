# Mouseless Mac Excel

A Hammerspoon-based plugin that brings Windows-style keyboard shortcuts
to Microsoft Excel for Mac (developed against Excel 2019, should work
on any recent Mac Excel that supports AppleScript).

Built to be edited conversationally with an AI assistant: you describe
the shortcut you want, the assistant edits `shortcuts.lua` (and
`actions.lua` if needed), Hammerspoon picks the change up automatically.
You can of course also edit by hand — it is just plain Lua tables.

**Want to add shortcuts with an AI agent?** Open
[`ADD_SHORTCUTS.md`](ADD_SHORTCUTS.md), copy the prompt section into
a fresh Cursor chat that has this repo open, then start asking for
shortcuts. The prompt embeds all the context and gotchas the agent
needs to be productive immediately.

If you are an AI agent picking this up directly, also jump to
**[Adding a new shortcut](#adding-a-new-shortcut)** and the
**[Mac Excel implementation notes](#mac-excel-implementation-notes)**
below — those two sections are the canonical playbook.

---

## How it works

There are two kinds of shortcut, both active **only when Excel is the
frontmost application**:

1. **Sequences** (Windows-style menu chords)

   Tap and release the **Option** key alone, just like tapping **Alt**
   on Windows. A small `Excel menu` indicator appears. Then type the
   letters of the sequence. For example `e s v` runs Paste Special >
   Values, mirroring Windows `Alt, E, S, V`.

   Holding Option as a modifier still works normally (Option+E still
   types an acute accent). Only a clean tap-and-release of Option,
   with no other key pressed, enters menu mode.

2. **Single combos**

   Direct hotkeys like `Cmd+Shift+V`. Bound only while Excel is
   frontmost, so they never collide with other apps.

---

## Quick start

```bash
# 1. Install Hammerspoon (if you don't have it already)
brew install --cask hammerspoon

# 2. Clone the repo somewhere you'll keep it long-term
git clone https://github.com/fastackl/mouseless-mac-excel mouseless-mac-excel
cd mouseless-mac-excel

# 3. Wire Hammerspoon to this clone
./install/install.sh
```

`install/install.sh` writes `~/.hammerspoon/init.lua` so Hammerspoon
loads the plugin from your clone. It backs up any existing `init.lua`
first, so it is safe to re-run. It does not modify anything else on
your system.

After running it:

1. **Launch Hammerspoon** (or click its menu-bar icon → *Reload Config*
   if it is already running):

   ```bash
   open -a Hammerspoon
   ```

2. **Grant Accessibility permission** when macOS prompts. This is what
   lets Hammerspoon observe key presses and synthesise keystrokes:

   *System Settings → Privacy & Security → Accessibility → tick
   `Hammerspoon`.*

   If Hammerspoon was already running when you granted the permission,
   fully **quit and relaunch** Hammerspoon for it to take effect — the
   running process caches the old (denied) state.

3. **Optional**: in the Hammerspoon menu-bar icon, enable
   *Launch at login*.

4. Open Excel and tap-and-release the Option key. You should see a
   small `Excel menu` indicator. With something on the clipboard,
   `e s v` runs Paste Special > Values.

The first time a shortcut runs an action that scripts Excel, macOS will
also prompt **"Hammerspoon wants to control Microsoft Excel"** — say
yes. The first time a shortcut goes through System Events (e.g. the
Column Width dialog), macOS will similarly prompt for **System Events**
control. These are separate, per-app Automation permissions; allow
them once and they will not ask again.

### Setting up the bridge by hand

If you do not want to run `install/install.sh`, you can do its job
manually:

1. Open `install/hammerspoon-bootstrap.lua` from this clone.
2. Copy its contents into `~/.hammerspoon/init.lua` (creating the file
   if it does not exist; merging into your existing one if you already
   have a Hammerspoon config for other reasons).
3. Edit the `local PROJECT = "..."` line near the top so it points to
   the absolute path of your clone.
4. Reload Hammerspoon.

---

## Project layout

```
mouseless-mac-excel/
├── README.md                       ← this file
├── LICENSE                         ← MIT
├── init.lua                        ← project entry point (rarely edited)
├── config.lua                      ← tunables: leader timeout, debug, step delay
├── shortcuts.lua                   ← *** declarative shortcut table — edit this most ***
├── actions.lua                     ← action implementations (edit when adding new logic)
├── runtime.lua                     ← engine: app watcher, leader detector, sequence modal
├── install/
│   ├── install.sh                  ← writes ~/.hammerspoon/init.lua from the template
│   └── hammerspoon-bootstrap.lua   ← template for ~/.hammerspoon/init.lua
└── Agent/
    └── ADD_SHORTCUTS.md            ← copy-paste prompt for AI-assisted shortcut work
```

`~/.hammerspoon/init.lua` is a thin bootstrap that points Hammerspoon
at this folder and reloads the config whenever any `.lua` file here is
saved. Treat the cloned folder as the source of truth.

---

## Currently installed shortcuts

`shortcuts.lua` is the actual source of truth; this table is a
human-readable mirror that should be updated whenever you add or
remove a shortcut.

**Sequences** — tap Option, then type the letters:

| Keys | Action | Description |
| --- | --- | --- |
| `e s v` | `paste_values` | Paste Special > Values |
| `e s t` | `paste_formats` | Paste Special > Formats |
| `e s w` | `paste_column_widths` | Paste Special > Column Widths |
| `o c w` | `column_width_dialog` | Format > Column > Width… (opens dialog) |
| `o r e` | `row_height_dialog` | Format > Row > Height… (opens dialog) |
| `o h r` | `rename_sheet` | Format > Sheet > Rename (inline edit on the sheet tab) |
| `o h m` | `move_sheet_dialog` | Edit > Sheet > Move or Copy… (opens dialog) |
| `o w s` | `insert_sheet` | Insert a new worksheet immediately after the active sheet |
| `e l` | `delete_sheet` | Delete the active worksheet (Excel shows its native confirmation dialog first) |

**Single combos** — bound while Excel is frontmost:

| Keys | Action | Description |
| --- | --- | --- |
| `Cmd+Shift+V` | `paste_values` | Paste Values |
| `Ctrl+Shift+C` | `cycle_font_color` | Cycle font color through `config.font_color_cycle` |
| `Ctrl+Shift+I` | `zoom_in` | Zoom in by `config.zoom_step` (clamped to `zoom_max`) |
| `Ctrl+Shift+J` | `zoom_out` | Zoom out by `config.zoom_step` (clamped to `zoom_min`) |
| `Ctrl+Shift+Space` | `select_column` | Select the entire column(s) covered by the current selection |

---

## Adding a new shortcut

The recipe in five steps. The first time through, read the
[implementation notes](#mac-excel-implementation-notes) below before
writing the action — they explain *why* certain patterns exist.

### 1. Pick an implementation approach

Match the operation against this decision table; each row links to a
template in step 2.

| The operation is… | Use approach |
| --- | --- |
| A copy/paste/cell op exposed by Excel's AppleScript dictionary (paste variants, sheet ops, range manipulation) | **A — Excel AppleScript** |
| A menu item that just runs a command and does not open a dialog | **B — `M.menu()`** |
| A menu item that opens a dialog where the user types | **C — System Events click + Escape** |
| Anything with no scripting equivalent at all | **D — keystroke synthesis** (last resort) |

### 2. Write the action in `actions.lua`

#### A — Excel AppleScript

For paste-special variants, just delegate to the existing
`paste_special` helper (it already wires in the focus-restoring
Escape — see notes):

```lua
function M.paste_formulas() M.paste_special("paste formulas") end
```

For other Excel scripting operations, call `M.applescript()` directly:

```lua
function M.go_to_first_sheet()
  M.applescript([[
    tell application "Microsoft Excel"
      activate object first sheet of active workbook
    end tell
  ]])
end
```

#### B — `M.menu()`

For menu items that just *do something* (no dialog opens, no
follow-up text input):

```lua
function M.autofit_columns()
  M.menu({ "Format", "Column", "AutoFit Selection" })
end
```

If you are not sure of the exact menu label on this Excel version,
**probe the menu tree first** (see implementation notes).

#### C — System Events click + `focus_and_select_dialog_field`

For menu items that open a dialog the user will type into. Use the
helper rather than re-implementing the focus dance — see
[the focus-routing note](#focus-routing-artefact) for why.

```lua
function M.row_height_dialog()
  M.applescript([[
    tell application "System Events"
      tell process "Microsoft Excel"
        click menu item "Height..." of menu 1 of menu item "Row" of menu 1 of menu bar item "Format" of menu bar 1
      end tell
    end tell
  ]])
  M.focus_and_select_dialog_field("Row Height")
end
```

`M.focus_and_select_dialog_field(window_title)`:

- Waits `config.dialog_focus_click_delay_seconds` for the dialog to
  render.
- Locates the dialog by its window title.
- Finds its first `AXTextField`.
- Synthesises a mouse click on the field (with mouse position saved
  and restored so the visible cursor only flickers).
- Selects the field's existing value so the user can overtype.
- Is a no-op when `config.dialog_focus_click = false`.
- Logs to `/tmp/mouseless-mac-excel.log` on failure (dialog not
  found, field not found, etc.) and otherwise stays silent.

If a future dialog has multiple text fields we'll extend the helper
to take an index or selector. For now it picks the first.

#### D — Keystroke synthesis

Only when no scripting path exists. `M.send(mods, key)` is a single
key tap; `M.sequence(steps)` is several taps with `step_delay_seconds`
between them.

```lua
function M.go_to_a1()
  M.send({ "ctrl" }, "home")
end
```

### 3. Register the trigger in `shortcuts.lua`

Add to the `sequences` list (Windows-style chord) or the `combos`
list (single hotkey):

```lua
-- sequences:
{ keys = { "o", "r", "h" }, action = "row_height_dialog", desc = "Format > Row > Height..." },

-- combos:
{ mods = { "cmd", "shift" }, key = "f", action = "paste_formulas", desc = "Paste Formulas" },
```

`action` must match a function name on `M` in `actions.lua`. `desc`
shows up in the on-screen alerts and the log line.

### 4. Save and let Hammerspoon auto-reload

The bootstrap watches all `.lua` files in this folder and triggers
`hs.reload()` on save. You should see a brief `Reloading Mouseless
Mac Excel` alert. Confirm in the log:

```bash
tail -3 /tmp/mouseless-mac-excel.log
# expect:
#   started: <N> sequences, <M> combos
#   loaded ok; accessibility=true
```

If the count went up by what you expected, you are wired up.

### 5. Test in Excel and check focus

Trigger the shortcut in Excel and verify three things:

1. **Action ran.** The on-screen `desc` alert appears, and the
   spreadsheet reflects the change.
2. **Selector still moves with arrow keys.** If arrows feel locked,
   you have hit the focus-routing artefact — the action needs the
   post-action Escape (see notes). Pressing Escape manually should
   unlock arrows; that confirms the diagnosis.
3. **For dialog actions, the first keystroke registers.** Type a
   digit immediately after the dialog opens. If the first digit is
   eaten, the dialog text field has not been promoted to first
   responder yet — same fix.

If the action errored, the error message flashes on screen (often
truncated) — the full text is in `/tmp/mouseless-mac-excel.log`.

---

## Mac Excel implementation notes

The hard-won knowledge for anyone building shortcuts here. Read
before writing actions.

### AppleScript beats keystrokes for Excel-internal operations

The original Paste Special > Values action drove the dialog with
synthesised keystrokes (`Ctrl+Cmd+V`, then `V`, then `Return`). It
did not work: macOS dialogs do **not** honour the Windows-style
underlined-letter accelerators that select radio buttons. The `V`
keystroke had no effect on the Values radio, so `Return` confirmed
the dialog with its default selection.

**Lesson:** if Excel exposes the operation in its AppleScript
dictionary (`paste special`, `select`, `copy`, `clear`, etc.), use
that. It is faster, dialog-free, and doesn't depend on timing.

### AppleScript can lie about its own output — verify round-trips

Mac Excel's AppleScript bridge is **not symmetric** about value
formats: the units it accepts on write are not always the units
it returns on read. The canonical example is font colour:

- **Writing** `set color of font object of selection to {1028, 13107, 65535}`
  paints the cell blue (`#0433FF`). 16-bit RGB, as the standard
  AppleScript `RGB color` type expects.
- **Reading** the *same* cell's `color of font object` returns
  `{4, 51, 255}`. Same colour, but in 8-bit units this time. Read
  it back with the natural assumption that it's 16-bit and you'll
  get a hex value of `000001` — nothing like the `0433FF` you
  wrote — and any "did the user pick a colour I know about?" logic
  silently breaks.

This bit us in `M.cycle_font_color`: the first press cycled black
→ blue correctly, but every subsequent press kept the cell on blue
because the read-back colour never matched any cycle entry.

**Defensive pattern for any action that reads then writes the same
Excel property:**

1. Don't trust the docs (or your past experience) about units —
   the AppleScript dictionary doesn't expose them and behaviour
   has shifted across Excel versions.
2. Always log the raw read once during development. The shape of
   the data is the spec.
3. Where it's cheap to do so, **auto-detect** the units on read
   rather than hardcode them. For RGB the detection is trivial
   ("does any component exceed 255?"); for other property types
   the detection might be different but the principle stands.

This kind of asymmetric round-tripping shows up elsewhere too — be
suspicious whenever you read a property you previously wrote and
the value seems "off." It's almost certainly the bridge, not you.

### Extracting AppleScript errors

When an AppleScript snippet fails, `hs.osascript.applescript` gives
you back `(false, nil, descriptor)` where the descriptor is an opaque
table. `M.applescript` runs `hs.inspect` on it for the log, but the
keys haven't been stable across Hammerspoon versions and the dump
can still be uninformative. We hit this trying to debug
`make new worksheet`, which kept failing with bare `result=nil`.

The reliable pattern: **catch the error inside the AppleScript
itself** and return it as a regular string through the normal
result channel.

```applescript
tell application "Microsoft Excel"
  try
    -- the call(s) you want to make
  on error errMsg number errNum
    return "ERROR " & errNum & ": " & errMsg
  end try
end tell
```

The Lua side then checks whether `result` starts with `"ERROR "`
and logs the message. For multi-step scripts where you don't know
which call fails, wrap each step in its own `try / on error`
block returning a labelled message (e.g. `"ERROR step2 (make new
worksheet) -50: Parameter error"`) — `M.insert_sheet` uses this
pattern and is the reference implementation.

When in doubt, reach for this technique. It's strictly more
informative than relying on the descriptor, and the AppleScript
overhead is negligible.

### Focus-routing artefact

This shows up in two distinct flavours; both are fixed by the same
trick.

**Flavour 1 — after a programmatic paste.** Excel's "Paste Options"
overlay (the small clipboard icon at the bottom-right of the pasted
region) silently captures keyboard focus. Until something dismisses
it, arrow keys do nothing on the worksheet.

**Flavour 2 — after programmatically opening a dialog.** The dialog
window appears, but its text field has not been promoted to "first
responder", so the user's first keystroke is lost.

The two flavours have different fixes — one is unconditional and
lives inline in the paste actions, the other is centralised in a
helper and configurable.

**Flavour 1 (paste overlay): unconditional post-paste Escape.**
After an AppleScript-driven paste, send a single Escape via
`hs.eventtap` ~50 ms after the action. Escape dismisses the overlay
and arrow keys come back; the source's marching ants are preserved
(so you can still paste again, like normal Excel). This shows up
on every Mac/Excel combination tested, so the Escape is unconditional
inside `M.paste_special` and not exposed as a config knob.

**Flavour 2 (dialog text field): synthesised mouse click via
`M.focus_and_select_dialog_field()`.** We tried two approaches before
landing on the click:

1. *Post-open Escape* — works on some machines (Escape is absorbed
   by the focus router and settles the field), but on others the
   dialog renders fast enough that Escape arrives at a fully-focused
   dialog and just closes it. Unreliable across hardware.
2. *AX-level focus / selection* — setting `AXFocused = true` and
   `AXSelectedTextRange` on the field via `hs.axuielement`. We
   verified empirically that this updates the AX tree (the field
   reports as focused and the value as selected) but **does not
   actually promote the field in AppKit's first-responder chain**.
   The first keystroke still gets eaten.

The only reliable nudge we found is a **real synthesised mouse
click** at the field's screen frame. A real click bypasses the
AX→AppKit indirection and forces first-responder the same way a
hardware click would. We then select the existing value with two
mechanisms in sequence: first an AX `AXSelectedTextRange` write
(best-effort; on some machines it drives the visible selection,
on others the AX tree records the range but the field editor stays
cursor-only), then a synthesised `Cmd+A` through the real keyboard
pipeline (definitive — the click has made the field the first
responder, so `Cmd+A` hits the field editor and visibly highlights
the existing value). The result is Windows-style overtype: open
dialog, type new value, press Enter. The user's mouse cursor is
saved and restored around the click so it doesn't end up parked
on the dialog.

Gated on `dialog_focus_click` (default `true`), with the open→click
delay tunable via `dialog_focus_click_delay_seconds` (default 80 ms).
Dialog-opening actions should call `M.focus_and_select_dialog_field(window_title)`
rather than re-implementing this — the helper is the only place we
maintain this knowledge.

### Marching ants are preserved on purpose

Do **not** include `set cut copy mode to false` in paste actions. The
Windows-Excel UX is that copy mode stays armed after a paste so you
can paste the same source to multiple targets. We replicate that.
The user explicitly confirmed this is the desired behaviour.

### `M.menu()` vs System Events click — when to use which

| Need | Use |
| --- | --- |
| Invoke a menu item that just runs a command (no dialog, no follow-up keyboard input) | `M.menu({"...", "..."})` (Hammerspoon's `selectMenuItem`) |
| Invoke a menu item that opens a dialog with text input | `tell application "System Events" ... click menu item ...` |

Both invoke menus through the macOS accessibility tree, but the
focus-routing artefact described above shows up most aggressively on
dialogs opened via `selectMenuItem`. The System Events `click` form
mimics a real mouse click closely enough that — combined with the
post-action Escape — focus settles cleanly. On non-dialog menu items
either form works; we use `M.menu()` because it is shorter.

### Probing unknown menu paths

When you do not know the exact label of a menu item (Mac Excel uses
plain ASCII `...` in some menus, Unicode `…` in others, and labels
sometimes change between versions), **dump the menu tree first**
rather than guessing. Temporarily replace your action body with:

```lua
function M.<your_action>()
  local app = hs.application.find(config.excel_bundle_id)
  if not app then return end
  local items = app:getMenuItems()
  if not items then return end
  local lines = {}
  local function walk(node, depth)
    for _, mi in ipairs(node or {}) do
      lines[#lines + 1] = string.rep("  ", depth) .. "- " .. (mi.AXTitle or "(no title)")
      local children = mi.AXChildren
      if children and children[1] then walk(children[1], depth + 1) end
    end
  end
  walk(items, 0)
  local f = io.open("/tmp/mle-menus.txt", "w")
  if f then f:write(table.concat(lines, "\n")); f:close() end
end
```

Trigger it once, read `/tmp/mle-menus.txt`, find the literal label,
restore the real implementation. Use exact-string matching when the
label is fixed; reach for the regex form of `selectMenuItem` only
when the label genuinely varies across versions.

### `M.menu` boolean coercion

Hammerspoon's `selectMenuItem(path, isRegex)` rejects a literal
`nil` for the second argument — it requires a real boolean. The
`M.menu` helper coerces with `is_regex == true`. If you ever build
a similar helper, make sure to do the same coercion or you will
get a mysterious "incorrect type 'nil' for argument 3" alert.

### AppleScript runs in-process inside Hammerspoon

`hs.osascript.applescript()` executes via NSAppleScript inside the
Hammerspoon process, which means System Events `tell process …`
blocks inherit Hammerspoon's Accessibility permission. You do **not**
need to grant osascript its own Accessibility — you do, however,
need to allow Hammerspoon → System Events the first time it asks.

### Watch the log when debugging

Every interesting runtime event lands in `/tmp/mouseless-mac-excel.log`:

- `bootstrap start` and `started: N sequences, M combos` on every
  reload.
- `accessibility=true|false` immediately after.
- `Excel focused; bindings active` / `Excel blurred; bindings inactive`.
- Any action error, with file and line number.
- `menu not found: ...` when `M.menu` cannot resolve a path.

Tail it during development:

```bash
tail -f /tmp/mouseless-mac-excel.log
```

---

## Tuning

`config.lua` exposes the dials you'll want most:

| setting | purpose |
| --- | --- |
| `excel_bundle_id` | Bundle id used to detect "Excel is frontmost". The default is `com.microsoft.Excel`. Change if your Excel install reports a different id. |
| `leader_tap_max_seconds` | Maximum tap duration to enter menu mode. Increase if your taps feel slow. |
| `menu_idle_timeout_seconds` | How long menu mode waits before silently exiting. |
| `step_delay_seconds` | Delay between scripted keystrokes inside multi-step `M.sequence()` actions. Bump if Excel dialogs feel slow. |
| `dialog_focus_click` | When `true` (default), dialog-opening actions synthesise a mouse click on the dialog's text field and select its existing value so typing replaces it (Windows-Excel-style overtype). The cursor briefly flickers to the field and back. Turn off to keep the cursor undisturbed and click into dialogs manually. See the [focus-routing note](#focus-routing-artefact). |
| `dialog_focus_click_delay_seconds` | Delay (seconds) between opening a dialog and looking it up in the AX tree to click. Default `0.08`. Bump if the log shows `focus_and_select: dialog "..." not found`. |
| `font_color_cycle` | Ordered list of 6-character hex strings the `cycle_font_color` action walks through. Edit to taste — order is significant. Leading `#` tolerated, case-insensitive. |
| `zoom_min` / `zoom_max` / `zoom_step` | Lower bound, upper bound, and step (all integer percentages) for `zoom_in` / `zoom_out`. Defaults: `50` / `200` / `10`. Snap-to-step: zoom-in from 117% lands on 120%, from 120% lands on 130%. |
| `debug` | When true, shows on-screen alerts and writes `/tmp/mouseless-mac-excel.log`. |

To see the live Hammerspoon console: menu-bar icon → *Console…*.

---

## Troubleshooting

- **Nothing happens when I tap Option.** Make sure Excel is the
  frontmost window. Check that Hammerspoon has Accessibility
  permission, and that you have *quit and relaunched* Hammerspoon
  since granting it. Look for `accessibility=true` in
  `/tmp/mouseless-mac-excel.log`.
- **A sequence runs but Excel does the wrong thing.** Open
  `actions.lua` and re-read the [implementation notes](#mac-excel-implementation-notes)
  — most regressions are about the focus-routing Escape, the marching
  ants, or a wrong menu path. The full error text is always in the log.
- **Selector locks (arrow keys do nothing) after my new action.**
  Add the post-action Escape (see flavour 1 of the focus-routing note).
- **First keystroke is eaten by my new dialog action.** Make sure
  the action calls `M.focus_and_select_dialog_field("<Window Title>")`
  after opening the dialog, and that `config.dialog_focus_click` is
  `true` (it is by default). Tail the log: if you see
  `focus_and_select: dialog "..." not found`, the dialog hadn't
  rendered yet when we went looking — bump
  `config.dialog_focus_click_delay_seconds`.
- **Selected value isn't replaced when I type.** That's the same
  helper not having found the field. Same diagnosis as above; check
  the log for `focus_and_select: AXTextField not found in ...`.
- **Cursor visibly flickers on every dialog and I'd rather it didn't.**
  Set `dialog_focus_click = false`. You'll lose overtype UX and the
  first keystroke into the dialog will be eaten — you'll need to
  click into the field manually before typing.
- **macOS keeps prompting for "Hammerspoon wants to control Microsoft
  Excel" or "...System Events".** Allow it once and it will stop. If
  you previously denied it, re-enable it under *System Settings →
  Privacy & Security → Automation → Hammerspoon → \<target\>*.
- **I want to disable it temporarily.** Click the Hammerspoon menu-bar
  icon → *Disable*. Re-enable to resume.
- **I want to verify what's loaded.** Tail `/tmp/mouseless-mac-excel.log`
  or open the Hammerspoon console; look for
  `[mouseless-mac-excel] started: N sequences, M combos` after a reload.

---

## License

MIT. See `LICENSE`.
