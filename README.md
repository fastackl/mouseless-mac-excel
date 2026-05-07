# Mouseless Mac Excel

A Hammerspoon-based plugin that brings Windows-style keyboard shortcuts
to Microsoft Excel for Mac (developed against Excel 2019, should work
on any recent Mac Excel that supports AppleScript).

Built to be edited conversationally with an AI assistant: you describe
the shortcut you want, the assistant edits `shortcuts.lua` (and
`actions.lua` if needed), Hammerspoon picks the change up automatically.
You can of course also edit by hand — it is just plain Lua tables.

If you are an AI agent picking this up for the first time, jump straight
to **[Adding a new shortcut](#adding-a-new-shortcut)** and the
**[Mac Excel implementation notes](#mac-excel-implementation-notes)**
below — those two sections are the playbook.

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
./install.sh
```

`install.sh` writes `~/.hammerspoon/init.lua` so Hammerspoon loads the
plugin from your clone. It backs up any existing `init.lua` first, so
it is safe to re-run. It does not modify anything else on your system.

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

If you do not want to run `install.sh`, you can do its job manually:

1. Open `hammerspoon-bootstrap.lua` from this clone.
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
├── README.md                     ← this file
├── LICENSE                       ← MIT
├── install.sh                    ← writes ~/.hammerspoon/init.lua from the template
├── hammerspoon-bootstrap.lua     ← template for ~/.hammerspoon/init.lua
├── init.lua                      ← project entry point (rarely edited)
├── config.lua                    ← tunables: leader timeout, debug, step delay
├── shortcuts.lua                 ← *** declarative shortcut table — edit this most ***
├── actions.lua                   ← action implementations (edit when adding new logic)
└── runtime.lua                   ← engine: app watcher, leader detector, sequence modal
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

**Single combos** — bound while Excel is frontmost:

| Keys | Action | Description |
| --- | --- | --- |
| `Cmd+Shift+V` | `paste_values` | Paste Values |

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

#### C — System Events click + Escape

For menu items that open a dialog the user will type into:

```lua
function M.row_height_dialog()
  M.applescript([[
    tell application "System Events"
      tell process "Microsoft Excel"
        click menu item "Height..." of menu 1 of menu item "Row" of menu 1 of menu bar item "Format" of menu bar 1
      end tell
    end tell
  ]])
  hs.timer.doAfter(0.05, function()
    hs.eventtap.keyStroke({}, "escape", 0)
  end)
end
```

The Escape ~50 ms after the click is **not** a no-op or a cancel
press — see [the focus-routing note](#focus-routing-artefact) below
for why.

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

**Fix for both:** send a single Escape via `hs.eventtap` ~50 ms after
the action.

For flavour 1, Escape dismisses the overlay and arrow keys come back;
the source's marching ants are preserved (so you can still paste
again, like normal Excel).

For flavour 2, in the half-routed state Escape is **absorbed by the
focus router rather than cancelling the dialog** — it has the side
effect of settling focus on the text field, after which keystrokes
work. Yes, this is non-obvious; yes, it works reliably; verified
manually first by the user, who reported "I hit Escape and the
dialog stays open and then accepts my keystrokes."

The 50 ms delay matters — too short and the action target hasn't
rendered; too long and Escape arrives after focus has already been
routed somewhere else, doing the wrong thing. 50 ms has been the
sweet spot on every shortcut so far.

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
- **First keystroke is eaten by my new dialog action.** Add the
  post-action Escape (see flavour 2 of the focus-routing note).
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
