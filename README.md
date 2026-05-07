# Mouseless Mac Excel

A Hammerspoon-based plugin that brings Windows-style keyboard shortcuts
to Microsoft Excel for Mac (developed against Excel 2019, should work
on any recent Mac Excel that supports AppleScript).

Built to be edited conversationally with an AI assistant: you describe
the shortcut you want, the assistant edits `shortcuts.lua` (and
`actions.lua` if needed), Hammerspoon picks the change up automatically.
You can of course also edit by hand — it's just plain Lua tables.

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
git clone https://github.com/<your-fork-or-this-repo> mouseless-mac-excel
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

The first time the plugin runs an action that scripts Excel, macOS will
also prompt **"Hammerspoon wants to control Microsoft Excel"** — say
yes. This is a separate, per-app automation permission.

### Setting up the bridge by hand

If you don't want to run `install.sh`, you can do its job manually:

1. Open `hammerspoon-bootstrap.lua` from this clone.
2. Copy its contents into `~/.hammerspoon/init.lua` (creating the file
   if it doesn't exist; merging into your existing one if you already
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

## Adding a new shortcut

The intended workflow is: describe the shortcut to the assistant, it
edits the files, you save, Hammerspoon reloads. By hand it looks like
this:

### Add a sequence (Windows-style)

In `shortcuts.lua`, add to the `sequences` list:

```lua
{ keys = { "h", "v", "v" }, action = "paste_values", desc = "Home > Paste > Values" },
```

`keys` are the letters typed after the Option-tap leader. `action` must
match a function name in `actions.lua`. `desc` is the label shown in
on-screen alerts.

### Add a single combo

In `shortcuts.lua`, add to the `combos` list:

```lua
{ mods = { "cmd", "shift" }, key = "f", action = "paste_formulas", desc = "Paste Formulas" },
```

### Add a new action

In `actions.lua`, add a function to the `M` table. For most operations
the cleanest approach is Excel's AppleScript dictionary — there is a
helper for that:

```lua
function M.paste_formulas()
  M.applescript([[
    tell application "Microsoft Excel"
      paste special (get selection) what paste formulas
    end tell
  ]])
end
```

You can also drive things via raw keystrokes when there is no scripting
equivalent:

```lua
function M.go_to_a1()
  M.send({ "ctrl" }, "home")
end
```

Then reference it as `action = "paste_formulas"` in `shortcuts.lua`.
Save any file in this folder; Hammerspoon will reload itself.

---

## Tuning

`config.lua` exposes the dials you'll want most:

| setting | purpose |
| --- | --- |
| `excel_bundle_id` | Bundle id used to detect "Excel is frontmost". The default is `com.microsoft.Excel`. Change if your Excel install reports a different id. |
| `leader_tap_max_seconds` | Maximum tap duration to enter menu mode. Increase if your taps feel slow. |
| `menu_idle_timeout_seconds` | How long menu mode waits before silently exiting. |
| `step_delay_seconds` | Delay between scripted keystrokes inside multi-step actions. Bump if Excel dialogs feel slow. |
| `debug` | When true, shows on-screen alerts and writes a log to `/tmp/mouseless-mac-excel.log`. |

To see the live Hammerspoon console: menu-bar icon → *Console…*.

---

## Troubleshooting

- **Nothing happens when I tap Option.** Make sure Excel is the
  frontmost window. Check that Hammerspoon has Accessibility
  permission, and that you have *quit and relaunched* Hammerspoon
  since granting it. The plugin's own log lives at
  `/tmp/mouseless-mac-excel.log` — look for `accessibility=true`.
- **A sequence runs but Excel does the wrong thing.** Look at the
  action implementation in `actions.lua`. AppleScript-based actions
  are the most reliable; keystroke-based actions are sensitive to
  dialog timing — bump `step_delay_seconds` in `config.lua` if needed.
- **macOS keeps prompting for "Hammerspoon wants to control Microsoft
  Excel".** Allow it once and it will stop. If you previously denied
  it, re-enable it under *System Settings → Privacy & Security →
  Automation → Hammerspoon → Microsoft Excel*.
- **I want to disable it temporarily.** Click the Hammerspoon menu-bar
  icon → *Disable*. Re-enable to resume.
- **I want to verify what's loaded.** Tail `/tmp/mouseless-mac-excel.log`
  or open the Hammerspoon console; look for
  `[mouseless-mac-excel] started: N sequences, M combos` after a reload.

---

## License

MIT. See `LICENSE`.
