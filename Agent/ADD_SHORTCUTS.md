# Add-Shortcuts Agent Prompt

This file is a self-contained prompt. **Copy the whole thing below
the `--- BEGIN PROMPT ---` line into a fresh Cursor agent chat that
has this repo open, then start asking for shortcuts** — e.g. *"add
a shortcut Ctrl+Shift+B that toggles bold on the selection"* or
*"add Alt, V, F to toggle the formula bar"*. The agent will have
all the context, conventions, and gotchas it needs.

The prompt embeds lessons learned the hard way; resist the urge to
trim it — the gotchas section in particular has saved hours of
debugging more than once.

---

--- BEGIN PROMPT ---

You are extending **mouseless-mac-excel**, a Hammerspoon (Lua)
plugin that adds Windows-style keyboard shortcuts to Microsoft
Excel for Mac. Your job is to add new shortcuts on request from
the user. Read this whole prompt before doing anything.

## Repo layout (root-flat, four files matter)

- `shortcuts.lua` — declarations. Maps key sequences/combos to
  named actions. Edit this for every new shortcut.
- `actions.lua` — Lua implementations of every action. Edit this
  for every new shortcut unless the operation is already there.
- `config.lua` — tunable parameters. New magic numbers, key lists,
  bounds, delays, on/off toggles **must** live here, never
  hardcoded in `actions.lua`.
- `README.md` — user-facing docs. Update the "Currently installed
  shortcuts" tables and (if you added a config knob) the "Tuning"
  table.

Two more files exist but you almost never touch them: `runtime.lua`
(event-tap and dispatch) and `install/hammerspoon-bootstrap.lua`
(the sample `init.lua` users drop into `~/.hammerspoon/`, written
there by `install/install.sh`).

The README has the canonical "Adding a new shortcut" recipe with
worked templates A/B/C/D. This prompt is the condensed, agent-
oriented companion to it — when in doubt, also read the README.

## Wiring model

A shortcut is the union of four things:

1. A **trigger** in `shortcuts.lua` — either a leader sequence
   (Option-tap then letters, e.g. `{ "o", "c", "w" }`) or a combo
   (`mods` + `key`).
2. An **action name** (string) — must match a function key in
   the `M` table exported by `actions.lua`.
3. An **implementation** in `actions.lua` — usually a thin wrapper
   that calls one of the helpers below.
4. **Documentation** in `README.md` — add a row to the appropriate
   "Currently installed shortcuts" table; add to the "Tuning"
   table if you introduced a config knob.

## Existing helpers in `actions.lua` (use these, don't reinvent)

- `M.send(mods, key)` — one keystroke. Last resort.
- `M.sequence(steps)` — list of `{mods, key}` with
  `config.step_delay_seconds` between each.
- `M.menu(path, is_regex)` — click a menu item by path,
  e.g. `M.menu({ "Format", "Column", "AutoFit Selection" })`.
  Pass `true` as second arg if labels vary across Excel versions.
- `M.applescript(script)` — run AppleScript. Returns `(ok, result)`;
  on failure logs the descriptor (via `hs.inspect`) and shows a
  brief alert. For AppleScript that can fail in interesting ways,
  put the `try / on error` *inside* the script (see gotcha §2).
- `M.paste_special(what)` — drives Excel's paste-special variants
  through the AppleScript dictionary. Use this — keystroke-driving
  the Paste Special dialog on macOS is unreliable.
- `M.focus_and_select_dialog_field(window_title)` — call this
  immediately after opening a dialog the user will type into. It
  clicks the dialog's text field and selects its current value so
  the user gets Windows-style overtype. See gotcha §1.

## Decision table

| The operation is… | Implementation |
| --- | --- |
| A paste-special variant | `M.paste_special(...)` one-liner |
| Anything else exposed in Excel's AppleScript dictionary | `M.applescript([[ tell application "Microsoft Excel" ... end tell ]])` |
| A menu item that just runs a command, no dialog | `M.menu({ ... })` |
| A menu item that opens a dialog where the user types | System Events `click menu item` via `M.applescript`, then `M.focus_and_select_dialog_field("Dialog Title")` |
| Nothing scripting-accessible at all | `M.send` / `M.sequence` of keystrokes (last resort) |

Prefer Excel's AppleScript dictionary, then menu navigation, then
keystrokes. **Avoid clicking UI controls by screen coordinates;
that's not portable across machines or window sizes.**

## Mac Excel quirks — read these, they will bite you

### §1. Dialog focus is a lie

Programmatically opened dialogs (via System Events menu click)
report `AXFocused = true` on their text field, but AppKit hasn't
promoted the field to first responder yet. The user's first
keystroke gets swallowed settling focus. Setting `AXFocused` or
`AXSelectedTextRange` via the AX API does **not** fix this — we
verified empirically. The only reliable nudge is a synthesised
real mouse click on the field. That's what
`M.focus_and_select_dialog_field` does (with the cursor saved and
restored so the user sees only a tiny flicker).

Don't try to "fix" this with `Escape` after the dialog opens —
on some machines Escape closes the dialog instead of focusing
the field. `dialog_focus_click` in `config.lua` is the
opt-out switch.

### §2. AppleScript error reporting is broken

`hs.osascript.applescript` returns `(false, nil, descriptor)` on
AppleScript-level failures, but the descriptor table is opaque
and frequently uninformative. We hit this debugging
`insert_sheet`: every variant returned `result=nil descriptor=table:
0x600001e0f680` with no usable message.

**The fix: catch the error inside the AppleScript itself** and
return the message as a normal string.

```applescript
tell application "Microsoft Excel"
  try
    -- the call(s) you want to make
  on error errMsg number errNum
    return "ERROR " & errNum & ": " & errMsg
  end try
end tell
```

Then in Lua, check whether `result` starts with `"ERROR "` and log
accordingly. For multi-step scripts where you don't know which
call will fail, wrap each step in its own `try / on error` block
with a labelled message (`"ERROR step2 (make new worksheet) ..."`).
`M.insert_sheet` is the reference implementation.

### §3. AppleScript can lie about its own outputs

What you write isn't always what you can read back. The Cycle
Font Color shortcut writes 16-bit RGB (`{0, 13056, 65280}`) but
Excel returns 8-bit (`{0, 51, 255}`) on read in newer versions —
across different Excel builds either is possible. If your action
needs to round-trip a value (read back to confirm what you wrote),
log raw values during dev and auto-detect the bit-width if needed.

Also covered in `README.md` under "AppleScript can lie about its
own output". Whenever a read-back value seems "off," **assume the
bridge is lying before assuming your logic is wrong**.

### §4. Menu activation can leave the ribbon focused

`M.menu()` and System Events `click menu item` both activate
inline-edit / menu-item commands correctly but sometimes leave
keyboard focus parked on Excel's ribbon strip. The user's
inline-edit cursor blinks but typing hits the ribbon.

Fix: send a single Escape ~50 ms after the menu action.

```lua
hs.timer.doAfter(0.05, function()
  hs.eventtap.keyStroke({}, "escape", 0)
end)
```

Don't shorten the delay below ~50 ms (Escape gets absorbed by the
menu dismissal animation) or extend it past ~100 ms (the user is
already typing). See `M.rename_sheet` for the canonical use.

### §5. The "cell selector freeze" after AppleScript actions

After certain AppleScript-driven changes (paste, insert sheet,
etc.) Excel parks focus on an overlay or popover that silently
eats arrow keys until dismissed. The remedy is the same Escape
nudge as §4 — 50 ms delay, single Escape — and the same one is
already wired into `M.paste_special` and `M.insert_sheet`.

### §6. Read attribute access defensively

Hammerspoon's `hs.axuielement` attribute access can throw on
unexpected element kinds. Wrap reads in `pcall` (see the `attr`
helper inside `M.focus_and_select_dialog_field`) when traversing
the AX tree.

## Workflow per shortcut

1. **Clarify the request.** Ask one question if the trigger, the
   target operation, or the expected end-state is ambiguous. Don't
   guess at colour lists, bounds, or step sizes — those belong in
   `config.lua` and the user should pick them.

2. **Pick the implementation approach** from the decision table.

3. **Implement minimally.** Write the action in `actions.lua`,
   declare the trigger in `shortcuts.lua`, add any config in
   `config.lua` with a doc-comment explaining what tuning it.

4. **Tell the user how to test.** They'll restart Hammerspoon
   themselves (menu bar → Reload Config — don't try to do this for
   them). They will run the shortcut on their machine and report
   back. Excel UI behaviour varies across machines and versions;
   testing on their actual setup is the only ground truth.

5. **Iterate with diagnostics, not guesses.** When something
   doesn't work, instrument before changing logic:
   - `_G.__mme_log("label %s", value)` writes to
     `/tmp/mouseless-mac-excel.log` (tail it from the user's
     terminal).
   - `hs.inspect(thing)` dumps tables.
   - For AppleScript: the `try / on error` pattern in §2.
   - For AX trees / unknown UI: a temporary "probe" action
     that walks the tree and writes findings to a file in
     `/tmp/` is fair game.

6. **Clean up before committing.**
   - Remove probe actions and probe shortcuts.
   - Strip diagnostic logs added for debugging.
   - Keep production-useful diagnostics (e.g. per-step
     `try / on error` blocks — they self-document and pinpoint
     future regressions).
   - Trim comments to "why," not "what." Avoid narrating
     line-by-line; explain non-obvious intent, trade-offs, and
     Mac/Excel quirks.

7. **Update the README** — at minimum the relevant "Currently
   installed shortcuts" table; also the "Tuning" table if you
   added config.

8. **Commit, don't push.** One-line commit message, sentence case,
   no trailing period, action-first. Match the existing style:
   - `Add Insert Sheet shortcut (Alt, O, W, S)`
   - `Add Cycle Font Color shortcut (Ctrl+Shift+C)`
   - `Add Row Height dialog shortcut (Alt, O, R, E)`
   - `Replace dialog Escape nudge with click-and-select helper`

   The user pushes themselves.

## Conventions

- **Never restart Hammerspoon for the user.** They control that
  via the Hammerspoon menu bar icon. Tell them to reload after
  edits.
- **No new top-level files** unless absolutely necessary —
  shortcuts go in the existing four. The README is the
  documentation surface.
- **All tunables in `config.lua`** with a doc-comment above each
  field explaining what it does, what it trades off, and any
  bounds. Cross-link from `actions.lua` only when behaviour is
  surprising (e.g. the auto-detect in `cycle_font_color`).
- **One shortcut per commit** unless the user batched the
  request.
- **Run linters mentally**: 2-space indent in Lua, no trailing
  whitespace, align `=` only in shortcut tables in `shortcuts.lua`
  (existing style).
- **No emoji** in code, comments, commit messages, or PR bodies
  unless the user explicitly asks.

## Quick smoke-test checklist before declaring done

- [ ] `shortcuts.lua` declares the new trigger, action name
      matches a function in `actions.lua`, and `desc` is human
      readable.
- [ ] The action handles the success path; failures log via
      `_G.__mme_log` and show a brief `hs.alert.show(...)` so the
      user gets a visible signal.
- [ ] Any timing constants are pulled from `config.lua` or
      justified inline.
- [ ] No probe / debug shortcuts or actions left behind.
- [ ] README tables updated; new config knob documented in
      "Tuning" if relevant.
- [ ] Comments explain *why*, not *what*.
- [ ] One-line commit, the user pushes.

If the user just says "add X" without specifying a trigger, suggest
one in line with existing patterns (leader sequence for menu
operations, combo for frequent direct ops) and let them approve.
