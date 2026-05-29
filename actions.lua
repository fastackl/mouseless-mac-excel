-- Action implementations.
--
-- Each action is a Lua function. Reference it from shortcuts.lua by its
-- key name (e.g. action = "paste_values").
--
-- Helpers available:
--   M.send(mods, key)         -- one keystroke (mods is a list, e.g. {"cmd"})
--   M.sequence(steps)         -- list of {mods, key} pairs with config.step_delay between them
--   M.menu(path)              -- click a menu item by path, e.g. {"Edit", "Paste Special..."}
--
-- See README.md for the full add-a-shortcut workflow.

local config = require("config")

local M = {}

----------------------------------------------------------------------
-- Low-level helpers
----------------------------------------------------------------------

function M.send(mods, key)
  hs.eventtap.keyStroke(mods or {}, key, 0)
end

-- Run a list of {mods, key} keystrokes with a small delay between them.
-- The delay is needed because Excel dialogs take a moment to appear.
function M.sequence(steps)
  local delay = config.step_delay_seconds
  for i, step in ipairs(steps) do
    hs.timer.doAfter((i - 1) * delay, function()
      hs.eventtap.keyStroke(step[1] or {}, step[2], 0)
    end)
  end
end

-- Invoke an Excel menu item by path. When `is_regex` is true each path
-- component is treated as a Lua pattern, which is useful when a menu
-- label varies across Excel versions (e.g. trailing "..." vs the
-- single Unicode ellipsis "…"). Logs a message on failure so we can
-- see what went wrong from /tmp/mouseless-mac-excel.log.
function M.menu(path, is_regex)
  local app = hs.application.find(config.excel_bundle_id)
  if not app then
    if _G.__mme_log then _G.__mme_log("menu: Excel not found") end
    return false
  end
  -- Hammerspoon's selectMenuItem rejects a literal nil for arg 2; it
  -- requires a real boolean. Coerce to false when omitted.
  local ok = app:selectMenuItem(path, is_regex == true)
  if not ok and _G.__mme_log then
    _G.__mme_log("menu not found: " .. table.concat(path, " > "))
  end
  return ok
end

----------------------------------------------------------------------
-- Excel actions
----------------------------------------------------------------------

-- Run an AppleScript snippet against Excel and surface failures clearly.
-- Returns true on success; logs and shows an alert on failure.
--
-- hs.osascript.applescript returns three values: (ok, result, descriptor).
-- On AppleScript-level errors the `result` is nil and the actual error
-- text lives in the descriptor table under `NSAppleScriptErrorMessage`
-- (and the error number under `NSAppleScriptErrorNumber`). We pull
-- those out explicitly so the log entry tells us what went wrong.
function M.applescript(script)
  local ok, result, descriptor = hs.osascript.applescript(script)
  if not ok then
    -- The descriptor returned by hs.osascript.applescript on
    -- failure is an opaque table whose key naming we haven't
    -- reverse-engineered. hs.inspect dumps the whole structure so
    -- whatever's in there ends up in the log; for AppleScript
    -- errors where we can write the catch ourselves, prefer
    -- in-script try/on error (see M.insert_sheet for the pattern).
    local detail
    if type(descriptor) == "table" and hs.inspect then
      detail = hs.inspect(descriptor)
    else
      detail = string.format("result=%s descriptor=%s",
        tostring(result), tostring(descriptor))
    end
    if _G.__mme_log then
      _G.__mme_log("applescript error: %s", detail)
    end
    hs.alert.show("AppleScript error (see log)", 1.2)
  end
  return ok, result
end

-- Paste Special variants.
--
-- Driving the Mac Excel Paste Special dialog with synthetic keystrokes
-- is unreliable: macOS dialog radio buttons don't honour Windows-style
-- letter accelerators. Excel's AppleScript dictionary exposes each
-- paste variant directly via the `what` parameter, so we call it
-- instead. No dialog, no timing dance.
--
-- Post-paste Escape:
--
-- After an AppleScript-driven paste, Excel's "Paste Options" overlay
-- (the small clipboard icon at the bottom-right of the pasted region)
-- silently captures keyboard focus. Until something dismisses the
-- overlay, arrow keys do nothing on the worksheet. A single Escape
-- dismisses *just* the overlay; the source's marching-ants copy
-- state is preserved, so the user can paste again, matching the
-- standard Excel copy/paste UX where the source stays "armed" until
-- the user explicitly cancels with Escape on the worksheet itself.
--
-- The 50 ms timer gives Excel a moment to render the overlay before
-- we send Escape; without the delay, the keystroke can race the
-- overlay and arrive while focus is still on the worksheet, in
-- which case Escape would clear the marching ants instead of
-- dismissing the overlay.
function M.paste_special(what)
  M.applescript(string.format([[
    tell application "Microsoft Excel"
      paste special (get selection) what %s
    end tell
  ]], what))
  hs.timer.doAfter(0.05, function()
    hs.eventtap.keyStroke({}, "escape", 0)
  end)
end

function M.paste_values()        M.paste_special("paste values")        end
function M.paste_formats()       M.paste_special("paste formats")       end
function M.paste_column_widths() M.paste_special("paste column widths") end

----------------------------------------------------------------------
-- Dialog helpers
----------------------------------------------------------------------

-- Bring keyboard focus to a freshly opened dialog's first text field
-- and select its current value, giving the user Windows-style overtype:
-- open dialog, type new value, press Enter to confirm.
--
-- Why this exists:
--   When a dialog is opened programmatically (e.g. via System Events
--   menu click), AX reports the text field as AXFocused=true but
--   AppKit hasn't yet promoted it to first responder, so the user's
--   first keystroke is eaten settling focus. Setting AXFocused or
--   AXSelectedTextRange via AX has no effect on the responder chain —
--   we verified this empirically. The only reliable nudge is a real
--   synthesised mouse click on the field, which forces AppKit to
--   update first-responder the same way a hardware click would.
--
-- After the click, we set AXSelectedTextRange to cover the full
-- current value so that typing replaces it rather than appending.
--
-- The user's mouse cursor is saved before the click and restored
-- immediately after, so they don't end up parked on the dialog.
--
-- Opt-out via config.dialog_focus_click = false.
function M.focus_and_select_dialog_field(window_title)
  if not config.dialog_focus_click then return end

  hs.timer.doAfter(config.dialog_focus_click_delay_seconds, function()
    local ok, err = pcall(function()
      local ax = hs.axuielement
      if not ax then
        local ok_req, mod = pcall(require, "hs.axuielement")
        if ok_req then ax = mod end
      end
      if not ax then
        if _G.__mme_log then
          _G.__mme_log("focus_and_select: hs.axuielement unavailable")
        end
        return
      end

      local app = hs.application.find(config.excel_bundle_id)
      if not app then return end

      local app_ax = ax.applicationElement(app)
      if not app_ax then return end

      -- Defensive attribute reader: AX attribute access can throw on
      -- unexpected element kinds.
      local function attr(elem, name)
        local got, val = pcall(function() return elem[name] end)
        return got and val or nil
      end

      -- Locate the dialog window by title.
      local dialog
      for _, w in ipairs(attr(app_ax, "AXWindows") or {}) do
        if attr(w, "AXTitle") == window_title then
          dialog = w
          break
        end
      end
      if not dialog then
        if _G.__mme_log then
          _G.__mme_log("focus_and_select: dialog %q not found (try bumping config.dialog_focus_click_delay_seconds)",
            tostring(window_title))
        end
        return
      end

      -- Find the first AXTextField in the dialog. All current
      -- single-field dialogs (Column Width, Row Height, ...) have
      -- exactly one; if a future dialog has several we'll extend
      -- this helper to take an index/selector.
      local field
      for _, child in ipairs(attr(dialog, "AXChildren") or {}) do
        if attr(child, "AXRole") == "AXTextField" then
          field = child
          break
        end
      end
      if not field then
        if _G.__mme_log then
          _G.__mme_log("focus_and_select: AXTextField not found in %q",
            tostring(window_title))
        end
        return
      end

      -- Resolve the field's screen frame. AXFrame is the common case;
      -- AXPosition + AXSize is a safety fallback for older AX bridges.
      local frame = attr(field, "AXFrame")
      if not frame then
        local pos  = attr(field, "AXPosition")
        local size = attr(field, "AXSize")
        if pos and size then
          frame = { x = pos.x, y = pos.y, w = size.w, h = size.h }
        end
      end
      if not frame then
        if _G.__mme_log then
          _G.__mme_log("focus_and_select: field frame unavailable in %q",
            tostring(window_title))
        end
        return
      end

      local fw = frame.w or frame.width or 0
      local fh = frame.h or frame.height or 0
      local click_point = { x = frame.x + fw / 2, y = frame.y + fh / 2 }

      local original_mouse = hs.mouse.absolutePosition()
      hs.eventtap.leftClick(click_point)
      hs.mouse.absolutePosition(original_mouse)

      -- Brief delay so the click has time to traverse AppKit's
      -- responder chain before we select the field's contents.
      -- This is an AppKit internal timing constant, not user-tunable.
      hs.timer.doAfter(0.03, function()
        -- Best-effort AX-level select-all. On some Mac/Excel
        -- combinations this drives the field editor's visible
        -- selection; on others the AX tree records the range but
        -- the field editor stays cursor-only. We do it anyway, then
        -- have Cmd+A below for the cases where it doesn't take.
        pcall(function()
          local value = tostring(attr(field, "AXValue") or "")
          field:setAttributeValue("AXSelectedTextRange", { loc = 0, len = #value })
        end)
        -- Definitive select-all via the real keyboard pipeline.
        -- The click above has made the field the first responder,
        -- so Cmd+A hits the field editor and visibly highlights
        -- the existing value — typing now replaces it.
        hs.eventtap.keyStroke({ "cmd" }, "a", 0)
      end)
    end)

    if not ok and _G.__mme_log then
      _G.__mme_log("focus_and_select error: %s", tostring(err))
    end
  end)
end

-- After a System Events menu-bar click, macOS leaves keyboard focus
-- parked on Excel's ribbon strip even though the downstream control
-- (inline-edit cursor, modal dialog, ...) is the logical target.
-- The user's first keystroke hits the ribbon instead.
--
-- The fix is a single Escape ~50ms after the click: Escape clears
-- the ribbon focus, and the downstream control inherits first
-- responder on its own (verified empirically across rename_sheet
-- and delete_sheet — both inline-edit and modal-alert flavours).
--
-- For dialog cases:
--   At 50ms the alert dialog usually hasn't fully rendered yet, so
--   the Escape doesn't risk cancelling it; it lands on the ribbon
--   (which still has first responder) and clears that. By the time
--   the dialog finishes opening the responder chain is clean and
--   it grabs first responder on its own. We tried AX-locating the
--   dialog and synthesising a mouse-click on its top-centre — the
--   click was insufficient to promote first responder for
--   button-only dialogs (worked for text-field dialogs because the
--   click landed on a focusable control).
--
-- Don't shorten below ~50ms (Escape gets absorbed by the menu
-- dismissal animation) or extend past ~100ms (the user could
-- already be typing, and for genuinely-focused alerts the Escape
-- would cancel).
function M.dismiss_menu_focus()
  hs.timer.doAfter(0.05, function()
    hs.eventtap.keyStroke({}, "escape", 0)
  end)
end

----------------------------------------------------------------------
-- Format menu actions
----------------------------------------------------------------------

-- Open Format > Column > Width... and put the user in position to
-- type a new width: the existing value is selected, so the first
-- keystroke replaces it.
function M.column_width_dialog()
  M.applescript([[
    tell application "System Events"
      tell process "Microsoft Excel"
        click menu item "Width..." of menu 1 of menu item "Column" of menu 1 of menu bar item "Format" of menu bar 1
      end tell
    end tell
  ]])
  M.focus_and_select_dialog_field("Column Width")
end

-- Row equivalent of column_width_dialog: opens Format > Row > Height...
-- with the existing height pre-selected for overtype.
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

-- Rename the active sheet. Format > Sheet > Rename in Mac Excel
-- switches the active sheet's tab into inline-edit mode (no dialog
-- pops) — the user types the new name and presses Enter.
--
-- The trailing M.dismiss_menu_focus() clears Excel's
-- ribbon-retained focus so the inline-edit cursor actually
-- receives the user's keystrokes. See the helper for the why and
-- timing notes.
function M.rename_sheet()
  M.applescript([[
    tell application "System Events"
      tell process "Microsoft Excel"
        click menu item "Rename" of menu 1 of menu item "Sheet" of menu 1 of menu bar item "Format" of menu bar 1
      end tell
    end tell
  ]])
  M.dismiss_menu_focus()
end

-- Open Edit > Sheet > Move or Copy Sheet... — Excel's native
-- "Move or Copy" dialog where the user picks a destination workbook
-- and target position. The dialog's controls are dropdowns and
-- checkboxes (no text field for typing), so we don't need
-- M.focus_and_select_dialog_field — once the dialog is visible the
-- user navigates it with Tab and arrow keys natively.
--
-- Path note: Windows Excel keeps Move or Copy under Format > Sheet;
-- Mac Excel keeps it under Edit > Sheet. The trigger letters are
-- whatever the user wants — they don't have to mirror the Mac menu
-- path. Verified the label is the literal three-dot "Move or Copy
-- Sheet..." (not the Unicode ellipsis) on this Excel build via a
-- probe dump.
--
-- No post-action Escape: this opens a modal dialog that already
-- holds focus, and an Escape would just close it.
function M.move_sheet_dialog()
  M.applescript([[
    tell application "System Events"
      tell process "Microsoft Excel"
        click menu item "Move or Copy Sheet..." of menu 1 of menu item "Sheet" of menu 1 of menu bar item "Edit" of menu bar 1
      end tell
    end tell
  ]])
end

-- Insert a new worksheet via Excel's AppleScript dictionary.
--
-- On Mac Excel, `make new worksheet at active workbook` inserts
-- immediately *before* the active sheet tab and activates the new
-- sheet — that is native behaviour; repositioning it after the
-- previous sheet is not reliable via AppleScript on the builds
-- we've tested (`move … to after` and `at after active sheet`
-- both fail with parameter errors).
--
-- Post-action Escape clears the same arrow-key-eating cell selector
-- freeze we saw with paste actions; same 50 ms timing.
function M.insert_sheet()
  local ok, result = hs.osascript.applescript([[
    tell application "Microsoft Excel"
      try
        make new worksheet at active workbook
        return "ok"
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]])

  local failed = (not ok)
    or (type(result) == "string" and result:sub(1, 6) == "ERROR ")
  if failed then
    local detail = (type(result) == "string") and result or "engine error"
    if _G.__mme_log then
      _G.__mme_log("insert_sheet: %s", detail)
    end
    hs.alert.show("Insert sheet failed (see log)", 1.5)
    return
  end

  hs.timer.doAfter(0.05, function()
    hs.eventtap.keyStroke({}, "escape", 0)
  end)
end

-- Delete the active worksheet. Routes through Edit > Sheet >
-- Delete Sheet in the menu bar so Excel surfaces its own "Are you
-- sure?" confirmation dialog; the user confirms with Return or
-- cancels with Escape.
--
-- Why this shape (System Events menu click + dismiss_menu_focus)
-- rather than `delete active sheet` via AppleScript directly:
--   The direct AppleScript call did pop the confirmation dialog,
--   but the dialog opened without keyboard focus — the user's
--   first Return got eaten until they clicked the dialog or hit
--   Escape a couple of times. Same ribbon-retained-focus quirk as
--   rename_sheet, so the same trick applies: drive the action
--   through the menu bar and trail with a single Escape to clear
--   the ribbon. By the time the alert dialog finishes rendering,
--   the responder chain is clean and the dialog inherits first
--   responder on its own. Enter then activates the default Delete
--   button.
--
-- (An earlier attempt synthesised an AX-located mouse click on the
-- dialog's top heading band; that was sufficient for the
-- text-field dialogs but not for button-only alerts. The Escape
-- approach worked in manual testing and replaces it cleanly.)
function M.delete_sheet()
  M.applescript([[
    tell application "System Events"
      tell process "Microsoft Excel"
        click menu item "Delete Sheet" of menu 1 of menu item "Sheet" of menu 1 of menu bar item "Edit" of menu bar 1
      end tell
    end tell
  ]])
  M.dismiss_menu_focus()
end

----------------------------------------------------------------------
-- View actions
----------------------------------------------------------------------

-- Change the active window's zoom in the direction of `delta`,
-- snapping to a grid of |delta| percentage points and clamped to
-- [config.zoom_min, config.zoom_max]. Used by the M.zoom_in /
-- M.zoom_out wrappers below.
--
-- Snap behaviour:
--   - delta > 0 (zoom in): target is the next grid line strictly
--     ABOVE current. From 117% with step 10 you land on 120%; from
--     120% you land on 130%.
--   - delta < 0 (zoom out): target is the next grid line strictly
--     BELOW current. From 117% with step 10 you land on 110%; from
--     120% you land on 110%.
-- The grid is defined by |delta| itself, so changing config.zoom_step
-- changes the grid live — no separate config knob.
--
-- The read is wrapped in try/on error inside AppleScript so a
-- missing-active-window state (e.g. no workbook open) becomes a
-- silent no-op rather than a flash alert. The write goes through
-- M.applescript, so any unexpected write failure does surface.
function M.zoom_by(delta)
  local min = config.zoom_min or 50
  local max = config.zoom_max or 200

  if delta == 0 then return end

  local ok, current = hs.osascript.applescript([[
    tell application "Microsoft Excel"
      try
        return zoom of active window
      on error errMsg
        return "ERROR: " & errMsg
      end try
    end tell
  ]])

  local cur = ok and tonumber(current) or nil
  if not cur then
    if _G.__mme_log then
      _G.__mme_log("zoom_by: could not read current zoom (%s)", tostring(current))
    end
    return
  end

  local step = math.abs(delta)
  local target
  if delta > 0 then
    target = math.floor(cur / step) * step + step
  else
    target = math.ceil(cur / step) * step - step
  end

  if target < min then target = min end
  if target > max then target = max end
  if target == cur then return end  -- already at the bound in this direction

  M.applescript(string.format([[
    tell application "Microsoft Excel"
      set zoom of active window to %d
    end tell
  ]], target))
end

function M.zoom_in()  M.zoom_by( config.zoom_step or 10) end
function M.zoom_out() M.zoom_by(-(config.zoom_step or 10)) end

----------------------------------------------------------------------
-- Font actions
----------------------------------------------------------------------

-- Cycle the font colour of the current selection through the list
-- defined in config.font_color_cycle.
--
-- Per-cell semantics: read the font colour of the first cell in the
-- selection, find it in the cycle, apply the next colour (wrapping
-- at the end) to the whole selection. If the current colour isn't
-- in the cycle (the common case being Excel's default black on a
-- cell the user has never coloured), apply the first cycle entry.
--
-- Implementation notes:
--   - Excel for Mac is asymmetric about RGB units: it ACCEPTS 16-bit
--     values (0..65535) on write but RETURNS 8-bit values (0..255)
--     on read. Found this empirically: setting {1028, 13107, 65535}
--     for blue paints the cell correctly, but reading the same cell
--     back gives {4, 51, 255}. So writing scales hex 0..255 up by
--     257 to 0..65535; reading auto-detects 8-bit vs 16-bit based
--     on whether any component exceeds 255, and only rescales when
--     needed (future-proof against an Excel version that ever
--     returns 16-bit again).
--   - The read is wrapped in try/on error so a non-range selection
--     (chart, image, etc.) returns "ERROR: ..." instead of throwing.
--     In that case we fall through to applying cycle[1].
--   - The set goes through M.applescript, so any failure (e.g.
--     selection is read-only) surfaces as an alert and a log line.
function M.cycle_font_color()
  local cycle = config.font_color_cycle or {}
  if #cycle == 0 then
    if _G.__mme_log then
      _G.__mme_log("cycle_font_color: config.font_color_cycle is empty")
    end
    return
  end

  -- Conversion helpers between 8-bit "RRGGBB" hex (what the user
  -- edits in config) and Excel's AppleScript RGB tuples. See the
  -- docstring for why the read side has to auto-detect units.

  -- "RRGGBB" hex -> 16-bit RGB (0..65535) for the SET AppleScript.
  local function hex_to_rgb16(hex)
    local r = tonumber(hex:sub(1, 2), 16)
    local g = tonumber(hex:sub(3, 4), 16)
    local b = tonumber(hex:sub(5, 6), 16)
    if not (r and g and b) then return nil end
    return r * 257, g * 257, b * 257
  end

  -- Whatever Excel returned (8-bit or 16-bit per channel) -> "RRGGBB"
  -- hex. If any component exceeds 255 we treat the input as 16-bit
  -- and rescale; otherwise we trust the input is already 8-bit and
  -- format directly.
  local function read_rgb_to_hex(r, g, b)
    if math.max(r, g, b) > 255 then
      return string.format("%02X%02X%02X",
        math.floor(r / 257 + 0.5),
        math.floor(g / 257 + 0.5),
        math.floor(b / 257 + 0.5))
    end
    return string.format("%02X%02X%02X", r, g, b)
  end

  -- Normalise cycle entries: strip optional leading '#', uppercase,
  -- discard anything that isn't a valid 6-hex string.
  local cycle_norm = {}
  for _, c in ipairs(cycle) do
    local h = tostring(c):gsub("^#", ""):upper()
    if #h == 6 and hex_to_rgb16(h) then
      cycle_norm[#cycle_norm + 1] = h
    elseif _G.__mme_log then
      _G.__mme_log("cycle_font_color: ignoring invalid hex %q", tostring(c))
    end
  end
  if #cycle_norm == 0 then
    if _G.__mme_log then
      _G.__mme_log("cycle_font_color: no valid hex entries in config.font_color_cycle")
    end
    return
  end

  -- Read the current font colour of the first selected cell.
  local ok_read, result = hs.osascript.applescript([[
    tell application "Microsoft Excel"
      try
        set fc to color of font object of (cell 1 of selection)
        return ((item 1 of fc) as text) & "," & ((item 2 of fc) as text) & "," & ((item 3 of fc) as text)
      on error errMsg
        return "ERROR: " & errMsg
      end try
    end tell
  ]])

  local current_hex
  if ok_read and type(result) == "string" and not result:find("^ERROR") then
    local r, g, b = result:match("^(%-?%d+),(%-?%d+),(%-?%d+)$")
    if r and g and b then
      current_hex = read_rgb_to_hex(tonumber(r), tonumber(g), tonumber(b))
    end
  end

  local next_hex = cycle_norm[1]
  if current_hex then
    for i, c in ipairs(cycle_norm) do
      if c == current_hex then
        next_hex = cycle_norm[(i % #cycle_norm) + 1]
        break
      end
    end
  end

  local nr, ng, nb = hex_to_rgb16(next_hex)
  M.applescript(string.format([[
    tell application "Microsoft Excel"
      set color of font object of selection to {%d, %d, %d}
    end tell
  ]], nr, ng, nb))
end

-- Step the font size of the current selection through the discrete
-- ladder defined in config.font_size_cycle.
--
-- Semantics mirror zoom_in / zoom_out but with a list of explicit sizes
-- instead of a step grid. `direction` is +1 (step up) or -1 (step down).
--
-- Behaviour:
--   - If the current size is in the ladder, move to the next/previous
--     entry; clamp at the top/bottom (no wrap).
--   - If the current size is between ladder entries, snap to the next
--     entry strictly above (up) or strictly below (down). E.g. with
--     ladder {9,12,18,24} and current 14: up → 18, down → 12.
--   - If the current size is below the smallest entry, up → smallest;
--     down → smallest (clamped).
--   - If the current size is above the largest entry, down → largest;
--     up → largest (clamped).
--
-- The AppleScript form for read/write deliberately uses a `tell font
-- object of <range>` block. Setting `font size` directly on the
-- selection is rejected on some Excel builds with -10006; the block
-- form is accepted on every build we've tested.
function M.step_font_size(direction)
  if direction ~= 1 and direction ~= -1 then return end

  local cycle = config.font_size_cycle or {}
  if #cycle == 0 then
    if _G.__mme_log then
      _G.__mme_log("step_font_size: config.font_size_cycle is empty")
    end
    return
  end

  -- Normalise to a sorted ascending ladder of unique positive numbers.
  local seen, ladder = {}, {}
  for _, s in ipairs(cycle) do
    local n = tonumber(s)
    if n and n > 0 and not seen[n] then
      seen[n] = true
      ladder[#ladder + 1] = n
    elseif (not n or n <= 0) and _G.__mme_log then
      _G.__mme_log("step_font_size: ignoring invalid entry %q", tostring(s))
    end
  end
  table.sort(ladder)
  if #ladder == 0 then
    if _G.__mme_log then
      _G.__mme_log("step_font_size: no valid entries in config.font_size_cycle")
    end
    return
  end

  local ok_read, result = hs.osascript.applescript([[
    tell application "Microsoft Excel"
      try
        tell font object of (cell 1 of selection)
          return font size as text
        end tell
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]])

  local current
  if ok_read and type(result) == "number" then
    current = result
  elseif ok_read and type(result) == "string" and not result:find("^ERROR ") then
    current = tonumber(result)
  end

  local min, max = ladder[1], ladder[#ladder]
  local target

  if not current then
    -- Couldn't read; behave like a first press at the bound for the
    -- chosen direction so the user still gets a defined outcome.
    target = (direction > 0) and min or max
  elseif direction > 0 then
    for _, n in ipairs(ladder) do
      if n > current + 0.01 then target = n; break end
    end
    target = target or max  -- already at or above the top
  else
    for i = #ladder, 1, -1 do
      if ladder[i] < current - 0.01 then target = ladder[i]; break end
    end
    target = target or min  -- already at or below the bottom
  end

  if current and math.abs(target - current) < 0.01 then return end

  local ok_set, set_result = hs.osascript.applescript(string.format([[
    tell application "Microsoft Excel"
      try
        tell font object of selection
          set font size to %s
        end tell
        return "ok"
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]], tostring(target)))

  local failed = (not ok_set)
    or (type(set_result) == "string" and set_result:sub(1, 6) == "ERROR ")
  if failed then
    local detail = (type(set_result) == "string") and set_result or "engine error"
    if _G.__mme_log then
      _G.__mme_log("step_font_size: %s", detail)
    end
    hs.alert.show("Font size change failed (see log)", 1.5)
  end
end

function M.font_size_up()   M.step_font_size( 1) end
function M.font_size_down() M.step_font_size(-1) end

----------------------------------------------------------------------
-- Border actions
----------------------------------------------------------------------

-- Cycle outer borders on the selection through config.border_placement_cycle.
-- `weight` is an Excel AppleScript border-weight name (e.g. "thin",
-- "medium") taken from config.border_weight_normal / border_weight_thick.
--
-- Each press clears all four outer edges, then draws only the next
-- placement in the cycle (top only, left only, right only, or full
-- outline). Placement detection reads the first selected cell; if the
-- pattern is not recognised, the first cycle entry is applied.
function M.cycle_border(weight, weight_label)
  local placements = config.border_placement_cycle or {}
  local valid = { top = true, left = true, right = true, outline = true }
  local cycle_norm = {}
  for _, p in ipairs(placements) do
    local s = tostring(p):lower()
    if valid[s] then
      cycle_norm[#cycle_norm + 1] = s
    elseif _G.__mme_log then
      _G.__mme_log("cycle_border: ignoring invalid placement %q", tostring(p))
    end
  end
  if #cycle_norm == 0 then
    if _G.__mme_log then
      _G.__mme_log("cycle_border: config.border_placement_cycle is empty")
    end
    return nil
  end

  local weight_as = tostring(weight or config.border_weight_normal or "thin"):lower()
  local label = tostring(weight_label or weight_as):lower()
  if label ~= "thin" and label ~= "thick" then
    label = (weight_as == "medium" or weight_as == "thick") and "thick" or "thin"
  end
  local allowed_weight = { hairline = true, thin = true, medium = true, thick = true }
  if not allowed_weight[weight_as] then
    if _G.__mme_log then
      _G.__mme_log("cycle_border: invalid weight %q, using thin", tostring(weight))
    end
    weight_as = "thin"
  end

  local function classify(top, left, right, bottom)
    if top and not left and not right and not bottom then return "top" end
    if left and not top and not right and not bottom then return "left" end
    if right and not top and not left and not bottom then return "right" end
    if top and left and right and bottom then return "outline" end
    return nil
  end

  local ok_read, result = hs.osascript.applescript([[
    tell application "Microsoft Excel"
      try
        tell cell 1 of selection
          set t to false
          set l to false
          set r to false
          set b to false
          set bTop to get border which border edge top
          if line style of bTop is not line style none then set t to true
          set bLeft to get border which border edge left
          if line style of bLeft is not line style none then set l to true
          set bRight to get border which border edge right
          if line style of bRight is not line style none then set r to true
          set bBot to get border which border edge bottom
          if line style of bBot is not line style none then set b to true
          return (t as text) & "," & (l as text) & "," & (r as text) & "," & (b as text)
        end tell
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]])

  local current_placement
  if ok_read and type(result) == "string" and not result:find("^ERROR ") then
    local ts, ls, rs, bs = result:match("^(%w+),(%w+),(%w+),(%w+)$")
    if ts then
      local function as_bool(s) return s == "true" end
      current_placement = classify(as_bool(ts), as_bool(ls), as_bool(rs), as_bool(bs))
    end
  end

  local next_placement = cycle_norm[1]
  if current_placement then
    for i, p in ipairs(cycle_norm) do
      if p == current_placement then
        next_placement = cycle_norm[(i % #cycle_norm) + 1]
        break
      end
    end
  end

  -- Excel Mac: borders are accessed with `get border which border edge top`
  -- inside a `tell <range>` block — not `border index edge top of range`.
  local apply_block
  if next_placement == "outline" then
    apply_block = string.format(
      "border around it line style continuous weight border weight %s",
      weight_as)
  else
    local edge_sym = ({
      top = "edge top",
      left = "edge left",
      right = "edge right",
    })[next_placement]
    if not edge_sym then
      if _G.__mme_log then
        _G.__mme_log("cycle_border: unknown placement %q", tostring(next_placement))
      end
      return nil
    end
    apply_block = string.format([[
          set b to get border which border %s
          set line style of b to continuous
          set weight of b to border weight %s]], edge_sym, weight_as)
  end

  local ok_set, set_result, descriptor = hs.osascript.applescript(string.format([[
    tell application "Microsoft Excel"
      try
        tell selection
          repeat with wb in {edge top, edge left, edge right, edge bottom}
            set b to get border which border wb
            set line style of b to line style none
          end repeat
%s
        end tell
        return "ok"
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]], apply_block))

  local failed = (not ok_set)
    or (type(set_result) == "string" and set_result:sub(1, 6) == "ERROR ")
  if failed then
    local detail
    if type(set_result) == "string" and set_result:sub(1, 6) == "ERROR " then
      detail = set_result
    elseif type(descriptor) == "table" and hs.inspect then
      detail = hs.inspect(descriptor)
    else
      detail = string.format("engine error ok=%s result=%s",
        tostring(ok_set), tostring(set_result))
    end
    if _G.__mme_log then
      _G.__mme_log("cycle_border: %s", detail)
    end
    hs.alert.show("Border change failed (see log)", 1.5)
    return nil
  end

  return label .. " " .. next_placement
end

function M.cycle_border_thin()
  return M.cycle_border(config.border_weight_normal or "thin", "thin")
end

function M.cycle_border_thick()
  return M.cycle_border(config.border_weight_thick or "medium", "thick")
end

-- Reapply whatever outer borders the selection already has as dashed
-- lines. Reads which edges are active and each edge's weight from the
-- first selected cell, then applies the same edges on the whole
-- selection: lighter weights (hairline / thin) use
-- config.border_dotted_weight_light; medium / thick use
-- config.border_dotted_weight_heavy. Both share
-- config.border_dotted_line_style (dash by default).
function M.border_dotted()
  local line_style = tostring(config.border_dotted_line_style or "dash"):lower()
  local weight_light = tostring(config.border_dotted_weight_light or "thin"):lower()
  local weight_heavy = tostring(config.border_dotted_weight_heavy or "medium"):lower()
  local allowed_style = {
    continuous = true, dash = true, ["dash dot"] = true,
    ["dash dot dot"] = true, dot = true, double = true,
    ["slant dash dot"] = true,
  }
  local allowed_weight = { hairline = true, thin = true, medium = true, thick = true }
  if not allowed_style[line_style] then
    if _G.__mme_log then
      _G.__mme_log("border_dotted: invalid line style %q, using dash", line_style)
    end
    line_style = "dash"
  end
  if not allowed_weight[weight_light] then
    if _G.__mme_log then
      _G.__mme_log("border_dotted: invalid light weight %q, using thin", weight_light)
    end
    weight_light = "thin"
  end
  if not allowed_weight[weight_heavy] then
    if _G.__mme_log then
      _G.__mme_log("border_dotted: invalid heavy weight %q, using medium", weight_heavy)
    end
    weight_heavy = "medium"
  end

  local ok, result, descriptor = hs.osascript.applescript(string.format([[
    tell application "Microsoft Excel"
      try
        set specs to {}
        tell cell 1 of selection
          repeat with wb in {edge top, edge left, edge right, edge bottom}
            set b to get border which border wb
            if line style of b is not line style none then
              set end of specs to {wb, weight of b as text}
            end if
          end repeat
        end tell
        if (count of specs) is 0 then return "none"
        tell selection
          repeat with spec in specs
            set wb to item 1 of spec
            set wTxt to item 2 of spec
            set b to get border which border wb
            set line style of b to %s
            if wTxt contains "thick" or wTxt contains "medium" then
              set weight of b to border weight %s
            else
              set weight of b to border weight %s
            end if
          end repeat
        end tell
        return "ok"
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]], line_style, weight_heavy, weight_light))

  if ok and result == "none" then return end

  local failed = (not ok)
    or (type(result) == "string" and result:sub(1, 6) == "ERROR ")
  if failed then
    local detail
    if type(result) == "string" and result:sub(1, 6) == "ERROR " then
      detail = result
    elseif type(descriptor) == "table" and hs.inspect then
      detail = hs.inspect(descriptor)
    else
      detail = string.format("engine error ok=%s result=%s",
        tostring(ok), tostring(result))
    end
    if _G.__mme_log then
      _G.__mme_log("border_dotted: %s", detail)
    end
    hs.alert.show("Border dotted failed (see log)", 1.5)
    return
  end

  -- Same cell-selector freeze as paste / insert sheet: Excel parks
  -- focus on an overlay after AppleScript border changes until Escape.
  M.dismiss_menu_focus()
end

----------------------------------------------------------------------
-- Fill actions
----------------------------------------------------------------------

-- Cycle the fill (cell background) colour of the current selection
-- through the list defined in config.fill_color_cycle.
--
-- Semantics mirror cycle_font_color: read the fill of the first cell in
-- the selection, advance to the next entry (wrapping), and apply to the
-- whole selection. The cycle may include the literal string "none" to
-- clear fill.
function M.cycle_fill_color()
  local cycle = config.fill_color_cycle or {}
  if #cycle == 0 then
    if _G.__mme_log then
      _G.__mme_log("cycle_fill_color: config.fill_color_cycle is empty")
    end
    return
  end

  local function hex_to_rgb16(hex)
    local r = tonumber(hex:sub(1, 2), 16)
    local g = tonumber(hex:sub(3, 4), 16)
    local b = tonumber(hex:sub(5, 6), 16)
    if not (r and g and b) then return nil end
    return r * 257, g * 257, b * 257
  end

  local function read_rgb_to_hex(r, g, b)
    if math.max(r, g, b) > 255 then
      return string.format("%02X%02X%02X",
        math.floor(r / 257 + 0.5),
        math.floor(g / 257 + 0.5),
        math.floor(b / 257 + 0.5))
    end
    return string.format("%02X%02X%02X", r, g, b)
  end

  -- Normalise cycle entries:
  --   - "none" (any case) is preserved as the sentinel "NONE"
  --   - valid hex strings are uppercased without leading '#'
  local cycle_norm = {}
  for _, c in ipairs(cycle) do
    local s = tostring(c)
    if s:lower() == "none" then
      cycle_norm[#cycle_norm + 1] = "NONE"
    else
      local h = s:gsub("^#", ""):upper()
      if #h == 6 and hex_to_rgb16(h) then
        cycle_norm[#cycle_norm + 1] = h
      elseif _G.__mme_log then
        _G.__mme_log("cycle_fill_color: ignoring invalid entry %q", tostring(c))
      end
    end
  end
  if #cycle_norm == 0 then
    if _G.__mme_log then
      _G.__mme_log("cycle_fill_color: no valid entries in config.fill_color_cycle")
    end
    return
  end

  -- Read the current fill of the first selected cell.
  -- We try to distinguish "no fill" from a real colour by checking the
  -- interior pattern first; if that fails, we fall back to reading the
  -- colour tuple and treating "missing value" as NONE.
  local ok_read, result = hs.osascript.applescript([[
    tell application "Microsoft Excel"
      try
        set c to cell 1 of selection
        set i to interior object of c
        try
          set p to pattern of i
          if p is pattern none then return "NONE"
        end try

        set fc to color of i
        if fc is missing value then return "NONE"
        return ((item 1 of fc) as text) & "," & ((item 2 of fc) as text) & "," & ((item 3 of fc) as text)
      on error errMsg
        return "ERROR: " & errMsg
      end try
    end tell
  ]])

  local current_key
  if ok_read and type(result) == "string" and not result:find("^ERROR") then
    if result == "NONE" then
      current_key = "NONE"
    else
      local r, g, b = result:match("^(%-?%d+),(%-?%d+),(%-?%d+)$")
      if r and g and b then
        current_key = read_rgb_to_hex(tonumber(r), tonumber(g), tonumber(b))
      end
    end
  end

  local next_key = cycle_norm[1]
  if current_key then
    for i, c in ipairs(cycle_norm) do
      if c == current_key then
        next_key = cycle_norm[(i % #cycle_norm) + 1]
        break
      end
    end
  end

  if next_key == "NONE" then
    M.applescript([[
      tell application "Microsoft Excel"
        try
          set i to interior object of selection
          set pattern of i to pattern none
        on error errMsg
          return "ERROR: " & errMsg
        end try
      end tell
    ]])
    return
  end

  local nr, ng, nb = hex_to_rgb16(next_key)
  M.applescript(string.format([[
    tell application "Microsoft Excel"
      try
        set i to interior object of selection
        set pattern of i to pattern solid
        set color of i to {%d, %d, %d}
      on error errMsg
        return "ERROR: " & errMsg
      end try
    end tell
  ]], nr, ng, nb))
end

----------------------------------------------------------------------
-- Selection actions
----------------------------------------------------------------------

-- Expand the current selection to cover the full column(s) it spans.
-- Mirrors Excel's Windows-style Ctrl+Space, but bound to
-- Ctrl+Shift+Space here because plain Ctrl+Space is widely
-- intercepted by macOS for input-source switching before Excel ever
-- sees the keystroke.
--
-- `entire column of selection` returns a Range covering every
-- column that any cell of the current selection touches, so a
-- multi-column selection grows to all those columns at once —
-- matching native Excel behaviour. No post-action Escape needed:
-- this is a pure selection change, no overlay or modal lands.
function M.select_column()
  M.applescript([[
    tell application "Microsoft Excel"
      select (entire column of selection)
    end tell
  ]])
end

----------------------------------------------------------------------
-- Sheet navigation
----------------------------------------------------------------------

-- Activate the next / previous sheet in the workbook by piggy-
-- backing on Mac Excel's own native shortcut: Opt+Right (next) and
-- Opt+Left (prev). We just synthesise those keystrokes; Excel
-- handles all the workbook lookup, sheet ordering, and clamping at
-- the workbook ends.
--
-- Why we gave up on AppleScript here:
--   We tried three increasingly-defensive AppleScript variants —
--   `index of active sheet`, then a stored `active workbook`, then
--   `name of every worksheet of active workbook` — and each one
--   ended up coercing a `missing value` into an integer somewhere
--   in the chain. With per-step labelled `try/on error` we
--   narrowed it to the worksheet-name lookup (`name of worksheet i
--   of active workbook` is missing value on at least one i, on
--   this Excel build). Rather than keep guessing which AppleScript
--   subexpression will misfire next, we let Excel's existing
--   keyboard handler do the lookup for us.
--
-- Trigger is whatever the user wires up in shortcuts.lua. They've
-- chosen Shift+Opt+Down=next and Shift+Opt+Up=prev. Note the user
-- will be physically holding Shift+Opt when these fire; the
-- synthesised events carry only the Opt flag, so Excel sees clean
-- Opt+Right / Opt+Left. If you ever see Excel responding as if
-- Shift were also down, that's the place to look.
function M.next_sheet() M.send({ "alt" }, "right") end
function M.prev_sheet() M.send({ "alt" }, "left")  end

return M
