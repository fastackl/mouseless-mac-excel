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

-- Bring keyboard focus to a list/outline/table inside a dialog (e.g.
-- Move or Copy's "Before sheet" pane) via a synthesised click, then
-- optionally jump to the first row with Home. Same AppKit quirk as
-- text fields: AXFocused on the list is not enough for arrow keys.
--
-- opts.select_first (default true): send Home after the click so the
-- highlight starts on the first sheet name, not wherever Excel left it.
--
-- Opt-out via config.dialog_focus_click (shared with text-field dialogs).
function M.focus_dialog_list(window_title, opts)
  if not config.dialog_focus_click then return end
  opts = opts or {}
  local select_first = opts.select_first ~= false

  hs.timer.doAfter(config.dialog_focus_click_delay_seconds, function()
    local ok, err = pcall(function()
      local ax = hs.axuielement
      if not ax then
        local ok_req, mod = pcall(require, "hs.axuielement")
        if ok_req then ax = mod end
      end
      if not ax then
        if _G.__mme_log then
          _G.__mme_log("focus_dialog_list: hs.axuielement unavailable")
        end
        return
      end

      local app = hs.application.find(config.excel_bundle_id)
      if not app then return end

      local app_ax = ax.applicationElement(app)
      if not app_ax then return end

      local function attr(elem, name)
        local got, val = pcall(function() return elem[name] end)
        return got and val or nil
      end

      local dialog
      for _, w in ipairs(attr(app_ax, "AXWindows") or {}) do
        if attr(w, "AXTitle") == window_title then
          dialog = w
          break
        end
      end
      if not dialog then
        if _G.__mme_log then
          _G.__mme_log(
            "focus_dialog_list: dialog %q not found (try bumping config.dialog_focus_click_delay_seconds)",
            tostring(window_title))
        end
        return
      end

      local LIST_ROLES = {
        ["AXList"] = true,
        ["AXOutline"] = true,
        ["AXTable"] = true,
      }

      local function element_area(elem)
        local frame = attr(elem, "AXFrame")
        if not frame then
          local pos  = attr(elem, "AXPosition")
          local size = attr(elem, "AXSize")
          if pos and size then
            frame = { x = pos.x, y = pos.y, w = size.w, h = size.h }
          end
        end
        if not frame then return 0 end
        local fw = frame.w or frame.width or 0
        local fh = frame.h or frame.height or 0
        return fw * fh
      end

      -- Move or Copy has a pop-up for "To book" and a tall list for
      -- "Before sheet"; pick the largest list-like control in the tree.
      local list, best_area = nil, 0
      local function walk(elem)
        local role = attr(elem, "AXRole")
        if role and LIST_ROLES[role] then
          local area = element_area(elem)
          if area > best_area then
            list, best_area = elem, area
          end
        end
        for _, child in ipairs(attr(elem, "AXChildren") or {}) do
          walk(child)
        end
      end
      walk(dialog)

      if not list then
        if _G.__mme_log then
          _G.__mme_log("focus_dialog_list: list not found in %q",
            tostring(window_title))
        end
        return
      end

      local function click_center(elem)
        local frame = attr(elem, "AXFrame")
        if not frame then
          local pos  = attr(elem, "AXPosition")
          local size = attr(elem, "AXSize")
          if pos and size then
            frame = { x = pos.x, y = pos.y, w = size.w, h = size.h }
          end
        end
        if not frame then return false end
        local fw = frame.w or frame.width or 0
        local fh = frame.h or frame.height or 0
        local click_point = { x = frame.x + fw / 2, y = frame.y + fh / 2 }
        local original_mouse = hs.mouse.absolutePosition()
        hs.eventtap.leftClick(click_point)
        hs.mouse.absolutePosition(original_mouse)
        return true
      end

      -- Prefer clicking the first row so selection and focus align.
      local target = list
      local rows = attr(list, "AXRows") or attr(list, "AXVisibleRows")
      if type(rows) == "table" and #rows > 0 then
        target = rows[1]
      end

      if not click_center(target) then
        if _G.__mme_log then
          _G.__mme_log("focus_dialog_list: could not resolve click target in %q",
            tostring(window_title))
        end
        return
      end

      if select_first then
        hs.timer.doAfter(0.03, function()
          hs.eventtap.keyStroke({}, "home", 0)
        end)
      end
    end)

    if not ok and _G.__mme_log then
      _G.__mme_log("focus_dialog_list error: %s", tostring(err))
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

----------------------------------------------------------------------
-- Move or Copy dialog: ephemeral Alt+C toggles "Create a copy"
----------------------------------------------------------------------

local MOVE_OR_COPY_DIALOG_TITLE = "Move or Copy"
local move_sheet_copy_hotkey = nil
local move_sheet_copy_watch_timer = nil

local function disarm_move_sheet_copy_hotkey()
  if move_sheet_copy_hotkey then
    move_sheet_copy_hotkey:delete()
    move_sheet_copy_hotkey = nil
  end
  if move_sheet_copy_watch_timer then
    move_sheet_copy_watch_timer:stop()
    move_sheet_copy_watch_timer = nil
  end
end

local function move_sheet_ax_attr(elem, name)
  local got, val = pcall(function() return elem[name] end)
  return got and val or nil
end

local function move_sheet_excel_app_ax()
  local ax = hs.axuielement
  if not ax then
    local ok_req, mod = pcall(require, "hs.axuielement")
    if ok_req then ax = mod end
  end
  if not ax then return nil end
  local app = hs.application.find(config.excel_bundle_id)
  if not app then return nil end
  return ax.applicationElement(app)
end

local function find_move_or_copy_dialog()
  local app_ax = move_sheet_excel_app_ax()
  if not app_ax then return nil end
  for _, w in ipairs(move_sheet_ax_attr(app_ax, "AXWindows") or {}) do
    if move_sheet_ax_attr(w, "AXTitle") == MOVE_OR_COPY_DIALOG_TITLE then
      return w
    end
  end
  return nil
end

local function find_create_copy_checkbox(dialog)
  local function matches_checkbox_label(elem)
    local role = move_sheet_ax_attr(elem, "AXRole")
    if role ~= "AXCheckBox" then return false end
    local title = tostring(
      move_sheet_ax_attr(elem, "AXTitle")
        or move_sheet_ax_attr(elem, "AXDescription")
        or "")
    return title:lower():find("create a copy", 1, true) ~= nil
  end

  local function walk(elem)
    if matches_checkbox_label(elem) then return elem end
    for _, child in ipairs(move_sheet_ax_attr(elem, "AXChildren") or {}) do
      local found = walk(child)
      if found then return found end
    end
    return nil
  end
  return walk(dialog)
end

local function checkbox_is_on(cb)
  local val = move_sheet_ax_attr(cb, "AXValue")
  return val == 1 or val == true or val == "1"
end

local function click_ax_center(elem)
  local frame = move_sheet_ax_attr(elem, "AXFrame")
  if not frame then
    local pos  = move_sheet_ax_attr(elem, "AXPosition")
    local size = move_sheet_ax_attr(elem, "AXSize")
    if pos and size then
      frame = { x = pos.x, y = pos.y, w = size.w, h = size.h }
    end
  end
  if not frame then return false end
  local fw = frame.w or frame.width or 0
  local fh = frame.h or frame.height or 0
  local click_point = { x = frame.x + fw / 2, y = frame.y + fh / 2 }
  local original_mouse = hs.mouse.absolutePosition()
  hs.eventtap.leftClick(click_point)
  hs.mouse.absolutePosition(original_mouse)
  return true
end

local function toggle_move_sheet_create_copy()
  local dialog = find_move_or_copy_dialog()
  if not dialog then return nil end

  local checkbox = find_create_copy_checkbox(dialog)
  if not checkbox then
    if _G.__mme_log then
      _G.__mme_log("toggle_create_copy: checkbox not found")
    end
    hs.alert.show("Create a copy checkbox not found", 1.2)
    return nil
  end

  local was_on = checkbox_is_on(checkbox)
  local set_ok = pcall(function()
    checkbox:setAttributeValue("AXValue", was_on and 0 or 1)
  end)

  local now_on = checkbox_is_on(checkbox)
  if set_ok and now_on ~= was_on then
    return now_on and "Create a copy: on" or "Create a copy: off"
  end

  if not click_ax_center(checkbox) then
    if _G.__mme_log then
      _G.__mme_log("toggle_create_copy: could not click checkbox")
    end
    hs.alert.show("Could not toggle Create a copy", 1.2)
    return nil
  end

  now_on = checkbox_is_on(checkbox)
  return now_on and "Create a copy: on" or "Create a copy: off"
end

local function arm_move_sheet_copy_hotkey()
  if not config.move_sheet_copy_toggle_enabled then return end

  disarm_move_sheet_copy_hotkey()

  local mods = config.move_sheet_copy_toggle_mods or { "alt" }
  local key  = config.move_sheet_copy_toggle_key or "c"

  move_sheet_copy_hotkey = hs.hotkey.new(mods, key, function()
    if not find_move_or_copy_dialog() then
      disarm_move_sheet_copy_hotkey()
      return
    end
    local label = toggle_move_sheet_create_copy()
    if label then
      hs.alert.show(label, 0.4)
    end
  end)
  move_sheet_copy_hotkey:enable()

  local interval = config.move_sheet_copy_watch_interval_seconds or 0.25
  move_sheet_copy_watch_timer = hs.timer.doEvery(interval, function()
    local front = hs.application.frontmostApplication()
    if not front or front:bundleID() ~= config.excel_bundle_id then
      disarm_move_sheet_copy_hotkey()
      return
    end
    if not find_move_or_copy_dialog() then
      disarm_move_sheet_copy_hotkey()
    end
  end)
end

-- Open Edit > Sheet > Move or Copy Sheet... — Excel's native
-- "Move or Copy" dialog where the user picks a destination workbook
-- and target position. No text field to overtype; we nudge focus into
-- the "Before sheet" list and select the first row so arrow keys work
-- without reaching for the mouse (same click trick as Column Width).
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
--
-- While the dialog stays open, config.move_sheet_copy_toggle_mods +
-- key (default Alt+C) toggles "Create a copy"; the hotkey disarms when
-- the dialog closes or Excel loses focus.
function M.move_sheet_dialog()
  disarm_move_sheet_copy_hotkey()
  M.applescript([[
    tell application "System Events"
      tell process "Microsoft Excel"
        click menu item "Move or Copy Sheet..." of menu 1 of menu item "Sheet" of menu 1 of menu bar item "Edit" of menu bar 1
      end tell
    end tell
  ]])
  M.focus_dialog_list("Move or Copy")
  arm_move_sheet_copy_hotkey()
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

-- Cycle the number format of the current selection through
-- config.number_format_cycle.
--
-- Semantics mirror cycle_fill_color: read the format of the first cell,
-- advance to the next entry (wrapping), apply to the whole selection.
-- Entries "none" or "general" (any case) apply Excel's General format.
function M.cycle_number_format()
  local cycle = config.number_format_cycle or {}
  if #cycle == 0 then
    if _G.__mme_log then
      _G.__mme_log("cycle_number_format: config.number_format_cycle is empty")
    end
    return
  end

  local cycle_norm = {}
  for _, entry in ipairs(cycle) do
    local s = tostring(entry)
    if s:lower() == "none" or s:lower() == "general" then
      cycle_norm[#cycle_norm + 1] = "GENERAL"
    else
      cycle_norm[#cycle_norm + 1] = s
    end
  end

  local ok_read, result = hs.osascript.applescript([[
    tell application "Microsoft Excel"
      try
        return number format of (cell 1 of selection)
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]])

  local current_key
  if ok_read and type(result) == "string" and not result:find("^ERROR ") then
    if result:lower() == "general" then
      current_key = "GENERAL"
    else
      current_key = result
    end
  end

  local next_key = cycle_norm[1]
  if current_key then
    for i, fmt in ipairs(cycle_norm) do
      if fmt == current_key then
        next_key = cycle_norm[(i % #cycle_norm) + 1]
        break
      end
    end
  end

  -- Bake the format into the script literal (backslash-escape `\` and `"`).
  local fmt_to_apply = (next_key == "GENERAL") and "General" or next_key
  local escaped = fmt_to_apply:gsub("\\", "\\\\"):gsub('"', '\\"')
  local ok, result = M.applescript([[
    tell application "Microsoft Excel"
      try
        set number format of selection to "]] .. escaped .. [["
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]])

  if ok and type(result) == "string" and result:find("^ERROR ") then
    if _G.__mme_log then
      _G.__mme_log("cycle_number_format: %s", result)
    end
    hs.alert.show("Number format failed (see log)", 1.2)
  end
end

-- Cycle the horizontal alignment of the current selection through
-- config.alignment_cycle.
--
-- Semantics mirror cycle_number_format: read the first cell's
-- alignment, advance to the next entry (wrapping), apply to the whole
-- selection. Recognised tokens are normalised in `tokens` below; only
-- the canonical AppleScript enum names we control are ever spliced into
-- the script, so config values can't inject AppleScript.
--
-- Excel for Mac silently refuses horizontal-alignment writes while a
-- modal dialog (e.g. Format Cells) is open, returning a -10006 that
-- reads like a parse error ("Can't set selection to ..."). If a press
-- seems to do nothing, check for an open dialog before suspecting this
-- code.
function M.cycle_alignment()
  -- token -> { apply = <enum to write>, read = <enum as text on read> }
  local tokens = {
    left   = { apply = "horizontal align left",   read = "horizontal align left" },
    right  = { apply = "horizontal align right",  read = "horizontal align right" },
    center = { apply = "horizontal align center", read = "horizontal align center" },
    ["center across selection"] = {
      apply = "horizontal align center across selection",
      read  = "horizontal align center across selection",
    },
    none = { apply = "horizontal align general", read = "horizontal align general" },
  }

  local function canon(s)
    s = tostring(s):lower():gsub("^%s+", ""):gsub("%s+$", "")
    if s == "centre" then s = "center" end
    if s == "general" then s = "none" end
    return s
  end

  local cycle_norm = {}
  for _, entry in ipairs(config.alignment_cycle or {}) do
    local key = canon(entry)
    if tokens[key] then
      cycle_norm[#cycle_norm + 1] = key
    elseif _G.__mme_log then
      _G.__mme_log("cycle_alignment: ignoring unknown entry %q", tostring(entry))
    end
  end
  if #cycle_norm == 0 then
    if _G.__mme_log then
      _G.__mme_log("cycle_alignment: no valid entries in config.alignment_cycle")
    end
    return
  end

  local ok_read, result = hs.osascript.applescript([[
    tell application "Microsoft Excel"
      try
        return horizontal alignment of (cell 1 of selection) as text
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]])

  local current_key
  if ok_read and type(result) == "string" and not result:find("^ERROR ") then
    for key, info in pairs(tokens) do
      if info.read == result then
        current_key = key
        break
      end
    end
  end

  local next_key = cycle_norm[1]
  if current_key then
    for i, key in ipairs(cycle_norm) do
      if key == current_key then
        next_key = cycle_norm[(i % #cycle_norm) + 1]
        break
      end
    end
  end

  local ok, apply_result = M.applescript(string.format([[
    tell application "Microsoft Excel"
      try
        set horizontal alignment of selection to %s
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]], tokens[next_key].apply))

  if ok and type(apply_result) == "string" and apply_result:find("^ERROR ") then
    if _G.__mme_log then
      _G.__mme_log("cycle_alignment: %s", apply_result)
    end
    hs.alert.show("Alignment failed (see log)", 1.2)
    return
  end

  hs.alert.show("Align: " .. next_key, 0.7)
end

-- Increase / decrease the decimal places shown by the selection while
-- preserving the rest of its number format. Mirrors Excel's
-- Increase/Decrease Decimal toolbar buttons (Ctrl+Shift+, and
-- Ctrl+Shift+. by default).
--
-- We read the active cell's format code and rewrite only the decimal
-- placeholder run, leaving currency symbols, thousands separators,
-- percent signs, padding codes, and the negative/zero/text sections
-- untouched. The work is pure string manipulation in Lua (no Excel
-- "round decimals" command exists in the dictionary), so the helpers
-- below must skip over literal regions — quoted text, escaped chars,
-- [color]/[condition] brackets, and _/* width-skip codes — so a "."
-- or a "0" inside them is never mistaken for a real decimal point or
-- placeholder.

-- Classify each byte of one format section as a numeric placeholder
-- ("ph": 0 # ?), the decimal separator ("dot": an active "."), or
-- something we must not touch. Returns a per-index table of marks.
local function classify_section(section)
  local n = #section
  local marks = {}
  local i = 1
  local in_quote = false
  while i <= n do
    local c = section:sub(i, i)
    if in_quote then
      marks[i] = "lit"
      if c == '"' then in_quote = false end
      i = i + 1
    elseif c == '"' then
      in_quote = true; marks[i] = "lit"; i = i + 1
    elseif c == "\\" then
      -- Backslash escapes the following char as a literal.
      marks[i] = "lit"
      if i + 1 <= n then marks[i + 1] = "lit" end
      i = i + 2
    elseif c == "_" or c == "*" then
      -- _x skips the width of x; *x repeats x. The next char is literal.
      marks[i] = "lit"
      if i + 1 <= n then marks[i + 1] = "lit" end
      i = i + 2
    elseif c == "[" then
      -- [Red], [>0], [$-409] ... consume through the closing bracket.
      marks[i] = "lit"; i = i + 1
      while i <= n and section:sub(i, i) ~= "]" do marks[i] = "lit"; i = i + 1 end
      if i <= n then marks[i] = "lit"; i = i + 1 end
    elseif c == "0" or c == "#" or c == "?" then
      marks[i] = "ph"; i = i + 1
    elseif c == "." then
      marks[i] = "dot"; i = i + 1
    else
      marks[i] = "other"; i = i + 1
    end
  end
  return marks
end

-- Split a full format code into its ;-separated sections without
-- splitting on a ";" that lives inside a literal region.
local function split_format_sections(fmt)
  local sections = {}
  local buf = {}
  local n = #fmt
  local i = 1
  local in_quote = false
  while i <= n do
    local c = fmt:sub(i, i)
    if in_quote then
      buf[#buf + 1] = c
      if c == '"' then in_quote = false end
      i = i + 1
    elseif c == '"' then
      in_quote = true; buf[#buf + 1] = c; i = i + 1
    elseif c == "\\" then
      buf[#buf + 1] = c
      if i + 1 <= n then buf[#buf + 1] = fmt:sub(i + 1, i + 1) end
      i = i + 2
    elseif c == "[" then
      buf[#buf + 1] = c; i = i + 1
      while i <= n and fmt:sub(i, i) ~= "]" do buf[#buf + 1] = fmt:sub(i, i); i = i + 1 end
      if i <= n then buf[#buf + 1] = fmt:sub(i, i); i = i + 1 end
    elseif c == ";" then
      sections[#sections + 1] = table.concat(buf); buf = {}; i = i + 1
    else
      buf[#buf + 1] = c; i = i + 1
    end
  end
  sections[#sections + 1] = table.concat(buf)
  return sections
end

-- Adjust the decimal placeholders of a single section by one step.
-- direction > 0 adds a decimal, < 0 removes one. Sections with no
-- numeric placeholders (e.g. a literal "-" text section) are returned
-- unchanged so the rest of the format survives intact.
local function adjust_section_decimals(section, direction)
  -- General has no placeholders to grow; give it a defined landing spot
  -- (one decimal up, integer down) matching Excel's Increase/Decrease.
  if section:gsub("%s", ""):lower() == "general" then
    return (direction > 0) and "0.0" or "0"
  end

  local marks = classify_section(section)
  local n = #section

  local dot_idx
  for i = 1, n do
    if marks[i] == "dot" then dot_idx = i; break end
  end

  if direction > 0 then
    if dot_idx then
      local last = dot_idx
      local j = dot_idx + 1
      while j <= n and marks[j] == "ph" do last = j; j = j + 1 end
      return section:sub(1, last) .. "0" .. section:sub(last + 1)
    end
    local last_ph
    for i = 1, n do
      if marks[i] == "ph" then last_ph = i end
    end
    if not last_ph then return section end
    return section:sub(1, last_ph) .. ".0" .. section:sub(last_ph + 1)
  end

  -- direction < 0
  if not dot_idx then return section end
  local last = dot_idx
  local j = dot_idx + 1
  while j <= n and marks[j] == "ph" do last = j; j = j + 1 end
  if last == dot_idx then
    -- A bare "." with nothing after it: drop it.
    return section:sub(1, dot_idx - 1) .. section:sub(dot_idx + 1)
  end
  if last - dot_idx <= 1 then
    -- Last remaining decimal placeholder: drop it and the "." together.
    return section:sub(1, dot_idx - 1) .. section:sub(last + 1)
  end
  return section:sub(1, last - 1) .. section:sub(last + 1)
end

local function adjust_format_decimals(fmt, direction)
  local parts = split_format_sections(fmt)
  for i, sec in ipairs(parts) do
    parts[i] = adjust_section_decimals(sec, direction)
  end
  return table.concat(parts, ";")
end

function M.step_decimal_places(direction)
  local ok_read, current = hs.osascript.applescript([[
    tell application "Microsoft Excel"
      try
        return number format of (cell 1 of selection)
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]])

  if not (ok_read and type(current) == "string") or current:find("^ERROR ") then
    if _G.__mme_log then
      _G.__mme_log("step_decimal_places: read failed: %s", tostring(current))
    end
    hs.alert.show("Decimal places failed (see log)", 1.2)
    return
  end

  local new_fmt = adjust_format_decimals(current, direction)
  if new_fmt == current then return end  -- nothing to change (e.g. no decimals to drop)

  local escaped = new_fmt:gsub("\\", "\\\\"):gsub('"', '\\"')
  local ok, result = M.applescript([[
    tell application "Microsoft Excel"
      try
        set number format of selection to "]] .. escaped .. [["
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]])

  if ok and type(result) == "string" and result:find("^ERROR ") then
    if _G.__mme_log then
      _G.__mme_log("step_decimal_places: %s", result)
    end
    hs.alert.show("Decimal places failed (see log)", 1.2)
  end
end

function M.decimal_places_up()   M.step_decimal_places( 1) end
function M.decimal_places_down() M.step_decimal_places(-1) end

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

----------------------------------------------------------------------
-- Expand the selected-sheet group (group adjacent worksheets)
----------------------------------------------------------------------

-- AppleScript can read `selected sheets` but not write a tab group on
-- current Mac builds; the reliable path is a Shift-click on the target
-- tab (frame from the AX tree, never hardcoded coordinates). Group
-- extent comes from AppleScript (AX "Selected," only marks the active
-- tab); sheet count and tab frame from AX "Sheet N of M" labels.
--
-- Like Shift+Arrow: active sheet is the anchor, each press moves the
-- lead edge (the non-anchor endpoint) one tab in the pressed direction.
local function expand_sheet_selection(direction)
  local ok, result = M.applescript([[
    tell application "Microsoft Excel"
      try
        set selSheets to selected sheets of active window
        set lo to index of (item 1 of selSheets)
        set hi to lo
        repeat with s in selSheets
          set i to index of s
          if i < lo then set lo to i
          if i > hi then set hi to i
        end repeat
        set a to index of active sheet
        return (a as string) & "," & (lo as string) & "," & (hi as string)
      on error errMsg number errNum
        return "ERROR " & errNum & ": " & errMsg
      end try
    end tell
  ]])
  if not ok then return end
  local active, lo, hi = tostring(result):match("^(%d+),(%d+),(%d+)$")
  active, lo, hi = tonumber(active), tonumber(lo), tonumber(hi)
  if not (active and lo and hi) then
    if _G.__mme_log then
      _G.__mme_log("expand_sheet_selection: unexpected group read %q", tostring(result))
    end
    return
  end

  local lead
  if lo == hi then
    lead = active
  elseif active == lo then
    lead = hi
  elseif active == hi then
    lead = lo
  else
    lead = (direction == "next") and hi or lo
  end

  local target_pos = (direction == "next") and (lead + 1) or (lead - 1)
  if target_pos < 1 then
    hs.alert.show("Already at first sheet", 1.0)
    return
  end

  local ax = hs.axuielement
  if not ax then
    local ok_req, mod = pcall(require, "hs.axuielement")
    if ok_req then ax = mod end
  end
  if not ax then
    if _G.__mme_log then _G.__mme_log("expand_sheet_selection: hs.axuielement unavailable") end
    hs.alert.show("Accessibility unavailable", 1.2)
    return
  end

  local app = hs.application.find(config.excel_bundle_id)
  if not app then return end
  local app_ax = ax.applicationElement(app)
  if not app_ax then return end

  -- AX attribute access can throw on unexpected element kinds.
  local function attr(elem, name)
    local got, val = pcall(function() return elem[name] end)
    return got and val or nil
  end

  local total, target_frame
  local function walk(elem)
    if target_frame and total then return end
    if attr(elem, "AXRole") == "AXButton" then
      local desc = tostring(attr(elem, "AXDescription") or "")
      local pos, m = desc:match("Sheet (%d+) of (%d+)")
      if pos then
        total = total or tonumber(m)
        if tonumber(pos) == target_pos then
          local frame = attr(elem, "AXFrame")
          if not frame then
            local p, s = attr(elem, "AXPosition"), attr(elem, "AXSize")
            if p and s then frame = { x = p.x, y = p.y, w = s.w, h = s.h } end
          end
          target_frame = frame
        end
      end
    end
    for _, child in ipairs(attr(elem, "AXChildren") or {}) do
      walk(child)
    end
  end
  for _, w in ipairs(attr(app_ax, "AXWindows") or {}) do
    walk(w)
  end

  if total and target_pos > total then
    hs.alert.show("Already at last sheet", 1.0)
    return
  end

  if not target_frame then
    if _G.__mme_log then
      _G.__mme_log("expand_sheet_selection: target tab %d not found/visible", target_pos)
    end
    hs.alert.show("Adjacent sheet tab not visible", 1.2)
    return
  end

  local f = target_frame
  local fw = f.w or f.width or 0
  local fh = f.h or f.height or 0
  local point = { x = f.x + fw / 2, y = f.y + fh / 2 }

  -- Shift flag only on the synthesised click (user may hold Alt+Shift).
  local e = hs.eventtap.event
  e.newMouseEvent(e.types.leftMouseDown, point):setFlags({ shift = true }):post()
  hs.timer.usleep(math.floor(config.sheet_group_shift_click_gap_seconds * 1e6))
  e.newMouseEvent(e.types.leftMouseUp, point):setFlags({ shift = true }):post()
end

function M.expand_sheet_selection_next() expand_sheet_selection("next") end
function M.expand_sheet_selection_prev() expand_sheet_selection("prev") end

return M
