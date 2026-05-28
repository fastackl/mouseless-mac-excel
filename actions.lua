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
function M.applescript(script)
  local ok, result = hs.osascript.applescript(script)
  if not ok then
    if _G.__mme_log then
      _G.__mme_log("applescript error: %s", tostring(result))
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

return M
