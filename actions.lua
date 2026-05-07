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

function M.menu(path)
  local app = hs.application.find(config.excel_bundle_id)
  if not app then return false end
  return app:selectMenuItem(path)
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

return M
