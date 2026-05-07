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
    if _G.__mle_log then
      _G.__mle_log("applescript error: %s", tostring(result))
    end
    hs.alert.show("AppleScript error (see log)", 1.2)
  end
  return ok, result
end

-- Paste Special > Values
--
-- Driving the Mac Excel Paste Special dialog with synthetic keystrokes
-- is unreliable: macOS dialog radio buttons don't honour Windows-style
-- letter accelerators. Excel's AppleScript dictionary exposes the
-- operation directly, so we call it instead. No dialog, no timing.
function M.paste_values()
  M.applescript([[
    tell application "Microsoft Excel"
      paste special (get selection) what paste values
    end tell
  ]])
end

return M
