-- Mouseless Excel: Hammerspoon bootstrap template.
--
-- This file is meant to live at ~/.hammerspoon/init.lua. Its only job
-- is to point Hammerspoon at the project clone and reload itself when
-- any project file is saved, so the project folder stays the source
-- of truth.
--
-- The PROJECT path below is filled in by install.sh. If you prefer to
-- set things up by hand, replace it with the absolute path to your
-- local clone, e.g. "/Users/yourname/code/mouseless-excel".
--
-- If you already have a ~/.hammerspoon/init.lua for other purposes,
-- copy the contents of this file into it rather than overwriting.

local PROJECT = "__MOUSELESS_EXCEL_PATH__"
local LOG_PATH = "/tmp/mouseless-excel.log"

local function logf(fmt, ...)
  local line = "[" .. os.date("%Y-%m-%d %H:%M:%S") .. "] " .. string.format(fmt, ...)
  print("[mouseless-excel] " .. line)
  local f = io.open(LOG_PATH, "a")
  if f then f:write(line .. "\n"); f:close() end
end
_G.__mle_log = logf
logf("bootstrap start (HS=%s)", hs and hs.processInfo and hs.processInfo.version or "?")

package.path = package.path
  .. ";" .. PROJECT .. "/?.lua"
  .. ";" .. PROJECT .. "/?/init.lua"

-- Drop cached project modules so dofile() picks up the latest source.
local PROJECT_MODULES = { "init", "config", "actions", "shortcuts", "runtime" }
for _, name in ipairs(PROJECT_MODULES) do
  package.loaded[name] = nil
end

local ok, err = pcall(dofile, PROJECT .. "/init.lua")
if not ok then
  hs.alert.show("Mouseless Excel load error: " .. tostring(err), 4)
  logf("load error: %s", tostring(err))
else
  logf("loaded ok; accessibility=%s", tostring(hs.accessibilityState()))
end

-- Auto-reload Hammerspoon whenever a .lua file in the project changes.
-- Saving shortcuts.lua or actions.lua is therefore a complete redeploy.
if not _G.__mle_pathwatcher then
  _G.__mle_pathwatcher = hs.pathwatcher.new(PROJECT, function(files)
    for _, f in ipairs(files) do
      if f:match("%.lua$") then
        hs.alert.show("Reloading Mouseless Excel", 0.6)
        hs.reload()
        return
      end
    end
  end)
  _G.__mle_pathwatcher:start()
end

hs.alert.show("Mouseless Excel ready", 0.8)
