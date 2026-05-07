-- Runtime engine for the Mouseless Excel plugin.
--
-- Responsibilities:
--   1. Watch macOS for Microsoft Excel becoming/leaving frontmost.
--   2. While Excel is frontmost:
--        - Bind single-combo hotkeys from shortcuts.combos.
--        - Run an event tap that detects "Option tapped alone" and uses
--          that as the leader to enter menu mode.
--   3. In menu mode, walk a tree built from shortcuts.sequences as
--      keys are pressed; on a leaf node, run the matching action.
--
-- You normally don't need to edit this file. Edit shortcuts.lua to add
-- shortcuts and actions.lua when an action needs new logic.

local config = require("config")

local M = {}

----------------------------------------------------------------------
-- State
----------------------------------------------------------------------

local app_watcher = nil
local event_tap = nil
local combo_hotkeys = {}

local sequence_modal = nil
local seq_tree = nil
local seq_pos = nil
local seq_crumbs = ""
local idle_timer = nil

local opt_was_down = false
local opt_down_at = 0
local opt_alone = false

----------------------------------------------------------------------
-- Logging
----------------------------------------------------------------------

local function log(msg)
  if config.debug then
    if _G.__mle_log then
      _G.__mle_log("%s", msg)
    else
      print("[mouseless-excel] " .. msg)
    end
  end
end

local function alert(msg, dur)
  if config.debug then
    hs.alert.show(msg, dur or 0.5)
  end
end

----------------------------------------------------------------------
-- Tree builder: turn shortcuts.sequences into a nested lookup table.
-- Leaves are { __action = fn, __desc = "..." }.
----------------------------------------------------------------------

local function build_tree(sequences, actions)
  local tree = {}
  for _, s in ipairs(sequences) do
    local fn = actions[s.action]
    if not fn then
      log("WARN unknown action: " .. tostring(s.action) ..
          " (sequence " .. table.concat(s.keys, ",") .. ")")
    else
      local node = tree
      local n = #s.keys
      for i, k in ipairs(s.keys) do
        k = string.lower(k)
        if i == n then
          if node[k] and node[k].__action == nil then
            log("WARN sequence prefix conflict at " .. table.concat(s.keys, ","))
          end
          node[k] = { __action = fn, __desc = s.desc or s.action }
        else
          if node[k] and node[k].__action then
            log("WARN sequence prefix conflict at " .. table.concat(s.keys, ","))
            break
          end
          node[k] = node[k] or {}
          node = node[k]
        end
      end
    end
  end
  return tree
end

----------------------------------------------------------------------
-- Sequence modal (menu mode)
----------------------------------------------------------------------

local function reset_idle_timer()
  if idle_timer then idle_timer:stop() end
  idle_timer = hs.timer.doAfter(config.menu_idle_timeout_seconds, function()
    M.exit_menu("timeout")
  end)
end

function M.enter_menu()
  if not seq_tree or not sequence_modal then return end
  seq_pos = seq_tree
  seq_crumbs = ""
  sequence_modal:enter()
  alert("Excel menu", 0.4)
  reset_idle_timer()
end

function M.exit_menu(reason)
  if sequence_modal then sequence_modal:exit() end
  seq_pos = nil
  seq_crumbs = ""
  if idle_timer then idle_timer:stop(); idle_timer = nil end
  if reason == "nomatch" then alert("X no match", 0.3) end
end

local function handle_modal_key(c)
  if not seq_pos then return end
  c = string.lower(c)
  seq_crumbs = seq_crumbs .. c
  local node = seq_pos[c]
  if node == nil then
    M.exit_menu("nomatch")
    return
  end
  if node.__action then
    local action = node.__action
    local desc = node.__desc or ""
    M.exit_menu("match")
    alert(seq_crumbs .. "  " .. desc, 0.5)
    local ok, err = pcall(action)
    if not ok then
      log("ERROR running action: " .. tostring(err))
      hs.alert.show("Action error: " .. tostring(err), 1.5)
    end
    return
  end
  seq_pos = node
  alert(seq_crumbs, 0.3)
  reset_idle_timer()
end

local function build_modal()
  local m = hs.hotkey.modal.new()
  for c = string.byte("a"), string.byte("z") do
    local key = string.char(c)
    m:bind({}, key, function() handle_modal_key(key) end)
  end
  for c = string.byte("0"), string.byte("9") do
    local key = string.char(c)
    m:bind({}, key, function() handle_modal_key(key) end)
  end
  m:bind({}, "escape", function() M.exit_menu("escape") end)
  return m
end

----------------------------------------------------------------------
-- Leader: detect "Option key tapped alone".
--
-- We watch flagsChanged + keyDown without consuming events, so normal
-- Option-as-modifier behaviour (Option+E for accents, etc.) still works.
-- A "tap alone" is: Option goes down with no other modifier, no other
-- key is pressed while it is held, and it goes back up within
-- leader_tap_max_seconds.
----------------------------------------------------------------------

local function build_event_tap()
  local types = hs.eventtap.event.types
  return hs.eventtap.new({ types.flagsChanged, types.keyDown }, function(e)
    local t = e:getType()
    if t == types.flagsChanged then
      local f = e:getFlags()
      local opt_now = f.alt == true
      local others = (f.cmd or f.ctrl or f.shift) and true or false
      if opt_now and not opt_was_down then
        opt_was_down = true
        opt_down_at = hs.timer.secondsSinceEpoch()
        opt_alone = not others
      elseif (not opt_now) and opt_was_down then
        opt_was_down = false
        local elapsed = hs.timer.secondsSinceEpoch() - opt_down_at
        local was_alone = opt_alone
        opt_alone = false
        if was_alone and elapsed < config.leader_tap_max_seconds then
          M.enter_menu()
        end
      end
    elseif t == types.keyDown then
      if opt_was_down then opt_alone = false end
    end
    return false
  end)
end

----------------------------------------------------------------------
-- Single-combo hotkeys
----------------------------------------------------------------------

local function build_combos(combos, actions)
  local hotkeys = {}
  for _, c in ipairs(combos) do
    local fn = actions[c.action]
    if not fn then
      log("WARN unknown action: " .. tostring(c.action))
    else
      local desc = c.desc or c.action
      local hk = hs.hotkey.new(c.mods or {}, c.key, function()
        alert(desc, 0.4)
        local ok, err = pcall(fn)
        if not ok then
          log("ERROR running action: " .. tostring(err))
          hs.alert.show("Action error: " .. tostring(err), 1.5)
        end
      end)
      table.insert(hotkeys, hk)
    end
  end
  return hotkeys
end

local function enable_combos()
  for _, hk in ipairs(combo_hotkeys) do hk:enable() end
end

local function disable_combos()
  for _, hk in ipairs(combo_hotkeys) do hk:disable() end
end

----------------------------------------------------------------------
-- App watcher: enable/disable everything based on Excel focus
----------------------------------------------------------------------

local function on_excel_focus()
  log("Excel focused; bindings active")
  enable_combos()
  if event_tap and not event_tap:isEnabled() then event_tap:start() end
end

local function on_excel_blur()
  log("Excel blurred; bindings inactive")
  M.exit_menu("blur")
  disable_combos()
  if event_tap and event_tap:isEnabled() then event_tap:stop() end
end

local function build_app_watcher()
  return hs.application.watcher.new(function(_, event_type, app)
    if not app or app:bundleID() ~= config.excel_bundle_id then return end
    if event_type == hs.application.watcher.activated then
      on_excel_focus()
    elseif event_type == hs.application.watcher.deactivated
        or event_type == hs.application.watcher.terminated then
      on_excel_blur()
    end
  end)
end

----------------------------------------------------------------------
-- Public API
----------------------------------------------------------------------

function M.start(shortcuts, actions)
  M.stop()

  seq_tree = build_tree(shortcuts.sequences or {}, actions)
  sequence_modal = build_modal()
  combo_hotkeys = build_combos(shortcuts.combos or {}, actions)
  event_tap = build_event_tap()
  app_watcher = build_app_watcher()
  app_watcher:start()

  local front = hs.application.frontmostApplication()
  if front and front:bundleID() == config.excel_bundle_id then
    on_excel_focus()
  end

  log("started: " .. #(shortcuts.sequences or {}) .. " sequences, " ..
      #(shortcuts.combos or {}) .. " combos")
end

function M.stop()
  if app_watcher then app_watcher:stop(); app_watcher = nil end
  if event_tap then event_tap:stop(); event_tap = nil end
  for _, hk in ipairs(combo_hotkeys) do hk:delete() end
  combo_hotkeys = {}
  if sequence_modal then sequence_modal:exit(); sequence_modal = nil end
  seq_tree = nil
  seq_pos = nil
end

return M
