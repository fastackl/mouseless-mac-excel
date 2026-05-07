-- Shortcut declarations.
--
-- This is the file you'll edit most. Two sections:
--
--   sequences  -- multi-key Windows-style sequences. After tapping the
--                Option key alone (the "leader"), type the letters in
--                `keys`. Example: { "e", "s", "v" } means Alt,E,S,V.
--
--   combos     -- single keyboard combinations. `mods` is a list of
--                modifiers ("cmd", "ctrl", "alt", "shift"); `key` is
--                the key. Active only when Excel is frontmost.
--
-- `action` is the name of a function in actions.lua.
-- `desc`   is human-readable text shown in alerts and the console.

return {
  sequences = {
    { keys = { "e", "s", "v" }, action = "paste_values", desc = "Paste Special > Values" },
  },

  combos = {
    { mods = { "cmd", "shift" }, key = "v", action = "paste_values", desc = "Paste Values" },
  },
}
