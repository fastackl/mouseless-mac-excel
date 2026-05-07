-- Entry point for the Mouseless Mac Excel plugin.
--
-- Loaded by ~/.hammerspoon/init.lua via dofile(). Returns a table so it
-- can also be required, but the side effect of loading is to start the
-- runtime with the current shortcuts and actions.

local shortcuts = require("shortcuts")
local actions   = require("actions")
local runtime   = require("runtime")

runtime.start(shortcuts, actions)

return {
  shortcuts = shortcuts,
  actions   = actions,
  runtime   = runtime,
}
