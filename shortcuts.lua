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
    { keys = { "e", "s", "v" }, action = "paste_values",        desc = "Paste Special > Values" },
    { keys = { "e", "s", "t" }, action = "paste_formats",       desc = "Paste Special > Formats" },
    { keys = { "e", "s", "w" }, action = "paste_column_widths", desc = "Paste Special > Column Widths" },

    { keys = { "o", "c", "w" }, action = "column_width_dialog", desc = "Format > Column > Width..." },
    { keys = { "o", "r", "e" }, action = "row_height_dialog",   desc = "Format > Row > Height..." },
    { keys = { "o", "h", "r" }, action = "rename_sheet",        desc = "Format > Sheet > Rename" },
    { keys = { "o", "h", "m" }, action = "move_sheet_dialog",   desc = "Edit > Sheet > Move or Copy..." },
    { keys = { "o", "w", "s" }, action = "insert_sheet",        desc = "Insert new sheet (after active)" },
    { keys = { "e", "l" },      action = "delete_sheet",        desc = "Delete active sheet (confirms)" },
    { keys = { "b", "o" },      action = "border_dotted",       desc = "Reapply borders as dash (light/heavy by weight)" },
  },

  combos = {
    { mods = { "cmd", "shift" },  key = "v", action = "paste_values",     desc = "Paste Values" },
    { mods = { "ctrl", "shift" }, key = "c", action = "cycle_font_color", desc = "Cycle font color" },
    { mods = { "ctrl", "shift" }, key = "b", action = "cycle_border_thin",  desc = "Cycle border (thin)" },
    { mods = { "ctrl", "shift" }, key = "t", action = "cycle_border_thick", desc = "Cycle border (thick)" },
    { mods = { "ctrl", "shift" }, key = "a", action = "font_size_up",     desc = "Font size up" },
    { mods = { "ctrl", "shift" }, key = "s", action = "font_size_down",   desc = "Font size down" },
    { mods = { "ctrl", "shift" }, key = "v", action = "cycle_fill_color", desc = "Cycle fill color" },
    { mods = { "ctrl", "shift" }, key = "i", action = "zoom_in",          desc = "Zoom in" },
    { mods = { "ctrl", "shift" }, key = "j", action = "zoom_out",         desc = "Zoom out" },
    { mods = { "ctrl" }, key = "space", action = "select_column", desc = "Select current column" },
    { mods = { "shift", "alt" },   key = "down",  action = "next_sheet",    desc = "Activate next sheet" },
    { mods = { "shift", "alt" },   key = "up",    action = "prev_sheet",    desc = "Activate previous sheet" },
  },
}
