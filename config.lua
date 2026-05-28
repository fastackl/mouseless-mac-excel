-- Tunable settings for the Mouseless Mac Excel plugin.
-- Edit values here, save the file, and Hammerspoon will auto-reload.

return {
  -- Bundle ID Hammerspoon uses to detect "Excel is frontmost".
  excel_bundle_id = "com.microsoft.Excel",

  -- Maximum time (seconds) between Option-down and Option-up to count as
  -- a "tap alone" that enters menu mode. Increase if your taps are slow.
  leader_tap_max_seconds = 0.6,

  -- After entering menu mode, how long with no key press before we
  -- silently exit (mirrors Windows behaviour where Alt times out).
  menu_idle_timeout_seconds = 3.0,

  -- Delay between scripted keystrokes inside multi-step actions
  -- (e.g. opening Paste Special, then pressing V, then Return).
  -- Bump this up if Excel feels slow to open dialogs on your machine.
  step_delay_seconds = 0.18,

  -- When true, dialog-opening actions that pre-fill a numeric value
  -- (e.g. Column Width) will synthesise a mouse click on the dialog's
  -- text field and select the existing value, so the user can type a
  -- new value and have it replace the old one — matching Windows
  -- Excel's overtype UX.
  --
  -- This works around an AppKit quirk where a programmatically opened
  -- dialog reports AXFocused=true on its text field but hasn't yet
  -- promoted it to first responder, so the user's first keystroke is
  -- eaten. A real synthesised click is currently the only reliable
  -- way to force the responder chain to update.
  --
  -- Costs: the mouse cursor briefly flickers to the field and back.
  -- Turn off if you'd rather have the unobstructed cursor and don't
  -- mind tapping the field manually before typing.
  dialog_focus_click = true,

  -- Delay (seconds) between opening a dialog and looking it up in the
  -- AX tree to click its text field. 80 ms is comfortable on a fast
  -- Mac; bump it if you see "dialog not found" messages in the log
  -- (the dialog hadn't rendered yet when we went looking).
  dialog_focus_click_delay_seconds = 0.08,

  -- Bounds and step for the Zoom In / Zoom Out shortcuts (Ctrl+Shift+I
  -- and Ctrl+Shift+J by default). Each press reads Excel's current
  -- active-window zoom and lands on the next grid line of zoom_step
  -- in the direction of the press, clamped to [zoom_min, zoom_max].
  -- All values are integer percentages.
  --
  -- Snap behaviour: zoom-in from 117% lands on 120%; from 120% lands
  -- on 130%. Zoom-out from 117% lands on 110%; from 120% lands on
  -- 110%. The grid is defined by zoom_step itself, so changing
  -- zoom_step changes the grid live.
  zoom_min  = 50,
  zoom_max  = 200,
  zoom_step = 10,

  -- Font colours that the Cycle Font Color shortcut (Ctrl+Shift+C by
  -- default) walks through, in order. Each entry is a 6-character hex
  -- string (case-insensitive; a leading "#" is tolerated). Pressing
  -- the shortcut advances the current selection's font colour from
  -- its current entry to the next one in the list, wrapping at the
  -- end. If the current colour isn't in the list (typically because
  -- the cell is still Excel's default black), the first entry is
  -- applied instead. Edit freely — order is significant.
  font_color_cycle = { "0433FF", "FF2600", "008F00", "000000" },

  -- Fill colours (cell background) that the Cycle Fill Color shortcut
  -- (Ctrl+Shift+V by default) walks through, in order. Each entry is
  -- either a 6-character hex string (case-insensitive; a leading "#"
  -- is tolerated) or the literal string "none" to clear the fill.
  -- Pressing the shortcut advances the selection's current fill from
  -- its current entry to the next one in the list, wrapping at the
  -- end. If the current fill isn't in the list, the first entry is
  -- applied instead. Edit freely — order is significant.
  fill_color_cycle = { "BFF7FA", "D9D9D9", "none" },

  -- When true, show small on-screen alerts as you type a sequence and
  -- log every action to the Hammerspoon console. Helpful while we
  -- iterate; turn off once shortcuts feel stable.
  debug = true,
}
