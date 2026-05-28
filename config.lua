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

  -- When true, show small on-screen alerts as you type a sequence and
  -- log every action to the Hammerspoon console. Helpful while we
  -- iterate; turn off once shortcuts feel stable.
  debug = true,
}
