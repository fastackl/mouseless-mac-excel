-- Tunable settings for the Mouseless Excel plugin.
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

  -- When true, show small on-screen alerts as you type a sequence and
  -- log every action to the Hammerspoon console. Helpful while we
  -- iterate; turn off once shortcuts feel stable.
  debug = true,
}
