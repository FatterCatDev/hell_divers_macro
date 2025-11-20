# **If hotkeys do not fire, try running the exe as Administrator (or make sure Steam and HELLDIVERS 2 are not running as Administrator).**

# HELLDIVERS 2 Stratagem Macro

Creator: FatterCatDev

A simple Windows app that lets you trigger HELLDIVERS 2 stratagems with your numpad. Pick a template for each numpad key, start listening, and the app will play the stratagem inputs for you.

## [Download the macro here!](https://github.com/FatterCatDev/hell_divers_macro/releases/tag/helldivers_2_macro)

## Quick start (exe)
1) Run `helldivers_macro.exe` (one-file build, no install needed).
2) In the grid, click each numpad slot (7-9 / 4-6 / 1-3) and choose a stratagem.
3) Click **Start Listening**. Press the matching numpad key in-game to fire the macro.
   - Exit hotkey: `Ctrl+Shift+Q`.
4) Profiles save in the `saves/` folder next to the exe; the last profile reloads automatically.

## In-game presetup (recommended)
Set these in HELLDIVERS 2 before using the macro (they match the app defaults and can be changed later in Settings):
- Up: Up Arrow
- Down: Down Arrow
- Left: Left Arrow
- Right: Right Arrow
- Stratagem List (panel): Home **(Make it Key Press instead of Hold)**

## Controls and options
- **Auto Stratagem Panel**: Toggle on to press the panel key (default `Home`) before every macro. Change the panel key via Settings.
- **Settings (File > Settings)**:
  - Slot Hotkeys: rebind each numpad slot.
  - Direction Keys: rebind the Up/Down/Left/Right inputs used inside a stratagem.
  - Macro Delay: adjust delay and duration between key presses.
- **Edit Stratagem Templates (Edit > Edit Stratagem Templates)**:
  - Search bar filters by name or category.
  - Select a template, edit its category or the comma-separated direction sequence, click Apply, then **Save Templates** to persist.
- **Log**: Shows macro activity and listener status.

## Tips
- If hotkeys do not fire, try running the exe as Administrator (but keep Steam and HELLDIVERS 2 **not** elevated).
- Arrow keys will not trigger numpad bindings; macros can fire even while other keys (e.g., Shift/W) are held.
- Keep the exe in its own folder so the generated `saves/` stay contained.
- Saves stay local in `saves/`; delete `.last_profile` there if you want a clean start.
