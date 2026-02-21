# Archimedean Screw Blade Generator (Add-in)

Fusion 360 add-in that creates configurable hydraulic Archimedean screw flights around an existing shaft.

## Install
1. In Fusion 360 open `Utilities -> Add-Ins -> Scripts and Add-Ins`.
2. Open the `Add-Ins` tab.
3. Click `+` and pick this folder:
   `c:\Users\User\Fusion360\ArchimedeanScrewBladeGenerator`
4. Run the add-in.

## Command location
- Workspace: `Design`
- Panel: `Create`
- Command: `Archimedean Screw Blade`

## Inputs
1. `Shaft Cylindrical Face`
2. `Start End` (auto-detected from shaft end caps)
3. `Outer Radius` (auto-suggested from selected shaft)
4. `Hub Clearance`
5. `Pitch Mode`
6. `Blade Length` (constant mode) or `Pitch Start` + `Pitch End` (variable mode)
7. `Turns` + `Turns (drag)` slider
8. `Bucket Wrap (deg)` + drag slider
9. `Hub Thickness`
10. `Thickness Mode` + optional `Tip Thickness`
11. `Start Angle`
12. `Handedness`
13. `Blade Preset` (`1/2/3/4/Custom`)
14. `Flights`
15. `Operation` (`New Blade Body` / `Join Blade To Shaft`)
16. `RPM` for engineering feedback
17. `Saved Preset`, `Preset Name`, and `Save Preset` button

## Notes
- Command includes strict validation with readable messages for invalid geometric combinations.
- Variable pitch mode builds a segmented tapered helix path between `Pitch Start` and `Pitch End`.
- Tapered thickness mode approximates thicker-at-hub and thinner-at-tip flights by radial banding.
- Derived panel shows helix angle, pitch ratio, and theoretical/estimated flow at the set `RPM`.
- User presets are saved to `presets.json` next to the add-in script and can be team-shared.
- Generated timeline features and bodies are named to make edits in parametric history easier.

## Files
- `ArchimedeanScrewBladeGenerator.py`: main add-in logic.
- `ArchimedeanScrewBladeGenerator.manifest`: Fusion manifest.
- `resources/`: command icons.
- `assets/prompt-reference.png`: diagram reference.
