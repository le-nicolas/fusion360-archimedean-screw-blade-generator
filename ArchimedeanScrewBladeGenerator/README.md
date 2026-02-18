# Archimedean Screw Blade Generator (Add-in)

Fusion 360 add-in that creates configurable Archimedean screw turbine blades around an existing shaft.

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
3. `Outer Radius`
4. `Blade Length`
5. `Turns`
6. `Turns (drag)` slider
7. `Blade Thickness`
8. `Hub Clearance`
9. `Start Angle`
10. `Handedness`
11. `Flights`
12. `Segments / Turn`
13. `Operation` (`New Blade Body` / `Join Blade To Shaft`)

## Notes
- Only the shaft cylindrical face needs to be selected.
- Start plane is auto-detected from planar shaft end faces connected to the selected cylinder.
- Command supports live preview while you edit values.
- If screw direction is opposite, switch `Handedness` or rotate `Start Angle`.
- Higher `Segments / Turn` gives smoother blades but heavier features.

## Files
- `ArchimedeanScrewBladeGenerator.py`: main add-in logic.
- `ArchimedeanScrewBladeGenerator.manifest`: Fusion manifest.
- `resources/`: command icons.
- `assets/prompt-reference.png`: diagram reference.
