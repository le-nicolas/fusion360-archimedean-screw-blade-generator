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
3. `Outer Radius`
4. `Blade Length`
5. `Turns`
6. `Turns (drag)` slider
7. `Blade Thickness`
8. `Hub Clearance`
9. `Bucket Wrap (deg)` and drag slider
10. `Start Angle`
11. `Handedness`
12. `Flights`
13. `Operation` (`New Blade Body` / `Join Blade To Shaft`)

## Notes
- Only the shaft cylindrical face needs to be selected.
- Start plane is auto-detected from planar shaft end faces connected to the selected cylinder.
- Command supports live preview while you edit values.
- `Bucket Wrap` controls how cupped the hydraulic bucket shape becomes.
- If screw direction is opposite, switch `Handedness` or rotate `Start Angle`.
- Geometry is built from helical guide wires + ruled surface + thicken for smoother, faster updates.

## Files
- `ArchimedeanScrewBladeGenerator.py`: main add-in logic.
- `ArchimedeanScrewBladeGenerator.manifest`: Fusion manifest.
- `resources/`: command icons.
- `assets/prompt-reference.png`: diagram reference.
