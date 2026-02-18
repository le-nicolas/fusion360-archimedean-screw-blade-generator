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
2. `Start Plane / Face` (must be normal to shaft axis)
3. `Outer Radius`
4. `Blade Length`
5. `Turns`
6. `Blade Thickness`
7. `Hub Clearance`
8. `Start Angle`
9. `Handedness`
10. `Flights`
11. `Segments / Turn`
12. `Operation` (`New Blade Body` / `Join Blade To Shaft`)

## Notes
- Shaft and start references must be in the same component.
- If screw direction is opposite, switch `Handedness` or rotate `Start Angle`.
- Higher `Segments / Turn` gives smoother blades but heavier features.

## Files
- `ArchimedeanScrewBladeGenerator.py`: main add-in logic.
- `ArchimedeanScrewBladeGenerator.manifest`: Fusion manifest.
- `resources/`: command icons.
- `assets/prompt-reference.png`: diagram reference.
