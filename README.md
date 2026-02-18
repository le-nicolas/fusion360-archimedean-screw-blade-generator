# Archimedean Screw Blade Generator (Fusion 360)

![Prompt reference](ArchimedeanScrewBladeGenerator/assets/prompt-reference.png)

## The frustration that started this
I had already built the shaft.
Then I got stuck on the blade, because the blade is the math-heavy part of an Archimedean screw: helical geometry, pitch, handedness, start angle, and keeping it manufacturable.
This add-in exists so that part is no longer a manual sketch nightmare.

## What this add-in does

<img width="1275" height="884" alt="image" src="https://github.com/user-attachments/assets/0465abd4-b521-4c88-b504-b4648143ab98" />


- Builds configurable hydraulic Archimedean screw flight geometry around an existing shaft.
- Auto-detects the start plane from shaft end geometry, so you only select the shaft once.
- Shows a live preview while editing.
- Includes a draggable `Turns (drag)` slider for fast tuning.
- Includes a draggable `Bucket Wrap (drag)` slider to tune the hydraulic bucket cup.
- Supports `New Body` or direct `Join` into your shaft.

## Configurable parameters
- `Outer Radius`
- `Blade Length`
- `Turns`
- `Turns (drag)` slider
- `Start End` (auto-detected end-cap choice)
- `Blade Thickness`
- `Hub Clearance`
- `Bucket Wrap (deg)` + drag slider
- `Start Angle`
- `Handedness` (left/right)
- `Flights` (multi-start screw)

## Installation
1. Open Fusion 360.
2. Go to `Utilities -> Add-Ins -> Scripts and Add-Ins`.
3. In the `Add-Ins` tab, click `+`.
4. Select folder: `ArchimedeanScrewBladeGenerator`.
5. Run `Archimedean Screw Blade` from the `Create` panel.

## Project layout
- `ArchimedeanScrewBladeGenerator/ArchimedeanScrewBladeGenerator.py`: add-in command and geometry generation.
- `ArchimedeanScrewBladeGenerator/ArchimedeanScrewBladeGenerator.manifest`: Fusion add-in manifest.
- `ArchimedeanScrewBladeGenerator/resources/`: toolbar icons.
- `ArchimedeanScrewBladeGenerator/assets/prompt-reference.png`: reference screw diagram.

## Cross-reference used before publish
I cross-referenced your local Fusion add-in `HelixGenerator` at:
`%APPDATA%\Autodesk\Autodesk Fusion 360\API\AddIns\HelixGenerator\HelixGenerator`
and aligned this repo with common Fusion add-in patterns:
- add-in manifest + main Python entrypoint
- toolbar resource icon set (`16/32/64`, dark, disabled)
- installation-focused README documentation

## Reference image attribution
Diagram source: Wikimedia Commons (`ArchimedesSketch.svg`).
