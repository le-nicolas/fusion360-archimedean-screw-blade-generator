import adsk.core
import adsk.fusion
import json
import math
import os
import traceback
from typing import Dict, List, Optional


APP = adsk.core.Application.get()
UI = APP.userInterface if APP else None

ADDIN_NAME = 'Archimedean Screw Blade Generator'
CMD_ID = 'archimedean_screw_blade_generator_cmd_v3'
CMD_NAME = 'Archimedean Screw Blade'
CMD_DESCRIPTION = 'Create a configurable hydraulic Archimedean screw flight around an existing shaft.'
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidCreatePanel'
RESOURCE_FOLDER = os.path.join(os.path.dirname(__file__), 'resources')
PRESET_FILE = os.path.join(os.path.dirname(__file__), 'presets.json')

# Input IDs
INPUT_SHAFT_FACE = 'shaftFace'
INPUT_START_END = 'startEnd'
INPUT_OUTER_RADIUS = 'outerRadius'
INPUT_LENGTH = 'bladeLength'
INPUT_TURNS = 'turns'
INPUT_TURNS_SLIDER = 'turnsSlider'
INPUT_THICKNESS = 'bladeThickness'
INPUT_CLEARANCE = 'hubClearance'
INPUT_BUCKET_WRAP = 'bucketWrapDeg'
INPUT_BUCKET_WRAP_SLIDER = 'bucketWrapSlider'
INPUT_START_ANGLE = 'startAngle'
INPUT_HANDEDNESS = 'handedness'
INPUT_FLIGHT_PRESET = 'flightPreset'
INPUT_FLIGHTS = 'flights'
INPUT_OPERATION = 'operation'
INPUT_DERIVED = 'derivedInfo'
INPUT_PITCH_MODE = 'pitchMode'
INPUT_PITCH_START = 'pitchStart'
INPUT_PITCH_END = 'pitchEnd'
INPUT_THICKNESS_MODE = 'thicknessMode'
INPUT_TIP_THICKNESS = 'tipThickness'
INPUT_RPM = 'rpm'
INPUT_SAVED_PRESET = 'savedPreset'
INPUT_PRESET_NAME = 'presetName'
INPUT_SAVE_PRESET = 'savePreset'
INPUT_PRESET_STATUS = 'presetStatus'

HAND_RIGHT = 'Right-handed'
HAND_LEFT = 'Left-handed'
OP_NEW_BODY = 'New Blade Body'
OP_JOIN = 'Join Blade To Shaft'
START_END_MIN = 'End 1 (auto)'
START_END_MAX = 'End 2 (auto)'

FLIGHT_PRESET_SINGLE = '1 Blade'
FLIGHT_PRESET_DOUBLE = '2 Blades'
FLIGHT_PRESET_TRIPLE = '3 Blades'
FLIGHT_PRESET_QUAD = '4 Blades'
FLIGHT_PRESET_CUSTOM = 'Custom'

PITCH_CONSTANT = 'Constant pitch'
PITCH_VARIABLE = 'Variable pitch (taper)'

THICKNESS_CONSTANT = 'Constant thickness'
THICKNESS_TAPERED = 'Tapered hub -> tip'

PRESET_MANUAL = 'Manual (current values)'
PRESET_BUILTIN_PREFIX = 'Built-in: '
PRESET_USER_PREFIX = 'User: '

_BUILTIN_PRESETS = {
    '1-flight standard': {
        'flightPreset': FLIGHT_PRESET_SINGLE,
        'flights': 1,
        'bucketWrapDeg': 35.0,
        'thicknessMode': THICKNESS_CONSTANT,
    },
    '2-flight standard': {
        'flightPreset': FLIGHT_PRESET_DOUBLE,
        'flights': 2,
        'bucketWrapDeg': 35.0,
        'thicknessMode': THICKNESS_CONSTANT,
    },
    '3-flight irrigation': {
        'flightPreset': FLIGHT_PRESET_TRIPLE,
        'flights': 3,
        'bucketWrapDeg': 45.0,
        'pitchMode': PITCH_VARIABLE,
        'pitchStart': 120.0 / 10.0,
        'pitchEnd': 90.0 / 10.0,
        'thicknessMode': THICKNESS_TAPERED,
        'tipThickness': 2.0 / 10.0,
    },
    'prototype thin-wall': {
        'thicknessMode': THICKNESS_TAPERED,
        'bladeThickness': 4.0 / 10.0,
        'tipThickness': 2.0 / 10.0,
        'bucketWrapDeg': 30.0,
        'pitchMode': PITCH_VARIABLE,
        'pitchStart': 160.0 / 10.0,
        'pitchEnd': 120.0 / 10.0,
    },
}

# Session state
_handler_registry: Dict[str, List[object]] = {}
_HELIX_SIGN_MODE: Optional[str] = None  # 'turns_sign' | 'pitch_sign' | 'flip_axis'
_user_presets: Dict[str, dict] = {}
_outer_radius_user_edited = False
_applying_preset = False

_VAR_PITCH_SEGMENTS = 24
_TAPER_THICKNESS_BANDS = 6
_TAPER_BAND_OVERLAP = 0.015


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def _active_design() -> Optional[adsk.fusion.Design]:
    return adsk.fusion.Design.cast(APP.activeProduct)


def _default_length_units() -> str:
    design = _active_design()
    return design.unitsManager.defaultLengthUnits if design else 'cm'


def _format_length(value: float) -> str:
    design = _active_design()
    if not design:
        return f'{value:.4f} cm'
    um = design.unitsManager
    return um.formatInternalValue(value, um.defaultLengthUnits, True)


def _clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def _selection_entity(inputs: adsk.core.CommandInputs, input_id: str):
    selection_input = adsk.core.SelectionCommandInput.cast(inputs.itemById(input_id))
    if not selection_input or selection_input.selectionCount < 1:
        return None
    return selection_input.selection(0).entity


def _select_dropdown_item(dropdown: adsk.core.DropDownCommandInput, item_name: str):
    if not dropdown:
        return
    for i in range(dropdown.listItems.count):
        item = dropdown.listItems.item(i)
        item.isSelected = (item.name == item_name)


def _dropdown_selected_name(inputs: adsk.core.CommandInputs, input_id: str, fallback: str = '') -> str:
    dd = adsk.core.DropDownCommandInput.cast(inputs.itemById(input_id))
    if not dd or not dd.selectedItem:
        return fallback
    return dd.selectedItem.name


def _value_input_value(inputs: adsk.core.CommandInputs, input_id: str, fallback: float = 0.0) -> float:
    inp = adsk.core.ValueCommandInput.cast(inputs.itemById(input_id))
    return inp.value if inp else fallback


def _float_spinner_value(inputs: adsk.core.CommandInputs, input_id: str, fallback: float = 0.0) -> float:
    inp = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(input_id))
    return inp.value if inp else fallback


def _int_spinner_value(inputs: adsk.core.CommandInputs, input_id: str, fallback: int = 0) -> int:
    inp = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById(input_id))
    return inp.value if inp else fallback


def _set_value_input(inputs: adsk.core.CommandInputs, input_id: str, value: Optional[float]):
    if value is None:
        return
    inp = adsk.core.ValueCommandInput.cast(inputs.itemById(input_id))
    if inp:
        inp.value = float(value)


def _set_float_spinner(inputs: adsk.core.CommandInputs, input_id: str, value: Optional[float]):
    if value is None:
        return
    inp = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(input_id))
    if inp:
        inp.value = float(value)


def _set_int_spinner(inputs: adsk.core.CommandInputs, input_id: str, value: Optional[int]):
    if value is None:
        return
    inp = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById(input_id))
    if inp:
        inp.value = int(value)


def _set_dropdown(inputs: adsk.core.CommandInputs, input_id: str, item_name: Optional[str]):
    if not item_name:
        return
    dd = adsk.core.DropDownCommandInput.cast(inputs.itemById(input_id))
    _select_dropdown_item(dd, item_name)


def _axis_projection(origin: adsk.core.Point3D, axis: adsk.core.Vector3D, point: adsk.core.Point3D) -> float:
    vec = adsk.core.Vector3D.create(point.x - origin.x, point.y - origin.y, point.z - origin.z)
    return vec.dotProduct(axis)


def _point_offset(point: adsk.core.Point3D, direction: adsk.core.Vector3D, distance: float) -> adsk.core.Point3D:
    return adsk.core.Point3D.create(
        point.x + direction.x * distance,
        point.y + direction.y * distance,
        point.z + direction.z * distance,
    )


def _vector_between_points(p1: adsk.core.Point3D, p2: adsk.core.Point3D) -> adsk.core.Vector3D:
    return adsk.core.Vector3D.create(p2.x - p1.x, p2.y - p1.y, p2.z - p1.z)


def _safe_perpendicular(axis: adsk.core.Vector3D) -> adsk.core.Vector3D:
    x_axis = adsk.core.Vector3D.create(1, 0, 0)
    y_axis = adsk.core.Vector3D.create(0, 1, 0)
    perp = axis.crossProduct(x_axis)
    if perp.length < 1e-6:
        perp = axis.crossProduct(y_axis)
    perp.normalize()
    return perp


def _vector_linear_combo(v1: adsk.core.Vector3D, s1: float, v2: adsk.core.Vector3D, s2: float) -> adsk.core.Vector3D:
    return adsk.core.Vector3D.create(
        v1.x * s1 + v2.x * s2,
        v1.y * s1 + v2.y * s2,
        v1.z * s1 + v2.z * s2,
    )


def _unit_direction_from_angle(basis_u: adsk.core.Vector3D, basis_v: adsk.core.Vector3D, angle: float) -> adsk.core.Vector3D:
    direction = _vector_linear_combo(basis_u, math.cos(angle), basis_v, math.sin(angle))
    if direction.length < 1e-9:
        raise ValueError('Could not resolve shaft basis vectors for helical direction.')
    direction.normalize()
    return direction


def _phase_at_fraction(base_phase: float, handed_sign: float, bucket_wrap_rad: float, fraction: float) -> float:
    return base_phase + handed_sign * bucket_wrap_rad * fraction


def _flight_preset_for_count(flights: int) -> str:
    if flights == 1:
        return FLIGHT_PRESET_SINGLE
    if flights == 2:
        return FLIGHT_PRESET_DOUBLE
    if flights == 3:
        return FLIGHT_PRESET_TRIPLE
    if flights == 4:
        return FLIGHT_PRESET_QUAD
    return FLIGHT_PRESET_CUSTOM


def _sync_linked_controls(inputs: adsk.core.CommandInputs):
    turns_spin = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(INPUT_TURNS))
    turns_slider = adsk.core.FloatSliderCommandInput.cast(inputs.itemById(INPUT_TURNS_SLIDER))
    bucket_spin = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(INPUT_BUCKET_WRAP))
    bucket_slider = adsk.core.FloatSliderCommandInput.cast(inputs.itemById(INPUT_BUCKET_WRAP_SLIDER))
    flight_preset = adsk.core.DropDownCommandInput.cast(inputs.itemById(INPUT_FLIGHT_PRESET))
    flights_spin = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById(INPUT_FLIGHTS))

    if turns_spin and turns_slider:
        if turns_spin.value > turns_slider.maximumValue:
            turns_slider.maximumValue = turns_spin.value
        turns_slider.valueOne = turns_spin.value

    if bucket_spin and bucket_slider:
        bucket_slider.valueOne = bucket_spin.value

    if flight_preset and flights_spin:
        _select_dropdown_item(flight_preset, _flight_preset_for_count(flights_spin.value))


def _update_pitch_inputs_visibility(inputs: adsk.core.CommandInputs):
    variable = _dropdown_selected_name(inputs, INPUT_PITCH_MODE, PITCH_CONSTANT) == PITCH_VARIABLE
    length_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_LENGTH))
    pitch_start_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_PITCH_START))
    pitch_end_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_PITCH_END))
    if length_input:
        length_input.isVisible = not variable
    if pitch_start_input:
        pitch_start_input.isVisible = variable
    if pitch_end_input:
        pitch_end_input.isVisible = variable


def _update_tip_thickness_visibility(inputs: adsk.core.CommandInputs):
    tapered = _dropdown_selected_name(inputs, INPUT_THICKNESS_MODE, THICKNESS_CONSTANT) == THICKNESS_TAPERED
    tip_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_TIP_THICKNESS))
    if tip_input:
        tip_input.isVisible = tapered


def _sync_variable_mode_length(inputs: adsk.core.CommandInputs):
    if _dropdown_selected_name(inputs, INPUT_PITCH_MODE, PITCH_CONSTANT) != PITCH_VARIABLE:
        return
    turns = _float_spinner_value(inputs, INPUT_TURNS, 0.0)
    pitch_start = _value_input_value(inputs, INPUT_PITCH_START, 0.0)
    pitch_end = _value_input_value(inputs, INPUT_PITCH_END, 0.0)
    implied_length = turns * 0.5 * (pitch_start + pitch_end)
    if implied_length <= 0:
        return
    length_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_LENGTH))
    if length_input:
        length_input.value = implied_length


def _normalize_preset_name(name: str) -> str:
    return ' '.join(name.strip().split())


def _set_preset_status(inputs: adsk.core.CommandInputs, text: str):
    status = adsk.core.TextBoxCommandInput.cast(inputs.itemById(INPUT_PRESET_STATUS))
    if status:
        status.text = text


def _load_user_presets():
    global _user_presets
    _user_presets = {}

    if not os.path.exists(PRESET_FILE):
        return

    try:
        with open(PRESET_FILE, 'r', encoding='utf-8') as f:
            raw = json.load(f)
    except Exception:
        return

    entries = []
    if isinstance(raw, dict):
        entries = raw.get('presets', [])
    elif isinstance(raw, list):
        entries = raw

    if not isinstance(entries, list):
        return

    for item in entries:
        if not isinstance(item, dict):
            continue
        name = _normalize_preset_name(str(item.get('name', '')))
        data = item.get('data')
        if not name or not isinstance(data, dict):
            continue
        _user_presets[name] = data


def _save_user_presets():
    payload = {
        'version': 1,
        'presets': [{'name': k, 'data': _user_presets[k]} for k in sorted(_user_presets.keys())],
    }
    with open(PRESET_FILE, 'w', encoding='utf-8') as f:
        json.dump(payload, f, indent=2, sort_keys=True)


def _populate_saved_preset_dropdown(inputs: adsk.core.CommandInputs, selected_item_name: Optional[str] = None):
    dd = adsk.core.DropDownCommandInput.cast(inputs.itemById(INPUT_SAVED_PRESET))
    if not dd:
        return

    if selected_item_name is None:
        selected_item_name = PRESET_MANUAL

    dd.listItems.clear()
    dd.listItems.add(PRESET_MANUAL, selected_item_name == PRESET_MANUAL, '')

    for name in sorted(_BUILTIN_PRESETS.keys()):
        label = PRESET_BUILTIN_PREFIX + name
        dd.listItems.add(label, selected_item_name == label, '')

    for name in sorted(_user_presets.keys()):
        label = PRESET_USER_PREFIX + name
        dd.listItems.add(label, selected_item_name == label, '')


def _selected_preset_payload(inputs: adsk.core.CommandInputs) -> Optional[dict]:
    dd = adsk.core.DropDownCommandInput.cast(inputs.itemById(INPUT_SAVED_PRESET))
    if not dd or not dd.selectedItem:
        return None

    label = dd.selectedItem.name
    if label.startswith(PRESET_BUILTIN_PREFIX):
        name = label[len(PRESET_BUILTIN_PREFIX):]
        return _BUILTIN_PRESETS.get(name)
    if label.startswith(PRESET_USER_PREFIX):
        name = label[len(PRESET_USER_PREFIX):]
        return _user_presets.get(name)
    return None


def _current_preset_payload(inputs: adsk.core.CommandInputs) -> dict:
    return {
        'startEnd': _dropdown_selected_name(inputs, INPUT_START_END, START_END_MIN),
        'outerRadius': _value_input_value(inputs, INPUT_OUTER_RADIUS),
        'bladeLength': _value_input_value(inputs, INPUT_LENGTH),
        'turns': _float_spinner_value(inputs, INPUT_TURNS),
        'bladeThickness': _value_input_value(inputs, INPUT_THICKNESS),
        'hubClearance': _value_input_value(inputs, INPUT_CLEARANCE),
        'bucketWrapDeg': _float_spinner_value(inputs, INPUT_BUCKET_WRAP),
        'startAngle': _value_input_value(inputs, INPUT_START_ANGLE),
        'handedness': _dropdown_selected_name(inputs, INPUT_HANDEDNESS, HAND_RIGHT),
        'flightPreset': _dropdown_selected_name(inputs, INPUT_FLIGHT_PRESET, FLIGHT_PRESET_DOUBLE),
        'flights': _int_spinner_value(inputs, INPUT_FLIGHTS, 2),
        'operation': _dropdown_selected_name(inputs, INPUT_OPERATION, OP_NEW_BODY),
        'pitchMode': _dropdown_selected_name(inputs, INPUT_PITCH_MODE, PITCH_CONSTANT),
        'pitchStart': _value_input_value(inputs, INPUT_PITCH_START),
        'pitchEnd': _value_input_value(inputs, INPUT_PITCH_END),
        'thicknessMode': _dropdown_selected_name(inputs, INPUT_THICKNESS_MODE, THICKNESS_CONSTANT),
        'tipThickness': _value_input_value(inputs, INPUT_TIP_THICKNESS),
        'rpm': _float_spinner_value(inputs, INPUT_RPM, 60.0),
    }


def _apply_preset_payload(inputs: adsk.core.CommandInputs, payload: dict):
    global _applying_preset, _outer_radius_user_edited
    if not payload:
        return

    _applying_preset = True
    try:
        _set_dropdown(inputs, INPUT_START_END, payload.get('startEnd'))
        _set_value_input(inputs, INPUT_OUTER_RADIUS, payload.get('outerRadius'))
        _set_value_input(inputs, INPUT_LENGTH, payload.get('bladeLength'))
        _set_float_spinner(inputs, INPUT_TURNS, payload.get('turns'))
        _set_value_input(inputs, INPUT_THICKNESS, payload.get('bladeThickness'))
        _set_value_input(inputs, INPUT_CLEARANCE, payload.get('hubClearance'))
        _set_float_spinner(inputs, INPUT_BUCKET_WRAP, payload.get('bucketWrapDeg'))
        _set_value_input(inputs, INPUT_START_ANGLE, payload.get('startAngle'))
        _set_dropdown(inputs, INPUT_HANDEDNESS, payload.get('handedness'))
        _set_dropdown(inputs, INPUT_FLIGHT_PRESET, payload.get('flightPreset'))
        _set_int_spinner(inputs, INPUT_FLIGHTS, payload.get('flights'))
        _set_dropdown(inputs, INPUT_OPERATION, payload.get('operation'))
        _set_dropdown(inputs, INPUT_PITCH_MODE, payload.get('pitchMode'))
        _set_value_input(inputs, INPUT_PITCH_START, payload.get('pitchStart'))
        _set_value_input(inputs, INPUT_PITCH_END, payload.get('pitchEnd'))
        _set_dropdown(inputs, INPUT_THICKNESS_MODE, payload.get('thicknessMode'))
        _set_value_input(inputs, INPUT_TIP_THICKNESS, payload.get('tipThickness'))
        _set_float_spinner(inputs, INPUT_RPM, payload.get('rpm'))

        _sync_variable_mode_length(inputs)
        _update_pitch_inputs_visibility(inputs)
        _update_tip_thickness_visibility(inputs)
        _sync_linked_controls(inputs)
        _update_derived_text(inputs)

        if payload.get('outerRadius') is not None:
            _outer_radius_user_edited = True
    finally:
        _applying_preset = False


def _save_current_preset(inputs: adsk.core.CommandInputs) -> Optional[str]:
    name_input = adsk.core.StringValueCommandInput.cast(inputs.itemById(INPUT_PRESET_NAME))
    if not name_input:
        return None

    preset_name = _normalize_preset_name(name_input.value)
    if not preset_name:
        _set_preset_status(inputs, 'Preset save failed: enter a preset name.')
        return None

    try:
        _get_validated_parameters(inputs)
    except Exception as ex:
        _set_preset_status(inputs, f'Preset save failed: {str(ex)}')
        return None

    _user_presets[preset_name] = _current_preset_payload(inputs)
    try:
        _save_user_presets()
    except Exception as ex:
        _set_preset_status(inputs, f'Preset save failed: {str(ex)}')
        return None

    _set_preset_status(inputs, f'Saved preset: {preset_name}')
    return preset_name


def _suggested_outer_radius(shaft_radius: float) -> float:
    return shaft_radius + max(1.0, 0.25 * shaft_radius)


def _autofill_outer_radius_from_shaft(inputs: adsk.core.CommandInputs, force: bool = False):
    global _outer_radius_user_edited
    if _outer_radius_user_edited and not force:
        return

    shaft_entity = _selection_entity(inputs, INPUT_SHAFT_FACE)
    if not shaft_entity:
        return

    shaft_face = adsk.fusion.BRepFace.cast(shaft_entity)
    if not shaft_face:
        return

    cylinder = adsk.core.Cylinder.cast(shaft_face.geometry)
    if not cylinder:
        return

    outer = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_OUTER_RADIUS))
    if not outer:
        return

    outer.value = _suggested_outer_radius(cylinder.radius)
    if force:
        _outer_radius_user_edited = False


# ---------------------------------------------------------------------------
# Helix creation helpers
# ---------------------------------------------------------------------------

def _detect_helix_sign_mode(
    temp_mgr: adsk.fusion.TemporaryBRepManager,
    axis_point: adsk.core.Point3D,
    axis_vector: adsk.core.Vector3D,
    start_point: adsk.core.Point3D,
    pitch: float,
):
    global _HELIX_SIGN_MODE

    if temp_mgr.createHelixWire(axis_point, axis_vector, start_point, pitch, -1.0, 0.0):
        _HELIX_SIGN_MODE = 'turns_sign'
        return

    if temp_mgr.createHelixWire(axis_point, axis_vector, start_point, -pitch, 1.0, 0.0):
        _HELIX_SIGN_MODE = 'pitch_sign'
        return

    flipped = axis_vector.copy()
    flipped.scaleBy(-1.0)
    if temp_mgr.createHelixWire(axis_point, flipped, start_point, pitch, 1.0, 0.0):
        _HELIX_SIGN_MODE = 'flip_axis'
        return

    raise RuntimeError('createHelixWire is not working in this Fusion build.')


def _create_helix_wire(
    temp_mgr: adsk.fusion.TemporaryBRepManager,
    axis_point: adsk.core.Point3D,
    axis_vector: adsk.core.Vector3D,
    start_point: adsk.core.Point3D,
    pitch: float,
    turns: float,
    handed_sign: float,
):
    global _HELIX_SIGN_MODE
    if _HELIX_SIGN_MODE is None:
        _detect_helix_sign_mode(temp_mgr, axis_point, axis_vector, start_point, abs(pitch))

    if _HELIX_SIGN_MODE == 'turns_sign':
        return temp_mgr.createHelixWire(axis_point, axis_vector, start_point, pitch, turns * handed_sign, 0.0)

    if _HELIX_SIGN_MODE == 'pitch_sign':
        return temp_mgr.createHelixWire(axis_point, axis_vector, start_point, pitch * handed_sign, turns, 0.0)

    axis = axis_vector.copy()
    if handed_sign < 0:
        axis.scaleBy(-1.0)
    return temp_mgr.createHelixWire(axis_point, axis, start_point, pitch, turns, 0.0)


def _compute_variable_pitch_angles(turns: float, pitch_start: float, pitch_end: float, segments: int):
    angles = []
    z_values = []
    for i in range(segments + 1):
        t = i / segments
        angles.append(2.0 * math.pi * turns * t)
        z = turns * (pitch_start * t + 0.5 * (pitch_end - pitch_start) * t * t)
        z_values.append(z)
    return angles, z_values


def _create_variable_pitch_helix(
    temp_mgr: adsk.fusion.TemporaryBRepManager,
    params: dict,
    radius: float,
    phase_angle: float,
    segments: int = _VAR_PITCH_SEGMENTS,
):
    pitch_start = params['pitchStart']
    pitch_end = params['pitchEnd']
    turns = params['turns']
    signed = params['handedSign']
    axis = params['axisDirection']
    center = params['startCenter']
    basis_u = params['basisU']
    basis_v = params['basisV']

    angles, z_values = _compute_variable_pitch_angles(turns, pitch_start, pitch_end, segments)
    wire_bodies = []
    for i in range(segments):
        seg_turns = (angles[i + 1] - angles[i]) / (2.0 * math.pi)
        seg_z = z_values[i + 1] - z_values[i]
        if seg_turns < 1e-9 or seg_z < 1e-9:
            continue

        seg_pitch = seg_z / seg_turns
        seg_angle = phase_angle + signed * angles[i]
        seg_dir = _unit_direction_from_angle(basis_u, basis_v, seg_angle)

        seg_start_center = _point_offset(center, axis, z_values[i])
        seg_start = _point_offset(seg_start_center, seg_dir, radius)

        wire = _create_helix_wire(
            temp_mgr,
            seg_start_center,
            axis,
            seg_start,
            seg_pitch,
            seg_turns,
            signed,
        )
        if wire and wire.wires.count > 0:
            wire_bodies.append(wire)

    if not wire_bodies:
        raise RuntimeError('Variable-pitch helix: no wire segments were created.')
    return wire_bodies


# ---------------------------------------------------------------------------
# Shaft start-end detection
# ---------------------------------------------------------------------------

def _auto_detect_start_end_face(
    shaft_face: adsk.fusion.BRepFace,
    cylinder: adsk.core.Cylinder,
    prefer_max_end: bool,
):
    axis = cylinder.axis.copy()
    axis.normalize()
    axis_origin = cylinder.origin

    candidates = []
    seen_temp_ids = set()
    for edge in shaft_face.edges:
        for i in range(edge.faces.count):
            face = edge.faces.item(i)
            if face == shaft_face or face.tempId in seen_temp_ids:
                continue

            plane = adsk.core.Plane.cast(face.geometry)
            if not plane:
                continue

            normal = plane.normal.copy()
            normal.normalize()
            if abs(normal.dotProduct(axis)) < 0.995:
                continue

            seen_temp_ids.add(face.tempId)
            candidates.append(face)

    if not candidates:
        raise ValueError(
            'Unable to auto-detect shaft end face. Ensure planar shaft end caps are adjacent to the selected cylinder.'
        )

    selector = max if prefer_max_end else min
    return selector(candidates, key=lambda f: _axis_projection(axis_origin, axis, f.pointOnFace))


# ---------------------------------------------------------------------------
# Surface + solid creation
# ---------------------------------------------------------------------------

def _create_constant_pitch_surface_between(
    params: dict,
    radius_a: float,
    phase_a: float,
    radius_b: float,
    phase_b: float,
):
    temp_mgr = adsk.fusion.TemporaryBRepManager.get()
    pitch = params['pitchStart']

    dir_a = _unit_direction_from_angle(params['basisU'], params['basisV'], phase_a)
    dir_b = _unit_direction_from_angle(params['basisU'], params['basisV'], phase_b)

    start_a = _point_offset(params['startCenter'], dir_a, radius_a)
    start_b = _point_offset(params['startCenter'], dir_b, radius_b)

    wire_a = _create_helix_wire(
        temp_mgr,
        params['startCenter'],
        params['axisDirection'],
        start_a,
        pitch,
        params['turns'],
        params['handedSign'],
    )
    wire_b = _create_helix_wire(
        temp_mgr,
        params['startCenter'],
        params['axisDirection'],
        start_b,
        pitch,
        params['turns'],
        params['handedSign'],
    )

    if not wire_a or wire_a.wires.count < 1:
        raise RuntimeError('Failed to build first constant-pitch guide wire.')
    if not wire_b or wire_b.wires.count < 1:
        raise RuntimeError('Failed to build second constant-pitch guide wire.')

    surface = temp_mgr.createRuledSurface(wire_a.wires.item(0), wire_b.wires.item(0))
    if not surface:
        raise RuntimeError('Failed to create ruled surface between constant-pitch guides.')
    return surface


def _create_variable_pitch_surface_between(
    params: dict,
    radius_a: float,
    phase_a: float,
    radius_b: float,
    phase_b: float,
):
    temp_mgr = adsk.fusion.TemporaryBRepManager.get()
    seg_a = _create_variable_pitch_helix(temp_mgr, params, radius_a, phase_a, _VAR_PITCH_SEGMENTS)
    seg_b = _create_variable_pitch_helix(temp_mgr, params, radius_b, phase_b, _VAR_PITCH_SEGMENTS)

    seg_count = min(len(seg_a), len(seg_b))
    if seg_count == 0:
        raise RuntimeError('Variable-pitch surface: no matched segment pairs.')

    patches = []
    for i in range(seg_count):
        patch = temp_mgr.createRuledSurface(seg_a[i].wires.item(0), seg_b[i].wires.item(0))
        if patch:
            patches.append(patch)

    if not patches:
        raise RuntimeError('Variable-pitch surface: ruled-surface patch creation failed.')

    stitched = patches[0]
    for patch in patches[1:]:
        ok = temp_mgr.booleanOperation(stitched, patch, adsk.fusion.BooleanTypes.UnionBooleanType)
        if not ok:
            raise RuntimeError('Variable-pitch surface: failed to stitch segment patches.')
    return stitched


def _create_surface_between(params: dict, radius_a: float, phase_a: float, radius_b: float, phase_b: float):
    if params['variablePitch']:
        return _create_variable_pitch_surface_between(params, radius_a, phase_a, radius_b, phase_b)
    return _create_constant_pitch_surface_between(params, radius_a, phase_a, radius_b, phase_b)


def _persist_temp_body(component: adsk.fusion.Component, temp_body: adsk.fusion.BRepBody, name: str):
    design = _active_design()
    if design and design.designType == adsk.fusion.DesignTypes.ParametricDesignType:
        base_feature = component.features.baseFeatures.add()
        base_feature.name = name
        base_feature.startEdit()
        component.bRepBodies.add(temp_body, base_feature)
        base_feature.finishEdit()
        if base_feature.bodies.count < 1:
            raise RuntimeError(f'Failed to add temporary body "{name}" to parametric component.')
        return base_feature.bodies.item(base_feature.bodies.count - 1)

    body = component.bRepBodies.add(temp_body)
    if not body:
        raise RuntimeError(f'Failed to add temporary body "{name}" to component.')
    body.name = name
    return body


def _thicken_surface_body(
    component: adsk.fusion.Component,
    surface_body: adsk.fusion.BRepBody,
    thickness: float,
    flight_index: int,
    label: str,
):
    entities = adsk.core.ObjectCollection.create()
    entities.add(surface_body)

    thicken_features = component.features.thickenFeatures
    thicken_input = thicken_features.createInput(
        entities,
        adsk.core.ValueInput.createByReal(thickness),
        True,
        adsk.fusion.FeatureOperations.NewBodyFeatureOperation,
    )
    thicken_input.isChainSelection = False
    thicken_feature = thicken_features.add(thicken_input)

    if not thicken_feature or thicken_feature.bodies.count < 1:
        raise RuntimeError(f'Thicken failed for flight {flight_index + 1} ({label}).')

    thicken_feature.name = f'Arch Flight {flight_index + 1} Thicken ({label})'

    try:
        surface_body.deleteMe()
    except Exception:
        try:
            surface_body.isLightBulbOn = False
        except Exception:
            pass

    blade_body = thicken_feature.bodies.item(0)
    blade_body.name = f'Archimedean Flight {flight_index + 1} ({label})'
    return blade_body


def _join_bodies(
    component: adsk.fusion.Component,
    target_body: adsk.fusion.BRepBody,
    tool_bodies,
    feature_name: str,
):
    if not tool_bodies:
        return target_body

    tools = adsk.core.ObjectCollection.create()
    for body in tool_bodies:
        tools.add(body)

    combine_input = component.features.combineFeatures.createInput(target_body, tools)
    combine_input.operation = adsk.fusion.FeatureOperations.JoinFeatureOperation
    combine_input.isKeepToolBodies = False
    combine_feature = component.features.combineFeatures.add(combine_input)
    if not combine_feature:
        raise RuntimeError('Boolean join failed. Check that blade and shaft geometries overlap.')

    combine_feature.name = feature_name
    if combine_feature.bodies.count > 0:
        return combine_feature.bodies.item(0)
    return target_body


def _create_constant_thickness_flight(component: adsk.fusion.Component, params: dict, phase_angle: float, index: int):
    phase_inner = _phase_at_fraction(phase_angle, params['handedSign'], params['bucketWrap'], 0.0)
    phase_outer = _phase_at_fraction(phase_angle, params['handedSign'], params['bucketWrap'], 1.0)
    surface_temp = _create_surface_between(
        params,
        params['innerRadius'],
        phase_inner,
        params['outerRadius'],
        phase_outer,
    )

    mode_tag = 'VP' if params['variablePitch'] else 'CP'
    surf_name = f'Arch Flight Surface {index + 1} ({mode_tag})'
    surface_body = _persist_temp_body(component, surface_temp, surf_name)
    return _thicken_surface_body(component, surface_body, params['hubThickness'], index, 'full span')


def _create_tapered_thickness_flight(component: adsk.fusion.Component, params: dict, phase_angle: float, index: int):
    inner = params['innerRadius']
    outer = params['outerRadius']
    span = outer - inner
    band_count = _TAPER_THICKNESS_BANDS
    overlap = _TAPER_BAND_OVERLAP

    band_bodies = []
    for i in range(band_count):
        t0 = i / band_count
        t1 = (i + 1) / band_count

        t0e = max(0.0, t0 - overlap)
        t1e = min(1.0, t1 + overlap)

        r0 = inner + span * t0e
        r1 = inner + span * t1e
        phase0 = _phase_at_fraction(phase_angle, params['handedSign'], params['bucketWrap'], t0e)
        phase1 = _phase_at_fraction(phase_angle, params['handedSign'], params['bucketWrap'], t1e)

        band_surface = _create_surface_between(params, r0, phase0, r1, phase1)
        band_surface_body = _persist_temp_body(
            component,
            band_surface,
            f'Arch Flight Surface {index + 1} Band {i + 1}',
        )

        t_mid = 0.5 * (t0 + t1)
        band_thickness = params['hubThickness'] + (params['tipThickness'] - params['hubThickness']) * t_mid
        label = f'band {i + 1}/{band_count}'
        band_bodies.append(_thicken_surface_body(component, band_surface_body, band_thickness, index, label))

    flight_body = band_bodies[0]
    if len(band_bodies) > 1:
        flight_body = _join_bodies(
            component,
            flight_body,
            band_bodies[1:],
            f'Join Flight {index + 1} Taper Bands',
        )
    flight_body.name = f'Archimedean Flight {index + 1}'
    return flight_body


def _create_single_flight(component: adsk.fusion.Component, params: dict, phase_angle: float, index: int):
    if params['taperedThickness']:
        return _create_tapered_thickness_flight(component, params, phase_angle, index)
    return _create_constant_thickness_flight(component, params, phase_angle, index)


def _final_blade_body_name(params: dict) -> str:
    pitch_tag = 'VP' if params['variablePitch'] else 'CP'
    thickness_tag = 'Taper' if params['taperedThickness'] else 'Const'
    return (
        f'Archimedean Blade '
        f'({params["flights"]}F, {pitch_tag}, {thickness_tag}, '
        f'{params["turns"]:.2f} turns)'
    )


def _build_blades(params: dict):
    component = params['component']

    flight_bodies = []
    for i in range(params['flights']):
        phase = params['startAngle'] + (2.0 * math.pi * i / params['flights'])
        flight_bodies.append(_create_single_flight(component, params, phase, i))

    blade_body = flight_bodies[0]
    if len(flight_bodies) > 1:
        blade_body = _join_bodies(
            component,
            blade_body,
            flight_bodies[1:],
            f'Join {params["flights"]} Flights',
        )

    blade_body.name = _final_blade_body_name(params)

    if params['joinToShaft']:
        _join_bodies(
            component,
            params['shaftBody'],
            [blade_body],
            'Join Archimedean Blade To Shaft',
        )


# ---------------------------------------------------------------------------
# Parameter extraction + validation
# ---------------------------------------------------------------------------

def _get_validated_parameters(inputs: adsk.core.CommandInputs, preview_mode: bool = False):
    shaft_entity = _selection_entity(inputs, INPUT_SHAFT_FACE)
    if not shaft_entity:
        raise ValueError('Select the cylindrical shaft face.')

    shaft_face = adsk.fusion.BRepFace.cast(shaft_entity)
    if not shaft_face:
        raise ValueError('Shaft selection must be a face.')
    if shaft_face.assemblyContext:
        raise ValueError('Assembly occurrence selections are not supported. Select source-component geometry.')

    cylinder = adsk.core.Cylinder.cast(shaft_face.geometry)
    if not cylinder:
        raise ValueError('Selected shaft face must be cylindrical.')

    component = shaft_face.body.parentComponent
    prefer_max = _dropdown_selected_name(inputs, INPUT_START_END, START_END_MIN) == START_END_MAX
    start_face = _auto_detect_start_end_face(shaft_face, cylinder, prefer_max)

    axis = cylinder.axis.copy()
    axis.normalize()
    axis_origin = cylinder.origin.copy()

    axis_direction = axis.copy()
    if prefer_max:
        axis_direction.scaleBy(-1.0)

    start_scalar = _axis_projection(axis_origin, axis, start_face.pointOnFace)
    start_center = _point_offset(axis_origin, axis, start_scalar)

    shaft_point = shaft_face.pointOnFace
    shaft_scalar = _axis_projection(axis_origin, axis, shaft_point)
    shaft_axis_point = _point_offset(axis_origin, axis, shaft_scalar)
    basis_u = _vector_between_points(shaft_axis_point, shaft_point)
    if basis_u.length < 1e-6:
        basis_u = _safe_perpendicular(axis_direction)
    else:
        basis_u.normalize()
    basis_v = axis_direction.crossProduct(basis_u)
    if basis_v.length < 1e-6:
        basis_u = _safe_perpendicular(axis_direction)
        basis_v = axis_direction.crossProduct(basis_u)
    basis_v.normalize()

    length = _value_input_value(inputs, INPUT_LENGTH)
    outer_radius = _value_input_value(inputs, INPUT_OUTER_RADIUS)
    hub_thickness = _value_input_value(inputs, INPUT_THICKNESS)
    clearance = _value_input_value(inputs, INPUT_CLEARANCE)
    start_angle = _value_input_value(inputs, INPUT_START_ANGLE)
    turns = _float_spinner_value(inputs, INPUT_TURNS)
    bucket_wrap_deg = _float_spinner_value(inputs, INPUT_BUCKET_WRAP)
    flights = _int_spinner_value(inputs, INPUT_FLIGHTS)
    rpm = _float_spinner_value(inputs, INPUT_RPM, 60.0)

    handedness = _dropdown_selected_name(inputs, INPUT_HANDEDNESS, HAND_RIGHT)
    operation_name = _dropdown_selected_name(inputs, INPUT_OPERATION, OP_NEW_BODY)
    pitch_mode = _dropdown_selected_name(inputs, INPUT_PITCH_MODE, PITCH_CONSTANT)
    thickness_mode = _dropdown_selected_name(inputs, INPUT_THICKNESS_MODE, THICKNESS_CONSTANT)

    # Guard rails and human-readable errors
    if turns <= 0:
        raise ValueError('Turns must be greater than zero.')
    if flights < 1:
        raise ValueError('Flights must be at least 1.')
    if clearance < 0:
        raise ValueError('Hub clearance cannot be negative.')
    if hub_thickness <= 0:
        raise ValueError('Blade thickness must be greater than zero.')
    if bucket_wrap_deg < 0 or bucket_wrap_deg > 120:
        raise ValueError('Bucket wrap must be between 0° and 120°.')
    if rpm < 0:
        raise ValueError('RPM cannot be negative.')

    inner_radius = cylinder.radius + clearance
    if outer_radius <= inner_radius:
        raise ValueError(
            f'Outer radius ({_format_length(outer_radius)}) must exceed shaft radius + clearance ({_format_length(inner_radius)}).'
        )

    tip_thickness = hub_thickness
    tapered_thickness = thickness_mode == THICKNESS_TAPERED
    if tapered_thickness:
        tip_thickness = _value_input_value(inputs, INPUT_TIP_THICKNESS)
        if tip_thickness <= 0:
            raise ValueError('Tip thickness must be greater than zero.')
        if tip_thickness > hub_thickness:
            raise ValueError('Tip thickness cannot exceed hub thickness in tapered mode.')

    radial_span = outer_radius - inner_radius
    if max(hub_thickness, tip_thickness) >= radial_span:
        raise ValueError(
            'Blade thickness must be less than blade radial span '
            f'({_format_length(radial_span)}).'
        )

    variable_pitch = pitch_mode == PITCH_VARIABLE
    if variable_pitch:
        pitch_start = _value_input_value(inputs, INPUT_PITCH_START)
        pitch_end = _value_input_value(inputs, INPUT_PITCH_END)
        if pitch_start <= 0 or pitch_end <= 0:
            raise ValueError('Pitch start and end must be greater than zero in variable mode.')
        length = turns * 0.5 * (pitch_start + pitch_end)
    else:
        if length <= 0:
            raise ValueError('Blade length must be greater than zero.')
        pitch_start = length / turns
        pitch_end = pitch_start

    handed_sign = 1.0 if handedness == HAND_RIGHT else -1.0
    join_to_shaft = operation_name == OP_JOIN

    preview_flights = flights
    if preview_mode:
        preview_flights = min(flights, 2)
        join_to_shaft = False

    return {
        'component': component,
        'shaftBody': shaft_face.body,
        'startCenter': start_center,
        'axisDirection': axis_direction,
        'basisU': basis_u,
        'basisV': basis_v,
        'innerRadius': inner_radius,
        'outerRadius': outer_radius,
        'shaftRadius': cylinder.radius,
        'length': length,
        'turns': turns,
        'hubThickness': hub_thickness,
        'tipThickness': tip_thickness,
        'taperedThickness': tapered_thickness,
        'startAngle': start_angle,
        'bucketWrap': math.radians(bucket_wrap_deg),
        'bucketWrapDeg': bucket_wrap_deg,
        'handedSign': handed_sign,
        'flights': preview_flights,
        'joinToShaft': join_to_shaft,
        'variablePitch': variable_pitch,
        'pitchStart': pitch_start,
        'pitchEnd': pitch_end,
        'rpm': rpm,
    }


# ---------------------------------------------------------------------------
# Derived engineering feedback
# ---------------------------------------------------------------------------

def _estimated_fill_efficiency(params: dict) -> float:
    pitch_avg = 0.5 * (params['pitchStart'] + params['pitchEnd'])
    pitch_ratio = pitch_avg / max(2.0 * params['outerRadius'], 1e-6)
    wrap_factor = _clamp(params['bucketWrapDeg'] / 120.0, 0.0, 1.0)
    span_factor = _clamp((params['outerRadius'] - params['innerRadius']) / max(params['outerRadius'], 1e-6), 0.0, 1.0)
    ratio_term = math.exp(-((pitch_ratio - 0.9) ** 2) / 0.35)
    # Heuristic only, intended for sanity-check ranking.
    return _clamp(0.30 + 0.22 * wrap_factor + 0.28 * ratio_term + 0.15 * span_factor, 0.20, 0.92)


def _update_derived_text(inputs: adsk.core.CommandInputs):
    derived = adsk.core.TextBoxCommandInput.cast(inputs.itemById(INPUT_DERIVED))
    if not derived:
        return

    shaft_entity = _selection_entity(inputs, INPUT_SHAFT_FACE)
    if not shaft_entity:
        derived.text = 'Select shaft face to view derived values.'
        return

    try:
        params = _get_validated_parameters(inputs)
        pitch_start = params['pitchStart']
        pitch_end = params['pitchEnd']
        pitch_avg = 0.5 * (pitch_start + pitch_end)

        annular_area = math.pi * (params['outerRadius'] ** 2 - params['innerRadius'] ** 2)
        vol_per_rev_ml = annular_area * pitch_avg
        flow_ml_min = vol_per_rev_ml * params['rpm']
        est_eff = _estimated_fill_efficiency(params)
        useful_flow_ml_min = flow_ml_min * est_eff

        helix_start = math.degrees(math.atan2(pitch_start, 2.0 * math.pi * params['outerRadius']))
        helix_end = math.degrees(math.atan2(pitch_end, 2.0 * math.pi * params['outerRadius']))
        pitch_ratio_start = pitch_start / max(2.0 * params['outerRadius'], 1e-6)
        pitch_ratio_end = pitch_end / max(2.0 * params['outerRadius'], 1e-6)

        lines = [
            f'Hub radius:        {_format_length(params["innerRadius"])}',
            f'Axial length:      {_format_length(params["length"])}',
            f'Pitch start/end:   {_format_length(pitch_start)} / {_format_length(pitch_end)}',
            f'Pitch ratio:       {pitch_ratio_start:.2f} -> {pitch_ratio_end:.2f}',
            f'Helix angle:       {helix_start:.1f}° -> {helix_end:.1f}° (outer edge)',
            f'Theoretical flow:  {flow_ml_min / 1000.0:.2f} L/min @ {params["rpm"]:.0f} RPM',
            f'Est. useful flow:  {useful_flow_ml_min / 1000.0:.2f} L/min (eff. {est_eff * 100.0:.0f}%, heuristic)',
        ]
        derived.text = '\n'.join(lines)
    except Exception as ex:
        derived.text = str(ex)


# ---------------------------------------------------------------------------
# Event handlers
# ---------------------------------------------------------------------------

class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def notify(self, args: adsk.core.CommandEventArgs):
        try:
            params = _get_validated_parameters(args.firingEvent.sender.commandInputs)
            _build_blades(params)
        except Exception:
            if UI:
                UI.messageBox('Failed to create Archimedean blade:\n{}'.format(traceback.format_exc()))


class CommandPreviewHandler(adsk.core.CommandEventHandler):
    def notify(self, args: adsk.core.CommandEventArgs):
        try:
            params = _get_validated_parameters(args.firingEvent.sender.commandInputs, preview_mode=True)
            _build_blades(params)
            args.isValidResult = True
        except Exception:
            args.isValidResult = False


class InputChangedHandler(adsk.core.InputChangedEventHandler):
    def notify(self, args: adsk.core.InputChangedEventArgs):
        global _outer_radius_user_edited
        try:
            changed = args.input
            inputs = args.inputs
            if not changed:
                return

            turns_spin = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(INPUT_TURNS))
            turns_slider = adsk.core.FloatSliderCommandInput.cast(inputs.itemById(INPUT_TURNS_SLIDER))
            bucket_spin = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(INPUT_BUCKET_WRAP))
            bucket_slider = adsk.core.FloatSliderCommandInput.cast(inputs.itemById(INPUT_BUCKET_WRAP_SLIDER))
            flight_preset = adsk.core.DropDownCommandInput.cast(inputs.itemById(INPUT_FLIGHT_PRESET))
            flights_spin = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById(INPUT_FLIGHTS))

            if turns_spin and turns_slider:
                if changed.id == INPUT_TURNS_SLIDER:
                    turns_spin.value = turns_slider.valueOne
                elif changed.id == INPUT_TURNS:
                    if turns_spin.value > turns_slider.maximumValue:
                        turns_slider.maximumValue = turns_spin.value
                    turns_slider.valueOne = turns_spin.value

            if bucket_spin and bucket_slider:
                if changed.id == INPUT_BUCKET_WRAP_SLIDER:
                    bucket_spin.value = bucket_slider.valueOne
                elif changed.id == INPUT_BUCKET_WRAP:
                    bucket_slider.valueOne = bucket_spin.value

            if flight_preset and flights_spin:
                if changed.id == INPUT_FLIGHT_PRESET and flight_preset.selectedItem:
                    name = flight_preset.selectedItem.name
                    if name == FLIGHT_PRESET_SINGLE:
                        flights_spin.value = 1
                    elif name == FLIGHT_PRESET_DOUBLE:
                        flights_spin.value = 2
                    elif name == FLIGHT_PRESET_TRIPLE:
                        flights_spin.value = 3
                    elif name == FLIGHT_PRESET_QUAD:
                        flights_spin.value = 4
                elif changed.id == INPUT_FLIGHTS:
                    _select_dropdown_item(flight_preset, _flight_preset_for_count(flights_spin.value))

            if changed.id == INPUT_PITCH_MODE:
                _update_pitch_inputs_visibility(inputs)
            if changed.id in (INPUT_PITCH_MODE, INPUT_PITCH_START, INPUT_PITCH_END, INPUT_TURNS, INPUT_TURNS_SLIDER):
                _sync_variable_mode_length(inputs)

            if changed.id == INPUT_THICKNESS_MODE:
                _update_tip_thickness_visibility(inputs)

            if changed.id == INPUT_SHAFT_FACE:
                if not _applying_preset:
                    _outer_radius_user_edited = False
                _autofill_outer_radius_from_shaft(inputs, force=True)
            elif changed.id == INPUT_OUTER_RADIUS and not _applying_preset:
                _outer_radius_user_edited = True

            if changed.id == INPUT_SAVED_PRESET and not _applying_preset:
                payload = _selected_preset_payload(inputs)
                if payload:
                    _apply_preset_payload(inputs, payload)
                    selected = _dropdown_selected_name(inputs, INPUT_SAVED_PRESET, PRESET_MANUAL)
                    _set_preset_status(inputs, f'Applied preset: {selected}')
                elif _dropdown_selected_name(inputs, INPUT_SAVED_PRESET, PRESET_MANUAL) == PRESET_MANUAL:
                    _set_preset_status(inputs, 'Manual mode.')

            if changed.id == INPUT_SAVE_PRESET:
                save_btn = adsk.core.BoolValueCommandInput.cast(inputs.itemById(INPUT_SAVE_PRESET))
                if save_btn and save_btn.value:
                    saved_name = _save_current_preset(inputs)
                    if saved_name:
                        selected = PRESET_USER_PREFIX + saved_name
                        _populate_saved_preset_dropdown(inputs, selected)
                    save_btn.value = False

            _update_derived_text(inputs)
        except Exception:
            derived = adsk.core.TextBoxCommandInput.cast(args.inputs.itemById(INPUT_DERIVED))
            if derived:
                derived.text = 'Input error:\n' + traceback.format_exc(limit=3)


class ValidateInputsHandler(adsk.core.ValidateInputsEventHandler):
    def notify(self, args: adsk.core.ValidateInputsEventArgs):
        try:
            _get_validated_parameters(args.inputs)
            args.areInputsValid = True
        except Exception:
            args.areInputsValid = False


class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, args: adsk.core.CommandCreatedEventArgs):
        global _outer_radius_user_edited
        try:
            command = args.command
            command.isExecutedWhenPreEmpted = False
            inputs = command.commandInputs
            units = _default_length_units()

            _outer_radius_user_edited = False

            # Shaft selection
            shaft_input = inputs.addSelectionInput(
                INPUT_SHAFT_FACE,
                'Shaft Cylindrical Face',
                'Select the shaft cylindrical face.',
            )
            shaft_input.addSelectionFilter('Faces')
            shaft_input.setSelectionLimits(1, 1)

            # Start end
            start_end = inputs.addDropDownCommandInput(
                INPUT_START_END,
                'Start End',
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            start_end.listItems.add(START_END_MIN, True, '')
            start_end.listItems.add(START_END_MAX, False, '')

            # Core geometry
            inputs.addValueInput(INPUT_OUTER_RADIUS, 'Outer Radius', units, adsk.core.ValueInput.createByString('80 mm'))
            inputs.addValueInput(INPUT_CLEARANCE, 'Hub Clearance', units, adsk.core.ValueInput.createByString('2 mm'))

            pitch_mode = inputs.addDropDownCommandInput(
                INPUT_PITCH_MODE,
                'Pitch Mode',
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            pitch_mode.listItems.add(PITCH_CONSTANT, True, '')
            pitch_mode.listItems.add(PITCH_VARIABLE, False, '')

            inputs.addValueInput(INPUT_LENGTH, 'Blade Length', units, adsk.core.ValueInput.createByString('600 mm'))
            inputs.addValueInput(INPUT_PITCH_START, 'Pitch Start', units, adsk.core.ValueInput.createByString('170 mm'))
            inputs.addValueInput(INPUT_PITCH_END, 'Pitch End', units, adsk.core.ValueInput.createByString('120 mm'))

            inputs.addFloatSpinnerCommandInput(INPUT_TURNS, 'Turns', '', 0.1, 200.0, 0.1, 3.5)
            turns_slider = inputs.addFloatSliderCommandInput(
                INPUT_TURNS_SLIDER,
                'Turns (drag)',
                '',
                0.1,
                20.0,
                False,
            )
            turns_slider.spinStep = 0.1
            turns_slider.valueOne = 3.5

            inputs.addFloatSpinnerCommandInput(INPUT_BUCKET_WRAP, 'Bucket Wrap (deg)', '', 0.0, 120.0, 1.0, 35.0)
            bucket_slider = inputs.addFloatSliderCommandInput(
                INPUT_BUCKET_WRAP_SLIDER,
                'Bucket Wrap (drag)',
                '',
                0.0,
                120.0,
                False,
            )
            bucket_slider.spinStep = 1.0
            bucket_slider.valueOne = 35.0

            # Thickness controls
            inputs.addValueInput(INPUT_THICKNESS, 'Hub Thickness', units, adsk.core.ValueInput.createByString('3 mm'))
            tm = inputs.addDropDownCommandInput(
                INPUT_THICKNESS_MODE,
                'Thickness Mode',
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            tm.listItems.add(THICKNESS_CONSTANT, True, '')
            tm.listItems.add(THICKNESS_TAPERED, False, '')
            inputs.addValueInput(INPUT_TIP_THICKNESS, 'Tip Thickness', units, adsk.core.ValueInput.createByString('2 mm'))

            # Orientation + pattern
            inputs.addValueInput(INPUT_START_ANGLE, 'Start Angle', 'deg', adsk.core.ValueInput.createByString('0 deg'))

            handed = inputs.addDropDownCommandInput(
                INPUT_HANDEDNESS,
                'Handedness',
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            handed.listItems.add(HAND_RIGHT, True, '')
            handed.listItems.add(HAND_LEFT, False, '')

            flight_preset = inputs.addDropDownCommandInput(
                INPUT_FLIGHT_PRESET,
                'Blade Preset',
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            flight_preset.listItems.add(FLIGHT_PRESET_SINGLE, False, '')
            flight_preset.listItems.add(FLIGHT_PRESET_DOUBLE, True, '')
            flight_preset.listItems.add(FLIGHT_PRESET_TRIPLE, False, '')
            flight_preset.listItems.add(FLIGHT_PRESET_QUAD, False, '')
            flight_preset.listItems.add(FLIGHT_PRESET_CUSTOM, False, '')

            inputs.addIntegerSpinnerCommandInput(INPUT_FLIGHTS, 'Flights', 1, 6, 1, 2)

            operation = inputs.addDropDownCommandInput(
                INPUT_OPERATION,
                'Operation',
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            operation.listItems.add(OP_NEW_BODY, True, '')
            operation.listItems.add(OP_JOIN, False, '')

            inputs.addFloatSpinnerCommandInput(INPUT_RPM, 'RPM', '', 0.0, 5000.0, 1.0, 60.0)

            # Presets
            inputs.addDropDownCommandInput(
                INPUT_SAVED_PRESET,
                'Saved Preset',
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            inputs.addStringValueInput(INPUT_PRESET_NAME, 'Preset Name', '')
            inputs.addBoolValueInput(INPUT_SAVE_PRESET, 'Save Preset', True, '', False)
            inputs.addTextBoxCommandInput(INPUT_PRESET_STATUS, 'Preset Status', 'Preset file: presets.json', 2, True)

            # Derived output
            inputs.addTextBoxCommandInput(
                INPUT_DERIVED,
                'Derived',
                'Select shaft face to view derived values.',
                7,
                True,
            )

            _load_user_presets()
            _populate_saved_preset_dropdown(inputs)
            _update_pitch_inputs_visibility(inputs)
            _update_tip_thickness_visibility(inputs)
            _sync_variable_mode_length(inputs)
            _sync_linked_controls(inputs)
            _update_derived_text(inputs)

            cmd_handlers: List[object] = []

            on_execute = CommandExecuteHandler()
            command.execute.add(on_execute)
            cmd_handlers.append(on_execute)

            on_preview = CommandPreviewHandler()
            command.executePreview.add(on_preview)
            cmd_handlers.append(on_preview)

            on_changed = InputChangedHandler()
            command.inputChanged.add(on_changed)
            cmd_handlers.append(on_changed)

            on_validate = ValidateInputsHandler()
            command.validateInputs.add(on_validate)
            cmd_handlers.append(on_validate)

            _handler_registry[command.parentCommandDefinition.id] = cmd_handlers
        except Exception:
            if UI:
                UI.messageBox('Command creation failed:\n{}'.format(traceback.format_exc()))


# ---------------------------------------------------------------------------
# Add-in lifecycle
# ---------------------------------------------------------------------------

def run(context):
    try:
        if not UI:
            return

        cmd_def = UI.commandDefinitions.itemById(CMD_ID)
        if cmd_def:
            return

        cmd_def = UI.commandDefinitions.addButtonDefinition(
            CMD_ID,
            CMD_NAME,
            CMD_DESCRIPTION,
            RESOURCE_FOLDER,
        )

        on_created = CommandCreatedHandler()
        cmd_def.commandCreated.add(on_created)
        _handler_registry[CMD_ID + '_created'] = [on_created]

        workspace = UI.workspaces.itemById(WORKSPACE_ID)
        panel = workspace.toolbarPanels.itemById(PANEL_ID)
        control = panel.controls.itemById(CMD_ID)
        if not control:
            control = panel.controls.addCommand(cmd_def)
            control.isPromoted = True
            control.isPromotedByDefault = True
    except Exception:
        if UI:
            UI.messageBox('Add-in start failed:\n{}'.format(traceback.format_exc()))


def stop(context):
    global _HELIX_SIGN_MODE, _outer_radius_user_edited, _applying_preset
    try:
        if not UI:
            return

        workspace = UI.workspaces.itemById(WORKSPACE_ID)
        panel = workspace.toolbarPanels.itemById(PANEL_ID)
        control = panel.controls.itemById(CMD_ID)
        if control:
            control.deleteMe()

        cmd_def = UI.commandDefinitions.itemById(CMD_ID)
        if cmd_def:
            cmd_def.deleteMe()

        _handler_registry.clear()
        _HELIX_SIGN_MODE = None
        _outer_radius_user_edited = False
        _applying_preset = False
    except Exception:
        if UI:
            UI.messageBox('Add-in stop failed:\n{}'.format(traceback.format_exc()))
