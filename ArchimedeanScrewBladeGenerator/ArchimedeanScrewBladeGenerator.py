import adsk.core
import adsk.fusion
import math
import os
import traceback

APP = adsk.core.Application.get()
UI = APP.userInterface if APP else None

ADDIN_NAME = 'Archimedean Screw Blade Generator'
CMD_ID = 'archimedean_screw_blade_generator_cmd'
CMD_NAME = 'Archimedean Screw Blade'
CMD_DESCRIPTION = 'Create configurable Archimedean screw blades around an existing shaft.'
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidCreatePanel'
RESOURCE_FOLDER = os.path.join(os.path.dirname(__file__), 'resources')

INPUT_SHAFT_FACE = 'shaftFace'
INPUT_OUTER_RADIUS = 'outerRadius'
INPUT_LENGTH = 'bladeLength'
INPUT_TURNS = 'turns'
INPUT_TURNS_SLIDER = 'turnsSlider'
INPUT_THICKNESS = 'bladeThickness'
INPUT_CLEARANCE = 'hubClearance'
INPUT_START_ANGLE = 'startAngle'
INPUT_HANDEDNESS = 'handedness'
INPUT_FLIGHTS = 'flights'
INPUT_SEGMENTS = 'segmentsPerTurn'
INPUT_OPERATION = 'operation'
INPUT_START_END = 'startEnd'
INPUT_DERIVED = 'derivedInfo'

HAND_RIGHT = 'Right-handed'
HAND_LEFT = 'Left-handed'
OP_NEW_BODY = 'New Blade Body'
OP_JOIN = 'Join Blade To Shaft'
START_END_MIN = 'End 1 (auto)'
START_END_MAX = 'End 2 (auto)'

_handlers = []


def _active_design() -> adsk.fusion.Design:
    return adsk.fusion.Design.cast(APP.activeProduct)


def _default_length_units() -> str:
    design = _active_design()
    if not design:
        return 'cm'
    return design.unitsManager.defaultLengthUnits


def _format_length(value: float) -> str:
    design = _active_design()
    if not design:
        return f'{value:.4f} cm'
    units_mgr = design.unitsManager
    return units_mgr.formatInternalValue(value, units_mgr.defaultLengthUnits, True)


def _selection_entity(inputs: adsk.core.CommandInputs, input_id: str):
    selection_input = adsk.core.SelectionCommandInput.cast(inputs.itemById(input_id))
    if not selection_input or selection_input.selectionCount < 1:
        return None
    return selection_input.selection(0).entity


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


def _vector_linear_combo(
    v1: adsk.core.Vector3D,
    s1: float,
    v2: adsk.core.Vector3D,
    s2: float,
) -> adsk.core.Vector3D:
    return adsk.core.Vector3D.create(
        v1.x * s1 + v2.x * s2,
        v1.y * s1 + v2.y * s2,
        v1.z * s1 + v2.z * s2,
    )


def _auto_detect_start_end_face(
    shaft_face: adsk.fusion.BRepFace,
    cylinder: adsk.core.Cylinder,
    prefer_max_end: bool,
):
    axis = cylinder.axis.copy()
    axis.normalize()
    axis_origin = cylinder.origin

    candidate_faces = []
    seen_temp_ids = set()
    edges = shaft_face.edges
    for edge in edges:
        edge_faces = edge.faces
        for i in range(edge_faces.count):
            face = edge_faces.item(i)
            if face == shaft_face:
                continue
            if face.tempId in seen_temp_ids:
                continue

            plane = adsk.core.Plane.cast(face.geometry)
            if not plane:
                continue

            normal = plane.normal.copy()
            normal.normalize()
            if abs(normal.dotProduct(axis)) < 0.995:
                continue

            seen_temp_ids.add(face.tempId)
            candidate_faces.append(face)

    if not candidate_faces:
        raise ValueError(
            'Unable to auto-detect shaft end face. Make sure the shaft has planar end caps connected to the selected cylindrical face.'
        )

    selector = max if prefer_max_end else min
    return selector(
        candidate_faces,
        key=lambda f: _axis_projection(axis_origin, axis, f.pointOnFace),
    )


def _get_validated_parameters(inputs: adsk.core.CommandInputs, preview_mode: bool = False):
    shaft_entity = _selection_entity(inputs, INPUT_SHAFT_FACE)
    if not shaft_entity:
        raise ValueError('Select the cylindrical shaft face.')

    shaft_face = adsk.fusion.BRepFace.cast(shaft_entity)
    if not shaft_face:
        raise ValueError('Shaft selection must be a face.')

    if shaft_face.assemblyContext:
        raise ValueError('Assembly occurrences are not supported. Select entities in their source component.')

    cylinder = adsk.core.Cylinder.cast(shaft_face.geometry)
    if not cylinder:
        raise ValueError('Selected shaft face must be cylindrical.')
    component = shaft_face.body.parentComponent
    start_end_input = adsk.core.DropDownCommandInput.cast(inputs.itemById(INPUT_START_END))
    start_end_name = start_end_input.selectedItem.name if start_end_input and start_end_input.selectedItem else START_END_MIN
    prefer_max_end = start_end_name == START_END_MAX
    start_planar_entity = _auto_detect_start_end_face(shaft_face, cylinder, prefer_max_end)
    start_plane = adsk.core.Plane.cast(start_planar_entity.geometry)
    if not start_plane:
        raise ValueError('Auto-detected shaft end is not planar.')

    axis = cylinder.axis.copy()
    axis.normalize()
    axis_origin = cylinder.origin.copy()

    axis_direction = axis.copy()
    if prefer_max_end:
        axis_direction.scaleBy(-1.0)

    start_normal = start_plane.normal.copy()
    start_normal.normalize()
    offset_sign = 1.0 if start_normal.dotProduct(axis_direction) >= 0.0 else -1.0

    start_scalar = _axis_projection(axis_origin, axis, start_planar_entity.pointOnFace)
    start_center = _point_offset(axis_origin, axis, start_scalar)

    shaft_point = shaft_face.pointOnFace
    shaft_scalar = _axis_projection(axis_origin, axis, shaft_point)
    shaft_axis_point = _point_offset(axis_origin, axis, shaft_scalar)
    basis_u = _vector_between_points(shaft_axis_point, shaft_point)
    if basis_u.length < 1e-6:
        basis_u = _safe_perpendicular(axis)
    else:
        basis_u.normalize()
    basis_v = axis_direction.crossProduct(basis_u)
    if basis_v.length < 1e-6:
        basis_u = _safe_perpendicular(axis_direction)
        basis_v = axis_direction.crossProduct(basis_u)
    basis_v.normalize()

    length_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_LENGTH))
    outer_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_OUTER_RADIUS))
    thickness_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_THICKNESS))
    clearance_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_CLEARANCE))
    start_angle_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_START_ANGLE))
    turns_input = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(INPUT_TURNS))
    segments_input = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById(INPUT_SEGMENTS))
    flights_input = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById(INPUT_FLIGHTS))
    hand_input = adsk.core.DropDownCommandInput.cast(inputs.itemById(INPUT_HANDEDNESS))
    operation_input = adsk.core.DropDownCommandInput.cast(inputs.itemById(INPUT_OPERATION))

    length = length_input.value
    outer_radius = outer_input.value
    thickness = thickness_input.value
    clearance = clearance_input.value
    start_angle = start_angle_input.value
    turns = turns_input.value
    segments_per_turn = segments_input.value
    flights = flights_input.value

    if length <= 0:
        raise ValueError('Blade length must be greater than zero.')
    if turns <= 0:
        raise ValueError('Turns must be greater than zero.')
    if thickness <= 0:
        raise ValueError('Blade thickness must be greater than zero.')
    if segments_per_turn < 8:
        raise ValueError('Use at least 8 segments per turn.')
    if flights < 1:
        raise ValueError('Flights must be at least 1.')

    inner_radius = cylinder.radius + clearance
    if outer_radius <= inner_radius:
        raise ValueError('Outer radius must be greater than shaft radius + clearance.')

    handedness = hand_input.selectedItem.name if hand_input and hand_input.selectedItem else HAND_RIGHT
    handed_sign = 1.0 if handedness == HAND_RIGHT else -1.0

    operation_name = operation_input.selectedItem.name if operation_input and operation_input.selectedItem else OP_NEW_BODY
    join_to_shaft = operation_name == OP_JOIN

    station_count = max(2, int(math.ceil(turns * segments_per_turn)) + 1)
    preview_segment_cap = 3
    preview_flights = flights
    if preview_mode:
        preview_station_count = max(2, int(math.ceil(turns * min(segments_per_turn, preview_segment_cap))) + 1)
        station_count = min(station_count, preview_station_count)
        join_to_shaft = False
        preview_flights = min(flights, 2)

    return {
        'component': component,
        'shaftBody': shaft_face.body,
        'startPlanarEntity': start_planar_entity,
        'startCenter': start_center,
        'axisDirection': axis_direction,
        'basisU': basis_u,
        'basisV': basis_v,
        'offsetSign': offset_sign,
        'innerRadius': inner_radius,
        'outerRadius': outer_radius,
        'length': length,
        'turns': turns,
        'thickness': thickness,
        'startAngle': start_angle,
        'handedSign': handed_sign,
        'flights': preview_flights,
        'segmentsPerTurn': segments_per_turn,
        'stationCount': station_count,
        'joinToShaft': join_to_shaft,
        'isPreview': preview_mode,
    }


def _add_profile_section(
    component: adsk.fusion.Component,
    base_entity,
    plane_offset_distance: float,
    axis_distance: float,
    start_center: adsk.core.Point3D,
    axis_direction: adsk.core.Vector3D,
    basis_u: adsk.core.Vector3D,
    basis_v: adsk.core.Vector3D,
    inner_radius: float,
    outer_radius: float,
    thickness: float,
    angle: float,
):
    planes = component.constructionPlanes
    plane_input = planes.createInput()
    plane_input.setByOffset(base_entity, adsk.core.ValueInput.createByReal(plane_offset_distance))
    section_plane = planes.add(plane_input)
    section_plane.isLightBulbOn = False

    sketch = component.sketches.add(section_plane)
    sketch.isLightBulbOn = False

    cos_a = math.cos(angle)
    sin_a = math.sin(angle)

    center = _point_offset(start_center, axis_direction, axis_distance)
    radial = _vector_linear_combo(basis_u, cos_a, basis_v, sin_a)
    radial.normalize()
    tangential = axis_direction.crossProduct(radial)
    tangential.normalize()

    half_t = thickness / 2.0

    inner_center = _point_offset(center, radial, inner_radius)
    outer_center = _point_offset(center, radial, outer_radius)

    p1 = _point_offset(inner_center, tangential, half_t)
    p2 = _point_offset(outer_center, tangential, half_t)
    p3 = _point_offset(outer_center, tangential, -half_t)
    p4 = _point_offset(inner_center, tangential, -half_t)

    # Sketch lines require sketch-space coordinates; using model-space points directly
    # distorts profiles when the sketch plane is not aligned to world axes.
    p1s = sketch.modelToSketchSpace(p1)
    p2s = sketch.modelToSketchSpace(p2)
    p3s = sketch.modelToSketchSpace(p3)
    p4s = sketch.modelToSketchSpace(p4)

    lines = sketch.sketchCurves.sketchLines
    lines.addByTwoPoints(p1s, p2s)
    lines.addByTwoPoints(p2s, p3s)
    lines.addByTwoPoints(p3s, p4s)
    lines.addByTwoPoints(p4s, p1s)

    if sketch.profiles.count < 1:
        raise RuntimeError('Failed to create section profile. Try increasing blade thickness.')

    return sketch.profiles.item(0)


def _create_single_flight(component: adsk.fusion.Component, params: dict, phase_angle: float):
    loft_features = component.features.loftFeatures
    loft_input = loft_features.createInput(adsk.fusion.FeatureOperations.NewBodyFeatureOperation)
    loft_input.isSolid = True

    sections = loft_input.loftSections

    section_count = params['stationCount']
    for i in range(section_count):
        t = i / float(section_count - 1)
        axis_distance = params['length'] * t
        plane_offset_distance = axis_distance * params['offsetSign']
        angle = phase_angle + params['handedSign'] * (2.0 * math.pi * params['turns'] * t)

        profile = _add_profile_section(
            params['component'],
            params['startPlanarEntity'],
            plane_offset_distance,
            axis_distance,
            params['startCenter'],
            params['axisDirection'],
            params['basisU'],
            params['basisV'],
            params['innerRadius'],
            params['outerRadius'],
            params['thickness'],
            angle,
        )
        sections.add(profile)

    loft_feature = loft_features.add(loft_input)
    if loft_feature.bodies.count < 1:
        raise RuntimeError('Loft creation failed for this flight.')

    return loft_feature.bodies.item(0)


def _join_bodies(component: adsk.fusion.Component, target_body: adsk.fusion.BRepBody, tool_bodies):
    if not tool_bodies:
        return target_body

    tools = adsk.core.ObjectCollection.create()
    for body in tool_bodies:
        tools.add(body)

    combine_features = component.features.combineFeatures
    combine_input = combine_features.createInput(target_body, tools)
    combine_input.operation = adsk.fusion.FeatureOperations.JoinFeatureOperation
    combine_input.isKeepToolBodies = False
    combine_features.add(combine_input)
    return target_body


def _build_blades(params: dict):
    component = params['component']

    flight_bodies = []
    for i in range(params['flights']):
        phase = params['startAngle'] + (2.0 * math.pi * i / params['flights'])
        body = _create_single_flight(component, params, phase)
        flight_bodies.append(body)

    blade_body = flight_bodies[0]
    if len(flight_bodies) > 1:
        blade_body = _join_bodies(component, blade_body, flight_bodies[1:])

    if params['joinToShaft']:
        _join_bodies(component, params['shaftBody'], [blade_body])


def _update_derived_text(inputs: adsk.core.CommandInputs):
    text_input = adsk.core.TextBoxCommandInput.cast(inputs.itemById(INPUT_DERIVED))
    if not text_input:
        return

    shaft_entity = _selection_entity(inputs, INPUT_SHAFT_FACE)
    if not shaft_entity:
        text_input.text = 'Select shaft face to view derived values.'
        return

    try:
        params = _get_validated_parameters(inputs)
        pitch = params['length'] / params['turns']
        text_input.text = (
            f"Hub radius: {_format_length(params['innerRadius'])}\n"
            f"Pitch: {_format_length(pitch)}\n"
            f"Sections per flight: {params['stationCount']}"
        )
    except Exception as ex:
        text_input.text = str(ex)


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def notify(self, args: adsk.core.CommandEventArgs):
        try:
            command = args.firingEvent.sender
            inputs = command.commandInputs
            params = _get_validated_parameters(inputs)
            _build_blades(params)
        except Exception:
            if UI:
                UI.messageBox('Failed to create Archimedean blade:\n{}'.format(traceback.format_exc()))


class CommandPreviewHandler(adsk.core.CommandEventHandler):
    def notify(self, args: adsk.core.CommandEventArgs):
        try:
            command = args.firingEvent.sender
            inputs = command.commandInputs
            params = _get_validated_parameters(inputs, preview_mode=True)
            _build_blades(params)
            args.isValidResult = True
        except Exception:
            # Preview should fail quietly while the user is still selecting/editing inputs.
            args.isValidResult = False


class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, args: adsk.core.CommandCreatedEventArgs):
        try:
            command = args.command
            command.isExecutedWhenPreEmpted = False
            inputs = command.commandInputs

            units = _default_length_units()

            shaft_input = inputs.addSelectionInput(
                INPUT_SHAFT_FACE,
                'Shaft Cylindrical Face',
                'Select the shaft cylindrical face',
            )
            shaft_input.addSelectionFilter('Faces')
            shaft_input.setSelectionLimits(1, 1)

            start_end_input = inputs.addDropDownCommandInput(
                INPUT_START_END,
                'Start End',
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            start_end_input.listItems.add(START_END_MIN, True, '')
            start_end_input.listItems.add(START_END_MAX, False, '')

            inputs.addValueInput(
                INPUT_OUTER_RADIUS,
                'Outer Radius',
                units,
                adsk.core.ValueInput.createByString(f'5 {units}'),
            )
            inputs.addValueInput(
                INPUT_LENGTH,
                'Blade Length',
                units,
                adsk.core.ValueInput.createByString(f'30 {units}'),
            )

            # Empty unit string means this is a unitless spinner (number of turns).
            inputs.addFloatSpinnerCommandInput(INPUT_TURNS, 'Turns', '', 0.1, 200.0, 0.1, 3.0)
            turns_slider = inputs.addFloatSliderCommandInput(
                INPUT_TURNS_SLIDER,
                'Turns (drag)',
                '',
                0.1,
                20.0,
                False,
            )
            turns_slider.spinStep = 0.1
            turns_slider.valueOne = 3.0

            inputs.addValueInput(
                INPUT_THICKNESS,
                'Blade Thickness',
                units,
                adsk.core.ValueInput.createByString(f'0.4 {units}'),
            )
            inputs.addValueInput(
                INPUT_CLEARANCE,
                'Hub Clearance',
                units,
                adsk.core.ValueInput.createByString(f'0 {units}'),
            )
            inputs.addValueInput(
                INPUT_START_ANGLE,
                'Start Angle',
                'deg',
                adsk.core.ValueInput.createByString('0 deg'),
            )

            hand_input = inputs.addDropDownCommandInput(
                INPUT_HANDEDNESS,
                'Handedness',
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            hand_input.listItems.add(HAND_RIGHT, True, '')
            hand_input.listItems.add(HAND_LEFT, False, '')

            inputs.addIntegerSpinnerCommandInput(INPUT_FLIGHTS, 'Flights', 1, 6, 1, 1)
            inputs.addIntegerSpinnerCommandInput(INPUT_SEGMENTS, 'Segments / Turn', 8, 300, 1, 24)

            operation_input = inputs.addDropDownCommandInput(
                INPUT_OPERATION,
                'Operation',
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            operation_input.listItems.add(OP_NEW_BODY, True, '')
            operation_input.listItems.add(OP_JOIN, False, '')

            inputs.addTextBoxCommandInput(
                INPUT_DERIVED,
                'Derived',
                'Select shaft face to view derived values.',
                3,
                True,
            )

            on_execute = CommandExecuteHandler()
            command.execute.add(on_execute)
            _handlers.append(on_execute)

            on_preview = CommandPreviewHandler()
            command.executePreview.add(on_preview)
            _handlers.append(on_preview)

            on_input_changed = InputChangedHandler()
            command.inputChanged.add(on_input_changed)
            _handlers.append(on_input_changed)

            on_validate = ValidateInputsHandler()
            command.validateInputs.add(on_validate)
            _handlers.append(on_validate)

            _update_derived_text(inputs)

        except Exception:
            if UI:
                UI.messageBox('Command creation failed:\n{}'.format(traceback.format_exc()))


class InputChangedHandler(adsk.core.InputChangedEventHandler):
    def notify(self, args: adsk.core.InputChangedEventArgs):
        try:
            changed = args.input
            turns_input = adsk.core.FloatSpinnerCommandInput.cast(args.inputs.itemById(INPUT_TURNS))
            turns_slider = adsk.core.FloatSliderCommandInput.cast(args.inputs.itemById(INPUT_TURNS_SLIDER))

            if changed and turns_input and turns_slider:
                if changed.id == INPUT_TURNS_SLIDER:
                    turns_input.value = turns_slider.valueOne
                elif changed.id == INPUT_TURNS:
                    if turns_input.value > turns_slider.maximumValue:
                        turns_slider.maximumValue = turns_input.value
                    turns_slider.valueOne = turns_input.value

            _update_derived_text(args.inputs)
        except Exception:
            pass


class ValidateInputsHandler(adsk.core.ValidateInputsEventHandler):
    def notify(self, args: adsk.core.ValidateInputsEventArgs):
        try:
            _get_validated_parameters(args.inputs)
            args.areInputsValid = True
        except Exception:
            args.areInputsValid = False


def run(context):
    try:
        if not UI:
            return

        cmd_def = UI.commandDefinitions.itemById(CMD_ID)
        if not cmd_def:
            cmd_def = UI.commandDefinitions.addButtonDefinition(
                CMD_ID,
                CMD_NAME,
                CMD_DESCRIPTION,
                RESOURCE_FOLDER,
            )

        on_created = CommandCreatedHandler()
        cmd_def.commandCreated.add(on_created)
        _handlers.append(on_created)

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
    global _handlers
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

        _handlers = []

    except Exception:
        if UI:
            UI.messageBox('Add-in stop failed:\n{}'.format(traceback.format_exc()))
