import adsk.core
import adsk.fusion
import math
import os
import traceback

APP = adsk.core.Application.get()
UI = APP.userInterface if APP else None

ADDIN_NAME = 'Archimedean Screw Blade Generator'
CMD_ID = 'archimedean_screw_blade_generator_cmd_v3'
CMD_NAME = 'Archimedean Screw Blade'
CMD_DESCRIPTION = 'Create a configurable hydraulic Archimedean screw flight around an existing shaft.'
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidCreatePanel'
RESOURCE_FOLDER = os.path.join(os.path.dirname(__file__), 'resources')

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
INPUT_FLIGHTS = 'flights'
INPUT_OPERATION = 'operation'
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


def _unit_direction_from_angle(
    basis_u: adsk.core.Vector3D,
    basis_v: adsk.core.Vector3D,
    angle: float,
) -> adsk.core.Vector3D:
    direction = _vector_linear_combo(basis_u, math.cos(angle), basis_v, math.sin(angle))
    if direction.length < 1e-9:
        raise ValueError('Could not resolve shaft basis vectors for helical direction.')
    direction.normalize()
    return direction


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
    for edge in shaft_face.edges:
        for i in range(edge.faces.count):
            face = edge.faces.item(i)
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
            'Unable to auto-detect shaft end face. Ensure planar shaft end caps touch the selected cylindrical face.'
        )

    selector = max if prefer_max_end else min
    return selector(candidate_faces, key=lambda f: _axis_projection(axis_origin, axis, f.pointOnFace))


def _persist_temp_body(component: adsk.fusion.Component, temp_body: adsk.fusion.BRepBody, name: str):
    design = _active_design()
    if design and design.designType == adsk.fusion.DesignTypes.ParametricDesignType:
        base_feature = component.features.baseFeatures.add()
        base_feature.name = name
        base_feature.startEdit()
        source_body = component.bRepBodies.add(temp_body, base_feature)
        base_feature.finishEdit()
        if not source_body or base_feature.bodies.count < 1:
            raise RuntimeError('Failed to add temporary body to parametric component.')
        return base_feature.bodies.item(base_feature.bodies.count - 1)

    body = component.bRepBodies.add(temp_body)
    if not body:
        raise RuntimeError('Failed to add temporary body to component.')
    body.name = name
    return body


def _create_helix_wire_body(
    temp_mgr: adsk.fusion.TemporaryBRepManager,
    axis_point: adsk.core.Point3D,
    axis_vector: adsk.core.Vector3D,
    start_point: adsk.core.Point3D,
    pitch: float,
    turns: float,
    handed_sign: float,
):
    helix = temp_mgr.createHelixWire(axis_point, axis_vector, start_point, pitch, turns * handed_sign, 0.0)
    if helix:
        return helix

    helix = temp_mgr.createHelixWire(axis_point, axis_vector, start_point, pitch * handed_sign, turns, 0.0)
    if helix:
        return helix

    flipped_axis = axis_vector.copy()
    flipped_axis.scaleBy(-1.0)
    return temp_mgr.createHelixWire(axis_point, flipped_axis, start_point, pitch, turns, 0.0)


def _create_bucket_surface_temp_body(params: dict, phase_angle: float):
    temp_mgr = adsk.fusion.TemporaryBRepManager.get()
    pitch = params['length'] / params['turns']

    inner_phase = phase_angle
    outer_phase = phase_angle + params['handedSign'] * params['bucketWrap']

    inner_dir = _unit_direction_from_angle(params['basisU'], params['basisV'], inner_phase)
    outer_dir = _unit_direction_from_angle(params['basisU'], params['basisV'], outer_phase)

    inner_start = _point_offset(params['startCenter'], inner_dir, params['innerRadius'])
    outer_start = _point_offset(params['startCenter'], outer_dir, params['outerRadius'])

    inner_wire_body = _create_helix_wire_body(
        temp_mgr,
        params['startCenter'],
        params['axisDirection'],
        inner_start,
        pitch,
        params['turns'],
        params['handedSign'],
    )
    outer_wire_body = _create_helix_wire_body(
        temp_mgr,
        params['startCenter'],
        params['axisDirection'],
        outer_start,
        pitch,
        params['turns'],
        params['handedSign'],
    )

    if not inner_wire_body or inner_wire_body.wires.count < 1:
        raise RuntimeError('Failed to build the inner helical guide wire.')
    if not outer_wire_body or outer_wire_body.wires.count < 1:
        raise RuntimeError('Failed to build the outer helical guide wire.')

    surface_body = temp_mgr.createRuledSurface(inner_wire_body.wires.item(0), outer_wire_body.wires.item(0))
    if not surface_body:
        raise RuntimeError('Failed to create ruled surface between helical guides.')

    return surface_body


def _thicken_surface_body(component: adsk.fusion.Component, surface_body: adsk.fusion.BRepBody, thickness: float):
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
    if thicken_feature.bodies.count < 1:
        raise RuntimeError('Thicken failed to create a solid blade body.')

    try:
        surface_body.deleteMe()
    except Exception:
        try:
            surface_body.isLightBulbOn = False
        except Exception:
            pass

    blade_body = thicken_feature.bodies.item(0)
    blade_body.name = 'Archimedean Flight'
    return blade_body


def _create_single_flight(component: adsk.fusion.Component, params: dict, phase_angle: float):
    surface_temp = _create_bucket_surface_temp_body(params, phase_angle)
    surface_body = _persist_temp_body(component, surface_temp, 'Archimedean Flight Surface')
    return _thicken_surface_body(component, surface_body, params['thickness'])


def _join_bodies(component: adsk.fusion.Component, target_body: adsk.fusion.BRepBody, tool_bodies):
    if not tool_bodies:
        return target_body

    tools = adsk.core.ObjectCollection.create()
    for body in tool_bodies:
        tools.add(body)

    combine_input = component.features.combineFeatures.createInput(target_body, tools)
    combine_input.operation = adsk.fusion.FeatureOperations.JoinFeatureOperation
    combine_input.isKeepToolBodies = False
    component.features.combineFeatures.add(combine_input)
    return target_body


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

    start_end_input = adsk.core.DropDownCommandInput.cast(inputs.itemById(INPUT_START_END))
    start_end_name = start_end_input.selectedItem.name if start_end_input and start_end_input.selectedItem else START_END_MIN
    prefer_max_end = start_end_name == START_END_MAX

    start_planar_face = _auto_detect_start_end_face(shaft_face, cylinder, prefer_max_end)

    axis = cylinder.axis.copy()
    axis.normalize()
    axis_origin = cylinder.origin.copy()

    axis_direction = axis.copy()
    if prefer_max_end:
        axis_direction.scaleBy(-1.0)

    start_scalar = _axis_projection(axis_origin, axis, start_planar_face.pointOnFace)
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

    length_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_LENGTH))
    outer_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_OUTER_RADIUS))
    thickness_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_THICKNESS))
    clearance_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_CLEARANCE))
    start_angle_input = adsk.core.ValueCommandInput.cast(inputs.itemById(INPUT_START_ANGLE))
    turns_input = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(INPUT_TURNS))
    bucket_wrap_input = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(INPUT_BUCKET_WRAP))
    flights_input = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById(INPUT_FLIGHTS))
    hand_input = adsk.core.DropDownCommandInput.cast(inputs.itemById(INPUT_HANDEDNESS))
    operation_input = adsk.core.DropDownCommandInput.cast(inputs.itemById(INPUT_OPERATION))

    length = length_input.value
    outer_radius = outer_input.value
    thickness = thickness_input.value
    clearance = clearance_input.value
    start_angle = start_angle_input.value
    turns = turns_input.value
    bucket_wrap_deg = bucket_wrap_input.value
    flights = flights_input.value

    if length <= 0:
        raise ValueError('Blade length must be greater than zero.')
    if turns <= 0:
        raise ValueError('Turns must be greater than zero.')
    if thickness <= 0:
        raise ValueError('Blade thickness must be greater than zero.')
    if clearance < 0:
        raise ValueError('Hub clearance cannot be negative.')
    if flights < 1:
        raise ValueError('Flights must be at least 1.')
    if bucket_wrap_deg < 0 or bucket_wrap_deg > 120:
        raise ValueError('Bucket wrap must be between 0 and 120 degrees.')

    inner_radius = cylinder.radius + clearance
    if outer_radius <= inner_radius:
        raise ValueError('Outer radius must be greater than shaft radius + hub clearance.')

    handedness = hand_input.selectedItem.name if hand_input and hand_input.selectedItem else HAND_RIGHT
    handed_sign = 1.0 if handedness == HAND_RIGHT else -1.0

    operation_name = operation_input.selectedItem.name if operation_input and operation_input.selectedItem else OP_NEW_BODY
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
        'length': length,
        'turns': turns,
        'thickness': thickness,
        'startAngle': start_angle,
        'bucketWrap': math.radians(bucket_wrap_deg),
        'handedSign': handed_sign,
        'flights': preview_flights,
        'joinToShaft': join_to_shaft,
    }


def _build_blades(params: dict):
    component = params['component']

    flight_bodies = []
    for i in range(params['flights']):
        phase = params['startAngle'] + (2.0 * math.pi * i / params['flights'])
        flight_bodies.append(_create_single_flight(component, params, phase))

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
            f"Bucket wrap: {math.degrees(params['bucketWrap']):.1f} deg"
        )
    except Exception as ex:
        text_input.text = str(ex)


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def notify(self, args: adsk.core.CommandEventArgs):
        try:
            command = args.firingEvent.sender
            params = _get_validated_parameters(command.commandInputs)
            _build_blades(params)
        except Exception:
            if UI:
                UI.messageBox('Failed to create Archimedean blade:\n{}'.format(traceback.format_exc()))


class CommandPreviewHandler(adsk.core.CommandEventHandler):
    def notify(self, args: adsk.core.CommandEventArgs):
        try:
            command = args.firingEvent.sender
            params = _get_validated_parameters(command.commandInputs, preview_mode=True)
            _build_blades(params)
            args.isValidResult = True
        except Exception:
            args.isValidResult = False


class InputChangedHandler(adsk.core.InputChangedEventHandler):
    def notify(self, args: adsk.core.InputChangedEventArgs):
        try:
            changed = args.input
            inputs = args.inputs

            turns_input = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(INPUT_TURNS))
            turns_slider = adsk.core.FloatSliderCommandInput.cast(inputs.itemById(INPUT_TURNS_SLIDER))
            bucket_input = adsk.core.FloatSpinnerCommandInput.cast(inputs.itemById(INPUT_BUCKET_WRAP))
            bucket_slider = adsk.core.FloatSliderCommandInput.cast(inputs.itemById(INPUT_BUCKET_WRAP_SLIDER))

            if changed and turns_input and turns_slider:
                if changed.id == INPUT_TURNS_SLIDER:
                    turns_input.value = turns_slider.valueOne
                elif changed.id == INPUT_TURNS:
                    if turns_input.value > turns_slider.maximumValue:
                        turns_slider.maximumValue = turns_input.value
                    turns_slider.valueOne = turns_input.value

            if changed and bucket_input and bucket_slider:
                if changed.id == INPUT_BUCKET_WRAP_SLIDER:
                    bucket_input.value = bucket_slider.valueOne
                elif changed.id == INPUT_BUCKET_WRAP:
                    bucket_slider.valueOne = bucket_input.value

            _update_derived_text(inputs)
        except Exception:
            pass


class ValidateInputsHandler(adsk.core.ValidateInputsEventHandler):
    def notify(self, args: adsk.core.ValidateInputsEventArgs):
        try:
            _get_validated_parameters(args.inputs)
            args.areInputsValid = True
        except Exception:
            args.areInputsValid = False


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
                'Select the shaft cylindrical face.',
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
                adsk.core.ValueInput.createByString(f'80 mm'),
            )
            inputs.addValueInput(
                INPUT_LENGTH,
                'Blade Length',
                units,
                adsk.core.ValueInput.createByString(f'600 mm'),
            )

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

            inputs.addValueInput(
                INPUT_THICKNESS,
                'Blade Thickness',
                units,
                adsk.core.ValueInput.createByString('3 mm'),
            )
            inputs.addValueInput(
                INPUT_CLEARANCE,
                'Hub Clearance',
                units,
                adsk.core.ValueInput.createByString('2 mm'),
            )

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
