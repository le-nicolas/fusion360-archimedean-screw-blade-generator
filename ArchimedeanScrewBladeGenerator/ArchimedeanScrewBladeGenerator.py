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
INPUT_START_PLANE = 'startPlane'
INPUT_OUTER_RADIUS = 'outerRadius'
INPUT_LENGTH = 'bladeLength'
INPUT_TURNS = 'turns'
INPUT_THICKNESS = 'bladeThickness'
INPUT_CLEARANCE = 'hubClearance'
INPUT_START_ANGLE = 'startAngle'
INPUT_HANDEDNESS = 'handedness'
INPUT_FLIGHTS = 'flights'
INPUT_SEGMENTS = 'segmentsPerTurn'
INPUT_OPERATION = 'operation'
INPUT_DERIVED = 'derivedInfo'

HAND_RIGHT = 'Right-handed'
HAND_LEFT = 'Left-handed'
OP_NEW_BODY = 'New Blade Body'
OP_JOIN = 'Join Blade To Shaft'

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


def _planar_entity_info(entity):
    if not entity:
        return None, None, None

    if entity.objectType == adsk.fusion.BRepFace.classType():
        face = adsk.fusion.BRepFace.cast(entity)
        plane = adsk.core.Plane.cast(face.geometry)
        if not plane:
            return None, None, None
        return face, plane, face.body.parentComponent

    if entity.objectType == adsk.fusion.ConstructionPlane.classType():
        plane_entity = adsk.fusion.ConstructionPlane.cast(entity)
        plane = adsk.core.Plane.cast(plane_entity.geometry)
        if not plane:
            return None, None, None
        return plane_entity, plane, plane_entity.parentComponent

    return None, None, None


def _get_validated_parameters(inputs: adsk.core.CommandInputs):
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

    start_entity = _selection_entity(inputs, INPUT_START_PLANE)
    if not start_entity:
        raise ValueError('Select a start plane or planar face.')

    if hasattr(start_entity, 'assemblyContext') and start_entity.assemblyContext:
        raise ValueError('Assembly occurrences are not supported. Select entities in their source component.')

    start_planar_entity, start_plane, start_component = _planar_entity_info(start_entity)
    if not start_planar_entity or not start_plane:
        raise ValueError('Start selection must be a planar face or construction plane.')

    component = shaft_face.body.parentComponent
    if component != start_component:
        raise ValueError('Shaft face and start plane must belong to the same component.')

    axis = cylinder.axis.copy()
    axis.normalize()
    normal = start_plane.normal.copy()
    normal.normalize()

    if abs(axis.dotProduct(normal)) < 0.995:
        raise ValueError('Start plane must be normal to the shaft axis.')

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

    return {
        'component': component,
        'shaftBody': shaft_face.body,
        'startPlanarEntity': start_planar_entity,
        'innerRadius': inner_radius,
        'outerRadius': outer_radius,
        'length': length,
        'turns': turns,
        'thickness': thickness,
        'startAngle': start_angle,
        'handedSign': handed_sign,
        'flights': flights,
        'segmentsPerTurn': segments_per_turn,
        'stationCount': station_count,
        'joinToShaft': join_to_shaft,
    }


def _add_profile_section(
    component: adsk.fusion.Component,
    base_entity,
    distance: float,
    inner_radius: float,
    outer_radius: float,
    thickness: float,
    angle: float,
):
    planes = component.constructionPlanes
    plane_input = planes.createInput()
    plane_input.setByOffset(base_entity, adsk.core.ValueInput.createByReal(distance))
    section_plane = planes.add(plane_input)
    section_plane.isLightBulbOn = False

    sketch = component.sketches.add(section_plane)
    sketch.isLightBulbOn = False

    cos_a = math.cos(angle)
    sin_a = math.sin(angle)

    radial_x = cos_a
    radial_y = sin_a
    tangential_x = -sin_a
    tangential_y = cos_a

    half_t = thickness / 2.0

    inner_x = inner_radius * radial_x
    inner_y = inner_radius * radial_y
    outer_x = outer_radius * radial_x
    outer_y = outer_radius * radial_y

    p1 = adsk.core.Point3D.create(inner_x + tangential_x * half_t, inner_y + tangential_y * half_t, 0)
    p2 = adsk.core.Point3D.create(outer_x + tangential_x * half_t, outer_y + tangential_y * half_t, 0)
    p3 = adsk.core.Point3D.create(outer_x - tangential_x * half_t, outer_y - tangential_y * half_t, 0)
    p4 = adsk.core.Point3D.create(inner_x - tangential_x * half_t, inner_y - tangential_y * half_t, 0)

    lines = sketch.sketchCurves.sketchLines
    lines.addByTwoPoints(p1, p2)
    lines.addByTwoPoints(p2, p3)
    lines.addByTwoPoints(p3, p4)
    lines.addByTwoPoints(p4, p1)

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
        distance = params['length'] * t
        angle = phase_angle + params['handedSign'] * (2.0 * math.pi * params['turns'] * t)

        profile = _add_profile_section(
            params['component'],
            params['startPlanarEntity'],
            distance,
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
    start_entity = _selection_entity(inputs, INPUT_START_PLANE)

    if not shaft_entity or not start_entity:
        text_input.text = 'Select shaft + start plane to view derived values.'
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

            start_input = inputs.addSelectionInput(
                INPUT_START_PLANE,
                'Start Plane / Face',
                'Select a planar face or construction plane normal to shaft axis',
            )
            start_input.addSelectionFilter('PlanarFaces')
            start_input.addSelectionFilter('ConstructionPlanes')
            start_input.setSelectionLimits(1, 1)

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

            inputs.addFloatSpinnerCommandInput(INPUT_TURNS, 'Turns', 0.1, 200.0, 0.1, 3.0)

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
                'Select shaft + start plane to view derived values.',
                3,
                True,
            )

            on_execute = CommandExecuteHandler()
            command.execute.add(on_execute)
            _handlers.append(on_execute)

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
