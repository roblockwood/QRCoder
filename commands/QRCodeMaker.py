"""
QRCoder, a Fusion 360 add-in
================================
QRCoder is a Fusion 360 add-in for the creation of 3D QR Codes.

:copyright: (c) 2021 by Patrick Rainsberry.
:license: MIT, see LICENSE for more details.


THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


QRCoder leverages the pyqrcode library:

    https://github.com/mnooner256/pyqrcode
    Copyright (c) 2013, Michael Nooner
    All rights reserved.

    Redistribution and use in source and binary forms, with or without
    modification, are permitted provided that the following conditions are met:
        * Redistributions of source code must retain the above copyright
          notice, this list of conditions and the following disclaimer.
        * Redistributions in binary form must reproduce the above copyright
          notice, this list of conditions and the following disclaimer in the
          documentation and/or other materials provided with the distribution.
        * Neither the name of the copyright holder nor the names of its
          contributors may be used to endorse or promote products derived from
          this software without specific prior written permission

    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
    AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
    IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
    ARE DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
    DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
    (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
    LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
    ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
    (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
    SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""
import csv
import os.path

import adsk.core
import adsk.fusion
import adsk.cam

from ..apper import apper
from .. import config

# Defaults
HEIGHT = '.4 mm'
BASE = '0.0 in'
MESSAGE = 'input here'
QR_TARGET_SIZE_MM = '24 mm'  # Target QR code size in millimeters
FILE_NAME = 'QR-17x.csv'
BLOCK = '.03 in'

# Flag to signal if the batch generation should occur on execute
# This needs to be a class attribute to persist between on_input_changed and on_execute.
# It's set by successful CSV load, and checked in on_execute.
# Initialized to False.


def get_target_body(sketch_point):
    """Finds the target body at the given sketch point."""
    ao = apper.AppObjects()
    target_collection = ao.root_comp.findBRepUsingPoint(
        sketch_point.worldGeometry, adsk.fusion.BRepEntityTypes.BRepBodyEntityType, -1.0, True
    )

    if target_collection.count > 0:
        return target_collection.item(0)
    else:
        return None


def make_real_geometry(target_body: adsk.fusion.BRepBody, input_values, qr_data):
    """
    Creates the QR code geometry in a new component.

    Args:
        target_body: The target body.
        input_values: Input values from the command.
        qr_data: The QR code data.

    Returns:
        The newly created component containing the QR code geometry, or None on failure.
    """
    ao = apper.AppObjects()
    design = ao.design
    root_comp = design.rootComponent

    if not target_body:
        ao.ui.messageBox("No target body found.")
        return None  # Return None if no target body

    parent_comp = target_body.parentComponent

    # Create a new occurrence/component in the root
    transform = adsk.core.Matrix3D.create()
    new_occurrence = root_comp.occurrences.addNewComponent(transform)
    new_comp = new_occurrence.component

    # Update the component occurrence name
    # Use the message from input_values (set in on_preview/on_input_changed for CSV mode)
    # or a default if not available.
    new_comp.name = input_values.get('message', 'Imported QR Code')

    # Copy all bodies from the original component to the new one
    for body in parent_comp.bRepBodies:
        if body.isSolid:
            body.copyToComponent(new_occurrence)

    # Create the QR code temp geometry using default height/base
    # Pass input_values for qr_size and sketch_point
    temp_body = get_qr_temp_geometry(qr_data, input_values)

    # Add the QR code geometry as a base feature to the new component
    if temp_body: # Only add if geometry was created
        base_feature = new_comp.features.baseFeatures.add()
        base_feature.startEdit()
        new_comp.bRepBodies.add(temp_body, base_feature)
        base_feature.finishEdit()

        # Find the newly created body and rename it.
        for body in new_comp.bRepBodies:
            # Check if the body name is the default name for a box created by createBox
            # This might need adjustment if other box-like bodies are copied from the parent.
            # A more robust way might be to check the body's creation feature or properties
            # if possible, but for this script's structure, checking the default name is likely sufficient.
            if body.name == "Solid1":
                body.name = "text"
                break
    else:
        return None # Return None if no temp body was created


    return new_comp  # Return None if no target body


# Note: clear_graphics and make_graphics are not actively used in the current preview/generation logic
# but are kept in case they are needed for future visualization features.
def clear_graphics(graphics_group: adsk.fusion.CustomGraphicsGroup):
    """Clears all custom graphics from a group."""
    # Iterate through the group in reverse order to avoid issues with deleted items
    for i in range(graphics_group.count - 1, -1, -1):
        entity = graphics_group.item(i)
        if entity.isValid:
            entity.deleteMe()


def make_graphics(t_body: adsk.fusion.BRepBody, graphics_group: adsk.fusion.CustomGraphicsGroup):
    """Creates custom graphics for a temporary body."""
    clear_graphics(graphics_group)
    color = adsk.core.Color.create(250, 162, 27, 255) # Example color (Orange)
    color_effect = adsk.fusion.CustomGraphicsSolidColorEffect.create(color)
    graphics_body = graphics_group.addBRepBody(t_body)
    graphics_body.color = color_effect


def get_qr_temp_geometry(qr_data, input_values):
    """
    Creates temporary BRep geometry for the QR code based on the QR data.

    Args:
        qr_data: A list of lists containing '0' or '1' characters representing the QR code modules.
        input_values: Input values from the command, including 'qr_size' and 'sketch_point'.

    Returns:
        A temporary BRepBody representing the QR code geometry, or None if QR data is empty or inputs are invalid.
    """
    ao = apper.AppObjects()
    units_manager = ao.units_manager

    # Corrected: Pass the string expressions directly to evaluateExpression
    try:
        height: float = units_manager.evaluateExpression(HEIGHT, units_manager.defaultLengthUnits)
        base: float = units_manager.evaluateExpression(BASE, units_manager.defaultLengthUnits)
    except Exception as e:
        ao.ui.messageBox(f"Error evaluating default height or base: {e}")
        return None


    # Ensure sketch_point is valid before accessing its properties
    if 'sketch_point' not in input_values or not input_values['sketch_point']:
        # Error message shown in calling function (on_preview/on_input_changed)
        return None

    sketch_point: adsk.fusion.SketchPoint = input_values['sketch_point'][0]

    # Ensure the parent sketch and its directions are valid
    if not sketch_point.parentSketch or not sketch_point.parentSketch.xDirection or not sketch_point.parentSketch.yDirection:
         ao.ui.messageBox("Error: Invalid sketch associated with the selected point.")
         return None


    x_dir = sketch_point.parentSketch.xDirection
    x_dir.normalize()
    y_dir = sketch_point.parentSketch.yDirection
    y_dir.normalize()
    z_dir = x_dir.copy() # Create a copy before cross product
    z_dir.crossProduct(y_dir)
    z_dir.normalize()

    qr_size = len(qr_data)
    if qr_size == 0:
        return None # Return None if no QR data

    # Calculate block size from the overall size and qr_size
    # Ensure 'qr_size' is in input_values and is a valid number
    if 'qr_size' not in input_values or not isinstance(input_values['qr_size'], (int, float)):
         # Error message shown in calling function (on_preview/on_input_changed)
         return None

    overall_size: float = input_values['qr_size']
    if qr_size > 0:
        side = overall_size / qr_size
    else:
        side = 0 # Avoid division by zero if qr_size is 0

    input_values['block_size'] = side # store block size

    middle_point = sketch_point.worldGeometry
    start_point = middle_point.copy()

    # Corrected start point calculation
    x_start_move = x_dir.copy()
    x_start_move.scaleBy(-side * (qr_size / 2.0))
    start_point.translateBy(x_start_move)

    y_start_move = y_dir.copy()
    y_start_move.scaleBy(side * (qr_size / 2.0))
    start_point.translateBy(y_start_move)

    z_start_move = z_dir.copy()
    z_start_move.scaleBy((.5 * height) + base)
    start_point.translateBy(z_start_move)

    base_point = middle_point.copy()
    z_base_move = z_dir.copy()
    z_base_move.scaleBy((.5 * base))
    base_point.translateBy(z_base_move)

    b_mgr = adsk.fusion.TemporaryBRepManager.get()

    base_t_body = None
    has_base = (base > 0)
    if has_base:
        full_size = side * qr_size
        base_t_box = adsk.core.OrientedBoundingBox3D.create(base_point, x_dir, y_dir, full_size, full_size, base)
        base_t_body = b_mgr.createBox(base_t_box)

    # Build the QR code blocks
    for i, row in enumerate(qr_data):
        for j, col in enumerate(row):
            # Ensure col is a string before checking and converting
            if isinstance(col, str) and col == '1':
                x_move = x_dir.copy()
                y_move = y_dir.copy()
                x_move.scaleBy(j * side)
                y_move.scaleBy(-1 * i * side) # QR codes are typically rendered top-down

                c_point = start_point.copy()
                c_point.translateBy(x_move)
                c_point.translateBy(y_move)

                # Create a box for the QR module
                b_box = adsk.core.OrientedBoundingBox3D.create(c_point, x_dir, y_dir, side, side, height + base)
                t_body = b_mgr.createBox(b_box)

                if base_t_body is None:
                    base_t_body = t_body
                else:
                    # Perform boolean union with the base body or previous blocks
                    b_mgr.booleanOperation(base_t_body, t_body, adsk.fusion.BooleanTypes.UnionBooleanType)

    return base_t_body


def read_csv_and_extract_keys(file_path):
    """Reads a CSV file and extracts values from the 'KEY' column."""
    keys = []
    if os.path.exists(file_path):
        try:
            with open(file_path, newline='') as f:
                reader = csv.DictReader(f)
                if 'KEY' in reader.fieldnames:
                    for row in reader:
                        # Append the key value as a string
                        keys.append(str(row.get('KEY', ''))) # Use .get with default for safety
                else:
                    ao = apper.AppObjects()
                    ao.ui.messageBox("CSV file does not contain a 'KEY' header.")
        except Exception as e:
            ao = apper.AppObjects()
            ao.ui.messageBox(f"Error reading CSV file: {e}")
    return keys


def import_qr_from_file(file_name):
    """
    This function is currently not used in the updated logic,
    as QR data is generated from a key extracted from the CSV.
    It remains here in case future functionality requires reading
    a CSV that directly defines a QR pattern.
    """
    qr_data = []

    if os.path.exists(file_name):
        try:
            with open(file_name, newline='') as f:
                reader = csv.reader(f)
                # Assuming the CSV contains rows of '0' and '1' characters
                qr_data = [row for row in reader if row] # Read all rows, filtering empty ones
        except Exception as e:
             ao = apper.AppObjects()
             ao.ui.messageBox(f"Error reading QR pattern from CSV file: {e}")


    return qr_data


@apper.lib_import(config.lib_path)
def build_qr_code(message, args):
    """Builds QR code data (list of lists of '0'/'1' strings) from a message."""
    import pyqrcode
    try:
        # Ensure message is a string
        message_str = str(message)
        qr = pyqrcode.create(message_str, **args)
        qr_text = qr.text(quiet_zone=0)
        # Split into lines, strip whitespace, filter empty lines, then split each line into characters
        qr_data = [[char for char in y] for y in (x.strip() for x in qr_text.splitlines()) if y]
        return qr_data

    except ValueError as e:
        ao = apper.AppObjects()
        ao.ui.messageBox(f'Problem generating QR code from message: {e}')
        return []
    except Exception as e:
        ao = apper.AppObjects()
        ao.ui.messageBox(f'An unexpected error occurred during QR code generation: {e}')
        return []


def make_qr_from_message(input_values):
    """Generates QR code data from the 'message' input value."""
    # Ensure 'message' is in input_values
    if 'message' not in input_values:
        ao = apper.AppObjects()
        ao.ui.messageBox("Error: 'message' input value is missing.")
        return []

    message: str = input_values['message']
    # use_user_size: bool = input_values['use_user_size'] # Removed
    # user_size: int = input_values['user_size'] # Removed
    # mode: str = input_values['mode'] # Removed
    # error_type: str = input_values['error_type'] # Removed

    args = {}
    # if use_user_size: # Removed
    #     args['version'] = user_size
    # if mode != 'Automatic': # Removed
    #     args['mode'] = mode
    # if error_type != 'Automatic': # Removed
    #     args['error'] = error_type

    # Check for pyqrcode dependency before building
    success = apper.check_dependency('pyqrcode', config.lib_path)

    if success:
        qr_data = build_qr_code(message, args)
        return qr_data
    else:
        ao = apper.AppObjects()
        ao.ui.messageBox("Error: 'pyqrcode' library not found. Please ensure it is installed.")
        return []


def add_make_inputs(inputs: adsk.core.CommandInputs):
    """Adds inputs for creating a QR code from a message."""
    drop_style = adsk.core.DropDownStyles.TextListDropDownStyle
    inputs.addStringValueInput('message', 'Value to encode', MESSAGE)
    # inputs.addBoolValueInput('use_user_size', 'Specify size?', True, '', False) # Removed
    # size_spinner = inputs.addIntegerSpinnerCommandInput('user_size', 'QR Code Size (Version)', 1, 40, 1, 5) # Removed
    # size_spinner.isEnabled = False # Removed
    # size_spinner.isVisible = False # Removed
    #
    # mode_drop_down = inputs.addDropDownCommandInput('mode', 'Encoding Mode', drop_style) # Removed
    # mode_items = mode_drop_down.listItems # Removed
    # mode_items.add('Automatic', True, '') # Removed
    # mode_items.add('alphanumeric', False, '') # Removed
    # mode_items.add('numeric', False, '') # Removed
    # mode_items.add('binary', False, '') # Removed
    # mode_items.add('kanji', False, '') # Removed
    #
    # error_input = inputs.addDropDownCommandInput('error_type', 'Encoding Mode', drop_style) # Removed
    # error_items = error_input.listItems # Removed
    # error_items.add('Automatic', True, '') # Removed
    # error_items.add('L', False, '') # Removed
    # error_items.add('M', False, '') # Removed
    # error_items.add('Q', False, '') # Removed
    # error_items.add('H', False, '') # Removed



def add_csv_inputs(inputs: adsk.core.CommandInputs):
    """Adds inputs for importing QR codes from a CSV file."""
    # Removed the 'generate_csv_button' input as it's no longer used to trigger the action.

    file_name_input = inputs.addStringValueInput('file_name', "File to import", '')
    file_name_input.isReadOnly = True # Make the file name input read-only

    browse_button = inputs.addBoolValueInput('browse', 'Browse', False, '', False)
    browse_button.isFullWidth = True

    # Add a read-only text box to display the extracted keys
    # Set number of rows to accommodate up to 30 keys vertically
    inputs.addTextBoxCommandInput('extracted_keys', 'Extracted Keys', '', 30, True)


# Create file browser dialog box
def browse_for_csv():
    """Opens a file dialog for selecting a CSV file."""
    ao = apper.AppObjects()

    file_dialog = ao.ui.createFileDialog()
    file_dialog.initialDirectory = config.app_path
    file_dialog.filter = ".csv files (*.csv)"
    file_dialog.isMultiSelectEnabled = False
    file_dialog.title = 'Select csv file to import'
    dialog_results = file_dialog.showOpen()

    if dialog_results == adsk.core.DialogResults.DialogOK:
        file_names = file_dialog.filenames
        return file_names[0]
    else:
        return ''


def export_step_file(component: adsk.fusion.Component):
    """Exports a component as a STEP file."""
    ao = apper.AppObjects()
    export_manager = ao.design.exportManager
    step_options = export_manager.createStepOptions()

    # Get the default export folder
    default_folder = export_manager.defaultFolderPath

    # Create a filename based on the component name
    file_name = os.path.join(default_folder, f"{component.name}.step")
    step_options.filename = file_name
    step_options.component = component

    try:
        export_manager.execute(step_options)
        ao.ui.messageBox(f'Successfully exported STEP file to: {file_name}')
    except RuntimeError as error:
        ao.ui.messageBox(f'Failed to export STEP file: {error}')



class QRCodeMaker(apper.Fusion360CommandBase):
    """Fusion 360 Command for creating QR codes from message or CSV."""
    def __init__(self, name: str, options: dict):
        super().__init__(name, options)
        # self.graphics_group = None # Initialize graphics group if needed for preview graphics
        self.make_preview = True
        self.is_make_qr = options.get('is_make_qr', False)
        self.extracted_keys = [] # Store extracted keys
        # Reintroduced the flag to signal batch generation in on_execute
        self._batch_generate_triggered = False


    def on_input_changed(self, command, inputs, changed_input, input_values):
        """Handles changes to command inputs."""
        ao = apper.AppObjects()

        # Always trigger preview unless specifically disabled below
        self.make_preview = True
        # Reset batch trigger flag unless a CSV with keys is loaded
        self._batch_generate_triggered = False

        # if changed_input.id == 'use_user_size': # Removed
        #     if input_values['use_user_size']:
        #         inputs.itemById('user_size').isEnabled = True
        #         inputs.itemById('user_size').isVisible = True
        #     else:
        #         inputs.itemById('user_size').isEnabled = False
        #         inputs.itemById('user_size').isVisible = False

        if changed_input.id == 'browse':
            changed_input.value = False # Reset browse button state
            file_name = browse_for_csv()
            if len(file_name) > 0:
                inputs.itemById('file_name').value = file_name
                # Read the CSV and display keys when a file is selected
                self.extracted_keys = read_csv_and_extract_keys(file_name)
                # Display keys vertically
                inputs.itemById('extracted_keys').text = '\n'.join(self.extracted_keys) if self.extracted_keys else 'No keys found or "KEY" header missing.'

                # --- MODIFIED LOGIC HERE ---
                # If keys were successfully extracted, disable preview and set the batch trigger flag
                if self.extracted_keys:
                    self.make_preview = False # Disable preview
                    self._batch_generate_triggered = True # Set flag to trigger batch generate on execute
                    # Do NOT call command.doExecute() here
                else:
                    self.make_preview = True # Re-enable preview if no keys found
                    self._batch_generate_triggered = False # Ensure flag is false
                # ---------------------------

            else:
                 # Clear extracted keys display, re-enable preview, and reset flag if no file is selected
                 self.extracted_keys = []
                 inputs.itemById('extracted_keys').text = ''
                 self.make_preview = True # Re-enable preview
                 self._batch_generate_triggered = False # Reset flag


        elif changed_input.id == 'file_name':
             # Also update keys if the file name is manually changed (though input is read-only)
            file_name = input_values.get('file_name', '') # Use .get for safety
            if len(file_name) > 0:
                self.extracted_keys = read_csv_and_extract_keys(file_name)
                # Display keys vertically
                inputs.itemById('extracted_keys').text = '\n'.join(self.extracted_keys) if self.extracted_keys else 'No keys found or "KEY" header missing.'

                # --- MODIFIED LOGIC HERE ---
                # If keys were successfully extracted, disable preview and set the batch trigger flag
                if self.extracted_keys:
                    self.make_preview = False # Disable preview
                    self._batch_generate_triggered = True # Set flag to trigger batch generate on execute
                    # Do NOT call command.doExecute() here
                else:
                    self.make_preview = True # Re-enable preview if no keys found
                    self._batch_generate_triggered = False # Ensure flag is false
                # ---------------------------

            else:
                 # Clear extracted keys display, re-enable preview, and reset flag if file name is cleared
                 self.extracted_keys = []
                 inputs.itemById('extracted_keys').text = ''
                 self.make_preview = True # Re-enable preview
                 self._batch_generate_triggered = False # Reset flag


        # Removed the generate_csv_button handler


    def on_preview(self, command, inputs, args, input_values):
        """Handles the preview event."""
        # Only run preview if make_preview is True (controlled by input_changed)
        if self.make_preview:
            ao = apper.AppObjects()
            qr_data = []
            message_to_encode = None # Variable to hold the message used for QR generation

            # In CSV import mode, preview should be disabled if a file with keys is loaded.
            # This check is mostly for safety, as make_preview should be False in that case.
            if not self.is_make_qr and (len(input_values.get('file_name', '')) > 0 and self.extracted_keys):
                 args.isValidResult = False # No valid preview in CSV mode with loaded file
                 return # Exit preview early


            # Determine message to encode based on mode
            if not self.is_make_qr: # CSV import mode (will only preview if no file/keys loaded yet)
                # Use a placeholder message for preview if no file is loaded
                message_to_encode = "Preview: Select CSV"
                qr_data = make_qr_from_message({'message': message_to_encode})

            else: # Make QR Code mode
                 # Use the message input directly
                message_to_encode = input_values.get('message', MESSAGE)
                qr_data = make_qr_from_message({'message': message_to_encode})


            # Set the message in input_values for component naming in make_real_geometry
            # Use the message_to_encode if available, otherwise a default
            input_values['message'] = message_to_encode if message_to_encode is not None else 'Imported QR Code Preview'


            # Only attempt to create geometry if QR data was successfully generated
            if len(qr_data) > 0:
                # Ensure sketch_point is selected before proceeding
                if 'sketch_point' not in input_values or not input_values['sketch_point']:
                    # No message box here, let the user pick the point
                    args.isValidResult = False # Invalidate the preview result if no point selected
                    # Don't set make_preview to False here, allow it to re-run when point is selected
                    return # Exit the function

                sketch_point = input_values['sketch_point'][0]
                target_body = get_target_body(sketch_point)

                # Updated to match new function signature, and removed export
                new_component = make_real_geometry(target_body, input_values, qr_data)

                # Set isValidResult based on whether geometry was created
                args.isValidResult = (new_component is not None)

            else:
                 # If no QR data, the preview is not valid
                 args.isValidResult = False

            # The make_preview flag is now controlled by input_changed, no need to set to False here.


    def on_execute(self, command, inputs, args, input_values):
        """Handles the execute event (triggered by OK button)."""
        ao = apper.AppObjects()
        design = ao.design
        generated_count = 0

        # Check if we are in CSV mode and the batch trigger flag is set
        if not self.is_make_qr and self._batch_generate_triggered:
            # Reset the flag immediately
            self._batch_generate_triggered = False

            # Ensure sketch point and QR size are selected
            if 'sketch_point' not in input_values or not input_values['sketch_point']:
                ao.ui.messageBox("Please select a sketch point for the QR code location.")
                args.executeFailed = True # Indicate execution failed
                return

            if 'qr_size' not in input_values or not isinstance(input_values['qr_size'], (int, float)):
                 ao.ui.messageBox("Please specify a valid QR Code Size.")
                 args.executeFailed = True # Indicate execution failed
                 return

            sketch_point = input_values['sketch_point'][0]
            target_body = get_target_body(sketch_point)

            if not target_body:
                 ao.ui.messageBox("No target body found at the selected sketch point.")
                 args.executeFailed = True # Indicate execution failed
                 return

            # Loop through ALL keys in the extracted_keys list
            if len(self.extracted_keys) > 0:
                # Ensure design is available
                if design:
                    timeline = design.timeline
                    timelineGroups = timeline.timelineGroups
                    # Add the timeline group before creating features
                    start_index = timeline.count + 1
                    # Check if timelineGroups is valid before adding
                    newTimelineGroup = None # Initialize to None
                    if timelineGroups:
                         newTimelineGroup = timelineGroups.add(start_index, start_index)
                         newTimelineGroup.name = "Generated QR Codes"

                    for key in self.extracted_keys: # Loop through ALL keys
                        message_to_encode = key
                        qr_data = make_qr_from_message({'message': message_to_encode})

                        if len(qr_data) > 0:
                            # Create geometry for this key
                            # Pass a dictionary to make_real_geometry with necessary input_values
                            temp_input_values = {
                                'message': message_to_encode, # Use the current key as the message
                                'qr_size': input_values['qr_size'],
                                'sketch_point': input_values['sketch_point'] # Pass the sketch point
                            }
                            new_component = make_real_geometry(target_body, temp_input_values, qr_data)
                            if new_component:
                                generated_count += 1
                                # Optional: Export STEP file for each generated component
                                # export_step_file(new_component)
                        else:
                            ao.ui.messageBox(f"Could not generate QR code for key: {key}")

                    # Update the end index of the timeline group after creation
                    # Check if newTimelineGroup is valid before updating
                    if generated_count > 0 and newTimelineGroup:
                        newTimelineGroup.endTimeStep = timeline.count
                else:
                     ao.ui.messageBox("Error: Could not access the active design.")


            # Display the count of generated QR codes
            ao.ui.messageBox(f"Generated {generated_count} QR codes from the CSV.")

            # The command will automatically terminate after on_execute completes successfully.
            # No explicit command.terminate() is needed here.

        # This is the original single QR creation logic for the OK button
        # It will still run if not in CSV mode OR if in CSV mode but no keys were loaded
        # (meaning the batch trigger flag was not set).
        elif self.is_make_qr or (not self.is_make_qr and not self._batch_generate_triggered):
             # If in CSV mode and no keys were loaded, the user might still click OK
             # In this case, it will just close the dialog as no batch generation is triggered.

            if self.is_make_qr: # Only run single creation if in Make QR mode
                ao = apper.AppObjects()
                qr_data = make_qr_from_message(input_values)

                if len(qr_data) > 0:
                     # Ensure sketch_point is selected before proceeding
                    if 'sketch_point' not in input_values or not input_values['sketch_point']:
                        ao.ui.messageBox("Please select a sketch point for the QR code location.")
                        args.executeFailed = True # Indicate execution failed
                        return # Exit the function

                    sketch_point: adsk.fusion.SketchPoint = input_values['sketch_point'][0]
                    target_body = get_target_body(sketch_point)

                    if not target_body:
                         ao.ui.messageBox("No target body found at the selected sketch point.")
                         args.executeFailed = True # Indicate execution failed
                         return

                    # Set the message in input_values for component naming
                    input_values['message'] = input_values.get('message', MESSAGE)

                    new_component = make_real_geometry(target_body, input_values, qr_data)

                    # Optional: Export STEP file after successful creation
                    # if new_component:
                    #      export_step_file(new_component)
            # If in CSV mode and no keys were loaded, this block is entered,
            # but since _batch_generate_triggered is False, no generation happens,
            # and the command will simply close when the user clicks OK.


    def on_destroy(self, command, inputs, reason, input_values):
        """Cleans up resources when the command is destroyed."""
        # clear_graphics(self.graphics_group) # Clear graphics if used for preview
        # self.graphics_group.deleteMe() # Delete graphics group if created
        pass

    def on_create(self, command, inputs):
        """Creates the command inputs when the command is activated."""
        ao = apper.AppObjects()
        # self.graphics_group = ao.root_comp.customGraphicsGroups.add() # Create graphics group if needed
        self.make_preview = True # Ensure preview is triggered initially
        self.extracted_keys = [] # Ensure keys are cleared on command creation
        # Reintroduced the flag and initialize to False
        self._batch_generate_triggered = False

        selection_input = inputs.addSelectionInput('sketch_point', "Center Point", "Pick Sketch Point for center")
        selection_input.addSelectionFilter("SketchPoints")

        # Added qr_size input
        inputs.addValueInput('qr_size', 'QR Code Size (mm)', 'mm', adsk.core.ValueInput.createByString(QR_TARGET_SIZE_MM))
        # Inputs for block_height and base_height removed as requested.
        # They will now use the default values defined at the top of the script.

        # Renamed the group input title here
        group_input = inputs.addGroupCommandInput('group', 'CSV Import')

        # Corrected: Ensure the correct inputs are added based on is_make_qr
        if self.is_make_qr:
            add_make_inputs(group_input.children)
        else:
            add_csv_inputs(group_input.children)
