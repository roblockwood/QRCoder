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
# Import necessary libraries
import csv
import os.path
import traceback # Added for detailed error reporting

import adsk.core
import adsk.fusion
import adsk.cam

# Import Apper framework and configuration
from ..apper import apper
from .. import config

# --- Default Values ---
HEIGHT = '.4 mm'
BASE = '0.0 in'
MESSAGE = 'input here'
QR_TARGET_SIZE_MM = '24 mm'  # Target QR code size in millimeters
FILE_NAME = 'QR-17x.csv' # Default/Example CSV filename (if needed)
BLOCK = '.03 in' # Default block size (Note: This seems unused as block size is calculated later)

# --- Helper Functions ---

def get_target_body(sketch_point: adsk.fusion.SketchPoint) -> adsk.fusion.BRepBody | None:
    """
    Finds the BRepBody that contains the given sketch point.

    Args:
        sketch_point: The sketch point to search from.

    Returns:
        The BRepBody containing the point, or None if not found.
    """
    ao = apper.AppObjects()
    # Search for bodies near the sketch point's world geometry
    target_collection = ao.root_comp.findBRepUsingPoint(
        sketch_point.worldGeometry,
        adsk.fusion.BRepEntityTypes.BRepBodyEntityType,
        -1.0, # Use a negative tolerance for robust checking
        True  # Search through occurrences
    )

    if target_collection.count > 0:
        return target_collection.item(0)
    else:
        # Inform user if no body is found at the point
        ao.ui.messageBox(f"Could not find a target body associated with the selected sketch point: {sketch_point.name}")
        return None

def make_real_geometry(target_body: adsk.fusion.BRepBody, input_values: dict, qr_data: list) -> adsk.fusion.Component | None:
    """
    Creates the QR code geometry within a new component.

    Args:
        target_body: The target body (used to determine the parent component for copying).
        input_values: Input values from the command dialog.
        qr_data: The QR code data (list of lists, e.g., [[1,0],[0,1]]).

    Returns:
        The newly created component containing the QR code geometry, or None on failure.
    """
    ao = apper.AppObjects()
    design = ao.design
    root_comp = design.rootComponent

    if not target_body:
        # This check might be redundant if get_target_body handles it, but good for safety.
        ao.ui.messageBox("make_real_geometry called without a valid target body.")
        return None

    # Get the component where the target body resides
    parent_comp = target_body.parentComponent

    # --- Create a new component ---
    # An identity matrix means the new component is placed at the origin of the root
    transform = adsk.core.Matrix3D.create()
    # IMPORTANT: addNewComponent creates the component AND adds an occurrence to the specified parent (root_comp here)
    new_occurrence = root_comp.occurrences.addNewComponent(transform)
    if not new_occurrence:
        ao.ui.messageBox("Failed to create a new occurrence/component.")
        return None
    new_comp = new_occurrence.component
    if not new_comp:
        # Should not happen if occurrence creation succeeded, but check anyway
        ao.ui.messageBox("Failed to get the new component from the occurrence.")
        return None

    # Rename the new component based on the input message
    try:
        new_comp.name = input_values.get('message', 'QR_Component') # Use get for safety
    except Exception as e:
        # Handle potential errors if the name is invalid (e.g., too long, invalid characters)
        ao.ui.messageBox(f"Warning: Could not set component name to '{input_values.get('message', '')}': {e}\nUsing default name.")
        new_comp.name = "QR_Component" # Fallback name

    # --- Copy bodies from the original component ---
    # NOTE: This copies ALL solid bodies from the component containing the target_body.
    # If you only want to copy the target_body itself, modify this section.
    try:
        bodies_to_copy = parent_comp.bRepBodies
        if bodies_to_copy.count == 0:
             ao.ui.messageBox(f"Warning: The parent component '{parent_comp.name}' of the target body contains no bodies to copy.")
        for body in bodies_to_copy:
            if body.isSolid:
                # copyToComponent copies the body geometry into the new component via its occurrence
                copied_body = body.copyToComponent(new_occurrence)
                if not copied_body:
                    ao.ui.messageBox(f"Warning: Failed to copy body '{body.name}' to the new component.")
    except Exception as e:
        ao.ui.messageBox(f"Error copying bodies from '{parent_comp.name}' to '{new_comp.name}': {e}\n{traceback.format_exc()}")
        # Decide if this is a fatal error or if we can continue with just the QR code
        # For now, let's continue, but log the error.

    try:
        # --- Create the QR code temporary geometry ---
        # This uses the TemporaryBRepManager, which is efficient but doesn't add to the timeline
        temp_body = get_qr_temp_geometry(qr_data, input_values)
        if not temp_body:
             ao.ui.messageBox("Failed to generate temporary QR code geometry.")
             # Clean up the potentially empty new component? Or leave it? For now, leave it but return None.
             return None

        # --- Add the QR code geometry as a Base Feature ---
        # Base Features are non-parametric, suitable for imported/temporary geometry
        base_feature = new_comp.features.baseFeatures.add()
        if not base_feature:
             ao.ui.messageBox("Failed to create a Base Feature in the new component.")
             return None

        base_feature.startEdit()
        # Add the temporary body geometry to the Base Feature
        # Ensure temp_body is valid before adding
        if temp_body.isValid:
            added_body = new_comp.bRepBodies.add(temp_body, base_feature)
            if not added_body:
                 ao.ui.messageBox("Failed to add the QR code body to the Base Feature.")
                 base_feature.cancelEdit() # Cancel the edit if adding failed
                 return None
        else:
            ao.ui.messageBox("Temporary QR code body is invalid before adding to Base Feature.")
            base_feature.cancelEdit()
            return None


        base_feature.finishEdit()

        # --- Rename the newly added QR code body ---
        # Find the body added by the base feature (it might not be named "Solid1")
        qr_body_found = False
        # Check added_body directly if it's valid
        if added_body and added_body.isValid:
             try:
                 added_body.name = "QR_Code_Geometry"
                 qr_body_found = True
             except Exception as name_e:
                  ao.ui.messageBox(f"Warning: Could not rename the QR code body: {name_e}")
                  qr_body_found = True # Still found it, even if renaming failed
        else:
            # Fallback: Iterate through bodies if direct reference failed or is invalid
            for body in new_comp.bRepBodies:
                 # Check if the body belongs to the base_feature
                 if body.parentBaseFeature == base_feature:
                     try:
                         body.name = "QR_Code_Geometry"
                         qr_body_found = True
                         break # Found and renamed
                     except Exception as name_e:
                          ao.ui.messageBox(f"Warning: Could not rename the QR code body (fallback method): {name_e}")
                          qr_body_found = True # Still found it, even if renaming failed
                          break

        if not qr_body_found:
             ao.ui.messageBox("Warning: Could not identify the newly added QR code body to rename it.")


        # --- Success ---
        return new_comp # Return the populated component

    except Exception as e:
        # Catch any unexpected errors during geometry creation
        ao.ui.messageBox(f"Error during QR code geometry creation or addition:\n{e}\n{traceback.format_exc()}")
        # Clean up? Abort transaction? (Transaction handled in on_execute)
        return None # Return None on error

def clear_graphics(graphics_group: adsk.fusion.CustomGraphicsGroup):
    """Removes all entities from a custom graphics group."""
    # Iterate backwards safely while deleting
    for i in range(graphics_group.count - 1, -1, -1):
        entity = graphics_group.item(i)
        if entity and entity.isValid:
            entity.deleteMe()

def make_graphics(t_body: adsk.fusion.BRepBody, graphics_group: adsk.fusion.CustomGraphicsGroup):
    """Displays a temporary body using custom graphics."""
    clear_graphics(graphics_group)
    # Define a color (e.g., orange)
    color = adsk.core.Color.create(250, 162, 27, 255) # RGBA
    color_effect = adsk.fusion.CustomGraphicsSolidColorEffect.create(color)
    # Add the body geometry to the graphics group
    graphics_body = graphics_group.addBRepBody(t_body)
    graphics_body.color = color_effect
    # Optional: Set transparency, visibility, etc.
    # graphics_body.opacity = 0.8


# CORRECTED TYPE HINT HERE: Changed TemporaryBRepBody to BRepBody
def get_qr_temp_geometry(qr_data: list, input_values: dict) -> adsk.fusion.BRepBody | None:
    """
    Generates the temporary BRep geometry for the QR code blocks.

    Args:
        qr_data: The QR code data (list of lists).
        input_values: Input values from the command dialog.

    Returns:
        A BRepBody representing the temporary QR code geometry, or None on failure.
    """
    ao = apper.AppObjects()
    try:
        # --- Get Input Values ---
        # Use .get() with defaults for robustness
        # Note: input_values contains the *values*, not the input objects themselves
        height_val = input_values.get('block_height')
        base_val = input_values.get('base_height')
        qr_size_val = input_values.get('qr_size')
        sketch_point_list = input_values.get('sketch_point', [])

        # Validate inputs - basic check if values are present (more detailed check in on_preview)
        if not all([height_val is not None, base_val is not None, qr_size_val is not None]):
             ao.ui.messageBox("Missing required value inputs (height, base, or qr_size) in get_qr_temp_geometry.")
             return None
        if not sketch_point_list:
             ao.ui.messageBox("Sketch point selection is missing in get_qr_temp_geometry.")
             return None

        # Access the float values directly from input_values
        height: float = height_val # input_values already contains the float value
        base: float = base_val
        overall_size: float = qr_size_val
        sketch_point: adsk.fusion.SketchPoint = sketch_point_list[0]

        # --- Calculate Geometry Parameters ---
        qr_matrix_size = len(qr_data)
        if qr_matrix_size == 0:
            ao.ui.messageBox("QR data is empty.")
            return None

        # Calculate the size of each individual square block (in cm)
        side = overall_size / qr_matrix_size
        # Store calculated block size back into input_values if needed elsewhere (optional)
        # input_values['calculated_block_size_cm'] = side

        # --- Define Coordinate System based on Sketch ---
        sketch = sketch_point.parentSketch
        if not sketch:
             ao.ui.messageBox("Could not get parent sketch from the selected point.")
             return None

        x_dir = sketch.xDirection
        y_dir = sketch.yDirection
        # Ensure vectors are normalized (unit vectors)
        x_dir.normalize()
        y_dir.normalize()
        # Calculate Z direction (normal to the sketch plane)
        z_dir = x_dir.crossProduct(y_dir)
        z_dir.normalize()

        # --- Calculate Start Point for QR Grid ---
        # Center point in world coordinates
        center_point_world = sketch_point.worldGeometry
        start_point = center_point_world.copy()

        # Move from center to the corner of the first (top-left) block's center for OBB creation
        x_start_offset = x_dir.copy()
        x_start_offset.scaleBy(-side * (qr_matrix_size / 2.0) + (side / 2.0)) # Center X of the first block
        start_point.translateBy(x_start_offset)

        y_start_offset = y_dir.copy()
        y_start_offset.scaleBy(side * (qr_matrix_size / 2.0) - (side / 2.0)) # Center Y of the first block
        start_point.translateBy(y_start_offset)

        # Move point to the center of the block's height dimension for box creation
        z_start_offset = z_dir.copy()
        z_start_offset.scaleBy(base + (height / 2.0)) # Center of the block height, on top of the base
        start_point.translateBy(z_start_offset)

        # --- Calculate Base Geometry (if applicable) ---
        b_mgr = adsk.fusion.TemporaryBRepManager.get()
        final_qr_body = None # Initialize variable to hold the combined QR geometry

        if base > 1e-6: # Use tolerance for float comparison
            base_center_point = center_point_world.copy()
            z_base_offset = z_dir.copy()
            z_base_offset.scaleBy(base / 2.0) # Center of the base height
            base_center_point.translateBy(z_base_offset)

            # Create the base plate box
            base_obb = adsk.core.OrientedBoundingBox3D.create(base_center_point, x_dir, y_dir, overall_size, overall_size, base)
            final_qr_body = b_mgr.createBox(base_obb) # Start with the base
            if not final_qr_body:
                ao.ui.messageBox("Failed to create temporary base geometry.")
                # Decide how to proceed - maybe return None? For now, continue without base.
                final_qr_body = None # Reset if creation failed
        # else: No base needed

        # --- Create QR Blocks ---
        for i, row in enumerate(qr_data):
            for j, col_char in enumerate(row):
                # Check if the cell should be solid (1)
                if str(col_char).strip() == '1':
                    # Calculate center point for this block
                    block_center_point = start_point.copy()

                    x_block_offset = x_dir.copy()
                    x_block_offset.scaleBy(j * side)
                    block_center_point.translateBy(x_block_offset)

                    y_block_offset = y_dir.copy()
                    y_block_offset.scaleBy(-i * side) # Negative because sketch Y is often upwards, grid rows go down
                    block_center_point.translateBy(y_block_offset)

                    # Create the Oriented Bounding Box for the block
                    block_obb = adsk.core.OrientedBoundingBox3D.create(block_center_point, x_dir, y_dir, side, side, height)
                    # Create the temporary body for the block
                    block_body = b_mgr.createBox(block_obb)

                    if not block_body:
                         ao.ui.messageBox(f"Warning: Failed to create temporary geometry for block at row {i}, col {j}.")
                         continue # Skip this block

                    # --- Combine with existing geometry ---
                    if final_qr_body is None:
                        # This is the first block (and there's no base, or base creation failed)
                        final_qr_body = block_body
                    else:
                        # Union this block with the main body
                        # Perform boolean operation (Union)
                        op_success = b_mgr.booleanOperation(final_qr_body, block_body, adsk.fusion.BooleanTypes.UnionBooleanType)
                        if not op_success:
                            ao.ui.messageBox(f"Warning: Boolean union failed for block at row {i}, col {j}.")
                            # If union fails, the final_qr_body might be compromised.
                            # It might be safer to return None here or try to recover.
                            # For now, we just warn and continue. The final body might be incomplete.

        # Check if any geometry was created at all
        if final_qr_body is None:
            ao.ui.messageBox("No QR code geometry was generated (possibly empty QR data or all blocks failed).")
            return None

        return final_qr_body # Return the combined temporary body

    except Exception as e:
        ao.ui.messageBox(f"Error in get_qr_temp_geometry:\n{e}\n{traceback.format_exc()}")
        return None


def import_qr_from_file(file_name: str) -> list:
    """
    Reads QR data (0s and 1s) from a CSV file.
    Assumes each cell in the CSV represents a QR block state.

    Args:
        file_name: Path to the CSV file.

    Returns:
        A list of lists representing the QR data, or empty list on error.
    """
    qr_data = []
    ao = apper.AppObjects()
    try:
        if os.path.exists(file_name):
            with open(file_name, mode='r', newline='') as f:
                reader = csv.reader(f)
                qr_data = list(reader) # Reads all rows into a list of lists
            if not qr_data:
                 ao.ui.messageBox(f"Warning: CSV file '{file_name}' appears to be empty.")
        else:
             ao.ui.messageBox(f"Error: CSV file not found at '{file_name}'.")
    except Exception as e:
        ao.ui.messageBox(f"Error reading CSV file '{file_name}': {e}\n{traceback.format_exc()}")
        return [] # Return empty list on error

    return qr_data


@apper.lib_import(config.lib_path) # Decorator to handle importing pyqrcode
def build_qr_code(message: str, args: dict) -> list:
    """
    Generates QR code data using the pyqrcode library.

    Args:
        message: The string to encode.
        args: Dictionary of arguments for pyqrcode.create() (e.g., error correction level).

    Returns:
        A list of lists representing the QR data (0s and 1s), or empty list on error.
    """
    ao = apper.AppObjects()
    try:
        # Dynamically import pyqrcode - handled by @apper.lib_import
        import pyqrcode

        # Create the QR code object
        qr = pyqrcode.create(message, **args)

        # Get the text representation (matrix of 0s and 1s)
        # quiet_zone=0 removes the white border around the QR code data area
        qr_text = qr.text(quiet_zone=0)

        # Convert the text matrix into a list of lists of integers
        qr_data = []
        for line in qr_text.splitlines():
            stripped_line = line.strip()
            if stripped_line: # Ensure line is not empty
                 # Convert each character ('0' or '1') to an integer
                 try:
                     row = [int(char) for char in stripped_line]
                     qr_data.append(row)
                 except ValueError:
                     ao.ui.messageBox(f"Warning: Non-numeric character found in QR text line: '{stripped_line}'. Skipping line.")
                     continue # Skip lines with invalid characters

        if not qr_data:
             ao.ui.messageBox("Warning: No valid QR data rows generated from pyqrcode text.")
             return []

        return qr_data

    except ImportError:
        ao.ui.messageBox(f"Failed to import 'pyqrcode'. Ensure it's installed in:\n{config.lib_path}")
        return []
    except ValueError as e:
        # pyqrcode raises ValueError for issues like message too long for version/error level
        ao.ui.messageBox(f'Error generating QR code (pyqrcode): {e}')
        return []
    except Exception as e:
        # Catch any other unexpected errors from pyqrcode
        ao.ui.messageBox(f'An unexpected error occurred during QR code generation:\n{e}\n{traceback.format_exc()}')
        return []


def make_qr_from_message(input_values: dict) -> list:
    """
    Wrapper function to generate QR data from a message string in input_values.
    """
    message: str = input_values.get('message', '') # Get message safely
    if not message:
        ao = apper.AppObjects()
        ao.ui.messageBox("No message provided to encode into QR code.")
        return []

    # Define arguments for pyqrcode (e.g., error correction level)
    # L = Low (7%), M = Medium (15%), Q = Quartile (25%), H = High (30%)
    qr_args = {'error': 'M'} # Example: Medium error correction

    # Check if pyqrcode dependency is met (optional, as build_qr_code handles ImportError)
    # success = apper.check_dependency('pyqrcode', config.lib_path)
    # if not success:
    #     ao = apper.AppObjects()
    #     ao.ui.messageBox(f"Dependency 'pyqrcode' not found in {config.lib_path}")
    #     return []

    # Call the function that uses pyqrcode
    qr_data = build_qr_code(message, qr_args)
    return qr_data


# --- Command Input Creation Functions ---

def add_make_inputs(inputs: adsk.core.CommandInputs):
    """Adds inputs specific to creating a single QR code from text."""
    # Input for the text/data to encode
    inputs.addStringValueInput('message', 'Value to Encode', MESSAGE)
    # Add other options if needed (e.g., error correction level dropdown)

def add_csv_inputs(inputs: adsk.core.CommandInputs):
    """Adds inputs specific to creating QR codes from a CSV file."""
    # Input for the CSV file path
    file_input = inputs.addStringValueInput('file_name', "CSV File", FILE_NAME) # Provide default
    file_input.isReadOnly = True # Make it read-only, populated by browse button

    # Button to trigger file browser
    browse_button = inputs.addBoolValueInput('browse', 'Browse...', False, '', True) # Set resourceFolder to ""
    browse_button.isFullWidth = True

    # Read-only text box for CSV preview
    inputs.addTextBoxCommandInput('csv_preview', 'CSV Preview', 'Click Browse... to load file.', 5, True) # Rows, ReadOnly


# --- File Browser Logic ---

def browse_for_csv() -> str:
    """Opens a file dialog for selecting a CSV file."""
    ao = apper.AppObjects()
    file_dialog = ao.ui.createFileDialog()
    # Set dialog properties
    file_dialog.initialDirectory = config.app_path # Start in the add-in's directory
    file_dialog.filter = "CSV Files (*.csv);;All Files (*.*)"
    file_dialog.isMultiSelectEnabled = False
    file_dialog.title = 'Select CSV File for QR Codes'

    # Show the dialog
    dialog_results = file_dialog.showOpen()

    # Process results
    if dialog_results == adsk.core.DialogResults.DialogOK:
        return file_dialog.filename # Return the selected file path
    else:
        return '' # Return empty string if canceled


# --- STEP Export Function --- (Optional, can be called after execution)
def export_step_file(component: adsk.fusion.Component):
    """Exports the given component as a STEP file."""
    ao = apper.AppObjects()
    try:
        if not component or not component.isValid:
             ao.ui.messageBox("Cannot export invalid component.")
             return

        export_manager = ao.design.exportManager
        # Create STEP export options
        step_options = export_manager.createSTEPExportOptions() # Correct method name

        # Define filename and path
        # Consider using a file dialog for the save location
        default_folder = os.path.expanduser("~") # Default to user's home directory
        safe_comp_name = "".join(c if c.isalnum() else "_" for c in component.name) # Sanitize name
        file_name = os.path.join(default_folder, f"{safe_comp_name}.step")

        # Use the component's occurrence for export if available and valid, otherwise use the component itself
        export_target = component
        if component.occurrences and component.occurrences.count > 0:
             # Find a valid occurrence to export (usually the first one if it's the only one)
             valid_occurrence = None
             for occ in component.occurrences:
                 if occ.isValid:
                     valid_occurrence = occ
                     break
             if valid_occurrence:
                 step_options.fileName = file_name # Use fileName (lowercase N)
                 # Exporting an occurrence might not be directly supported this way.
                 # Often, you export the component and it contains the geometry defined within its context.
                 # Let's stick to exporting the component.
                 step_options.component = component

             else:
                  ao.ui.messageBox(f"Could not find a valid occurrence for component '{component.name}' to export.")
                  return # Cannot export without a valid occurrence context usually
        else:
             # Exporting a component directly (might be the root or uninstantiated)
             step_options.fileName = file_name
             step_options.component = component


        # Execute the export
        result = export_manager.execute(step_options)
        if result:
             ao.ui.messageBox(f'Successfully exported STEP file to:\n{file_name}')
        else:
             ao.ui.messageBox(f'STEP export failed for component: {component.name}')

    except Exception as e:
        ao.ui.messageBox(f'Failed to export STEP file for {component.name}:\n{e}\n{traceback.format_exc()}')


# --- Main Command Class ---

class QRCodeMaker(apper.Fusion360CommandBase):
    def __init__(self, name: str, options: dict):
        super().__init__(name, options)
        # Flag to control expensive preview generation
        self.make_preview = True
        # Flag to determine command mode (single QR vs CSV batch)
        self.is_make_qr = options.get('is_make_qr', False) # Default to False if not specified

    def on_input_changed(self, command: adsk.core.Command, inputs: adsk.core.CommandInputs,
                         changed_input: adsk.core.CommandInput, input_values: dict):
        """Handles changes to command inputs."""
        # Reset preview flag on any input change
        # Only reset if the changed input is relevant to the geometry
        relevant_inputs = ['sketch_point', 'qr_size', 'block_height', 'base_height', 'message', 'file_name']
        if changed_input.id in relevant_inputs:
            self.make_preview = True

        # Handle the 'Browse...' button click
        if changed_input.id == 'browse':
            # Set the boolean value back to False (it's just a trigger)
            browse_button = changed_input # Keep reference
            browse_button.value = False

            # Open the file dialog
            file_name = browse_for_csv()

            # If a file was selected, update the file name input and preview
            if file_name:
                file_name_input = inputs.itemById('file_name')
                if file_name_input:
                    file_name_input.value = file_name
                    self.make_preview = True # Need to revalidate if file changes
                # Update the preview text box
                preview_text_box = inputs.itemById('csv_preview')
                if preview_text_box:
                    preview_text = self._get_csv_preview(file_name)
                    preview_text_box.text = preview_text
            # else: User cancelled the dialog, do nothing

        # Add logic here if other inputs need to dynamically change the dialog
        # For example, enabling/disabling inputs based on selections.

    def _get_csv_preview(self, file_name: str) -> str:
        """Reads the CSV file and generates a preview string."""
        preview_lines = []
        max_preview_lines = 10
        ao = apper.AppObjects()
        try:
            with open(file_name, mode='r', newline='', encoding='utf-8-sig') as f: # Handle potential BOM
                reader = csv.DictReader(f)
                # Check if the required header 'KEY' exists
                if not reader.fieldnames:
                    return "CSV file appears to be empty or has no header row."
                # Case-insensitive check for 'KEY'
                key_header = None
                for field in reader.fieldnames:
                    if field.upper() == 'KEY':
                        key_header = field
                        break

                if not key_header:
                    return f"CSV file is missing the required header: 'KEY'.\nFound headers: {', '.join(reader.fieldnames)}"

                preview_lines.append(f"Found '{key_header}' column. Preview:")
                count = 0
                total_items = 0
                for row in reader:
                    total_items += 1
                    if count < max_preview_lines:
                        # Safely get the KEY value using the found header name
                        key_value = row.get(key_header, '*MISSING*')
                        preview_lines.append(f"- {key_value}")
                        count += 1

                if total_items == 0:
                    return f"CSV file has '{key_header}' header but no data rows."
                if total_items > max_preview_lines:
                    preview_lines.append(f"... (Total {total_items} items)")

                return "\n".join(preview_lines)

        except FileNotFoundError:
            return f"Error: File not found.\n'{file_name}'"
        except Exception as e:
            # Provide more specific error feedback if possible
            return f"Error reading CSV:\n{e}\nCheck file encoding and format."


    def on_preview(self, command: adsk.core.Command, inputs: adsk.core.CommandInputs,
                   args: adsk.core.CommandEventArgs, input_values: dict):
        """Generates a preview in the graphics window (optional)."""
        # Currently, preview only sets validity. No graphics generated here.
        # Consider adding custom graphics preview if needed, but be mindful of performance.

        # ALWAYS set isValidResult initially. It will be set to False if validation fails.
        args.isValidResult = True

        # Only recalculate preview state if relevant inputs changed or it's the first preview
        # Or if the command was just started (make_preview is True initially)
        if self.make_preview or not args.isValidResult: # Revalidate if previous check failed
            ao = apper.AppObjects()

            # --- Validate Required Inputs ---

            # Check if sketch point is selected (required for both modes)
            sketch_point_input = inputs.itemById('sketch_point') # Get the actual input object
            if not sketch_point_input or not sketch_point_input.isValid or sketch_point_input.selectionCount == 0:
                 args.isValidResult = False
                 # ao.ui.messageBox("Sketch point selection is required.") # Avoid message boxes in preview

            # Validate mode-specific inputs
            if self.is_make_qr:
                # Check if message is provided
                message_input = inputs.itemById('message') # Get the actual input object
                if not message_input or not message_input.isValid or not message_input.value:
                    args.isValidResult = False # Cannot execute without message
                    # ao.ui.messageBox("Message to encode is required.")

            else: # CSV mode
                 # Check if CSV file is selected and exists
                 file_name_input = inputs.itemById('file_name') # Get the actual input object
                 if not file_name_input or not file_name_input.isValid or not file_name_input.value or not os.path.exists(file_name_input.value):
                    args.isValidResult = False # Cannot execute without valid file
                    # ao.ui.messageBox("A valid CSV file must be selected.")
                 # Optionally, check if CSV has 'KEY' header here for better preview validation
                 # This might be too slow for on_preview, better in on_execute or after browse
                 # elif not self._check_csv_header(file_name_input.value):
                 #     args.isValidResult = False
                     # ao.ui.messageBox("CSV file is missing the required 'KEY' header.")


            # Check if dimension inputs are valid numbers and positive
            for dim_id in ['qr_size', 'block_height']:
                 val_input = inputs.itemById(dim_id) # *** Get the actual ValueInput object here ***
                 # Check if the input exists, is valid, and its value is positive
                 if not val_input or not val_input.isValid or val_input.value <= 0: # Now val_input IS a ValueInput
                      args.isValidResult = False
                      # ao.ui.messageBox(f"'{val_input.name}' must be a positive value.")
                      break # One invalid dimension is enough

            # Check base_height separately as it can be 0
            base_input = inputs.itemById('base_height') # *** Get the actual ValueInput object here ***
            if not base_input or not base_input.isValid or base_input.value < 0: # Base cannot be negative
                 args.isValidResult = False
                 # ao.ui.messageBox(f"'{base_input.name}' cannot be negative.")


            # If all checks pass up to this point, consider the preview valid for now
            # The actual geometry generation is in make_real_geometry/get_qr_temp_geometry
            # which will also return None on failure.
            # This validity check in on_preview primarily enables/disables the OK button.


            # Reset the flag so preview doesn't regenerate unless an input changes again
            # Only set to False if validity check passed, otherwise keep True to recheck next time
            if args.isValidResult:
                self.make_preview = False

        # Note: If you were generating custom graphics for preview, you would do it here
        # based on input_values and the validity check.
        # Example:
        # if args.isValidResult and self.make_preview: # Only if inputs are valid and preview flag is set
        #    qr_data = self._get_qr_data_for_preview(input_values) # Get QR data based on mode
        #    if qr_data:
        #        temp_body = get_qr_temp_geometry(qr_data, input_values) # Generate temp geometry
        #        if temp_body:
        #             # Display temp_body using custom graphics (requires setup)
        #             pass
        #    self.make_preview = False # Reset flag after generating graphics


    def on_execute(self, command: adsk.core.Command, inputs: adsk.core.CommandInputs,
                   args: adsk.core.CommandEventArgs, input_values: dict):
        """Executes when the user clicks OK."""
        ao = apper.AppObjects()
        design = ao.design
        root_comp = design.rootComponent

        # Ensure a transaction is active for operations that modify the design
        command.commandInputs.itemById('okButton').isEnabled = False # Disable OK button during execution
        adsk.doEvents() # Process events to update UI

        try:
            # --- Get Input Values ---
            # input_values dictionary already contains the values from the inputs
            # We validated these in on_preview, but a quick check here is good practice.
            sketch_point_list = input_values.get('sketch_point', [])
            if not sketch_point_list:
                 ao.ui.messageBox("Execution failed: Sketch point selection is missing.")
                 return

            sketch_point = sketch_point_list[0]

            # Get QR data based on the mode
            qr_data = []
            if self.is_make_qr:
                qr_data = make_qr_from_message(input_values)
            else: # CSV mode
                file_name = input_values.get('file_name', '')
                if not file_name or not os.path.exists(file_name):
                     ao.ui.messageBox("Execution failed: Invalid CSV file selected.")
                     return
                # For CSV mode, we would typically iterate through rows and create multiple QR codes.
                # This example currently only handles a single QR from message.
                # You would need to modify this section to read the CSV and loop,
                # calling get_qr_temp_geometry and make_real_geometry for each row.
                ao.ui.messageBox("CSV batch processing is not fully implemented in this example.")
                return # Exit for now as CSV mode isn't fully implemented

            if not qr_data:
                 ao.ui.messageBox("Execution failed: No QR data generated.")
                 return

            # --- Get Target Body ---
            # This is needed to determine the parent component for the new QR component
            target_body = get_target_body(sketch_point)
            if not target_body:
                 # get_target_body already shows a message box
                 return # Cannot proceed without a target body

            # --- Create the QR code geometry in a new component ---
            # Pass the target_body to determine the parent component context
            new_component = make_real_geometry(target_body, input_values, qr_data)

            if new_component:
                 ao.ui.messageBox(f"Successfully created QR code component: {new_component.name}")
                 # Optional: Export the component after creation
                 # export_step_file(new_component)
            else:
                 ao.ui.messageBox("Failed to create the QR code component.")

        except Exception as e:
            # Catch any exception that occurs during execution
            ao.ui.messageBox(f'Execution failed:\n{e}\n{traceback.format_exc()}')

        finally:
            # This block always executes, even if an exception occurred
            command.commandInputs.itemById('okButton').isEnabled = True # Re-enable OK button


    def on_create(self, command: adsk.core.Command, inputs: adsk.core.CommandInputs):
        """Creates the command inputs."""
        ao = apper.AppObjects()

        # Add common inputs
        # Use a SelectionCommandInput for the sketch point
        selection_input = inputs.addSelectionInput('sketch_point', 'Sketch Point', 'Select a sketch point')
        selection_input.addSelectionFilter('SketchPoints') # Filter to only allow sketch points
        selection_input.setSelectionLimits(1, 1) # Allow exactly one selection

        # Add dimension inputs (ValueInputs)
        # Use addValueInput for dimensions
        # Specify the unit type (e.g., 'mm', 'in')
        inputs.addValueInput('qr_size', 'QR Size', ao.units_manager.defaultLengthUnits, adsk.core.ValueInput.createByString(QR_TARGET_SIZE_MM))
        inputs.addValueInput('block_height', 'Block Height', ao.units_manager.defaultLengthUnits, adsk.core.ValueInput.createByString(HEIGHT))
        inputs.addValueInput('base_height', 'Base Height', ao.units_manager.defaultLengthUnits, adsk.core.ValueInput.createByString(BASE)) # Base can be 0

        # Add inputs based on the command mode
        if self.is_make_qr:
            add_make_inputs(inputs)
        else: # CSV mode
            add_csv_inputs(inputs)

        # Add OK and Cancel buttons (apper handles this by default, but you can add custom ones if needed)
        # inputs.addCommandInput('okButton', adsk.core.CommandInputTypes.ButtonCommandInput, 'OK')
        # inputs.addCommandInput('cancelButton', adsk.core.CommandInputTypes.ButtonCommandInput, 'Cancel')

    # Optional: Implement on_destroy if you need to clean up resources
    # def on_destroy(self, command: adsk.core.Command, inputs: adsk.core.CommandInputs, reason: int):
    #     pass

