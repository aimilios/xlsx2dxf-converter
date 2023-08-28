
# xlsx2dxf-converter: Electrical Symbol Block Generator

Converts xlsx-based electrical symbol description into DXF blocks represented by rectangles.

# Example-1
Input Data:


| Port Name    | Position       | Position Index|
|--------------|----------------|----------------|
| 1:5V             | left |0|
| 2:5V             | left |1|
| 4:Rx             | left |2|
| 5:Tx             | left |3|
| 8:GND            | left |4|
| RS485+       | right          |0|
| RS485-       | right          |1|

| Block Label  | Block Model  |Block Comment|
|--------------|----------------|----------------|
| PLC1          |   Delta PLC  |Main PLC|

Output File:  

![App Screenshot](https://via.placeholder.com/468x300?text=App+Screenshot+Here)

## Usage/Examples

To use the xlsx2dxf_converter, follow these steps:

1. **Prerequisites**: Ensure you have Python installed on your system and also the following Python modules installed:

- **ezdxf** (version '0.18.1')
- **openpyxl** (version '3.0.10')

2. **Setup**:
   - Create a folder named `Blocks` in the same directory as the script.
   - Place the XLSX block description files inside the `Blocks` folder. Feel free to use the examples already provided in the folder as references. Alternatively, you can utilize the "block_template.xlsx" file by copying it into the "Blocks" folder.


3. **Run From Terminal**:
To run from the terminal, navigate to the script's directory an run the command:

   ```
   python convert_dxf.py "<file_path>" "<block_label>" "<block_comment>" "<group_name>" <dxf_filename>
   ```
* file_path: The xlsx filename 
* block_label: A Unique word for every part in the circuit.
* block_comment: A comment describing the part.
* group_name: The group name of the xlsx file.
* dxf_filename: Output dxf file name.

Example-1:
   ```
    python xlsx2dxf_converter.py "Blocks\Teco L510.xlsx" "power" "VFD23" "Teco L510\n210-SH1F-P 0.75kW" "Main\nInverter\nRollers" "output.dxf"
   ```
Example-2:
   ```   
   python xlsx2dxf_converter.py "Blocks\Delta DVP12SA2.xlsx" "communication_all" "PLC1" "Delta PLC\nDVP-DVP-SA2" "Main\nPLC" "output2.dxf"
   ```
4. **Run Using the GUI**:
