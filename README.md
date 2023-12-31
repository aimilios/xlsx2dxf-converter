
# xlsx2dxf-converter: Electrical Symbol Block Generator

Converts xlsx-based electrical symbol description into DXF block represented by rectangles.

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
| PLC1          |   Delta PLC <br> DVP-DVP-SA2  |Main <br> PLC|

Output File:

![output2](https://github.com/aimilios/xlsx2dxf-converter/assets/7573375/cea1aa15-9783-4fdf-a38c-4e416ae01098)

# Example-2
Input Data:

| Port Name    | Position       | Position Index|
|--------------|----------------|----------------|
| L1(L)             | left |0|
| L3(N)            | left |1|
| T1             | right |0|
| T2             | right |1|
| T3            | right |2|
| GND       | right          |3|

| Block Label  | Block Model  |Block Comment|
|--------------|----------------|----------------|
| VFD23          |   Teco L510 <br> 210-SH1F-P 0.75kW  |Main <br> Inverter <br> Rollers|

Output File:

![output1](https://github.com/aimilios/xlsx2dxf-converter/assets/7573375/4ca67ad9-fa87-4906-8ba4-77081c7e6960)


## Script Objective
When designing an electrical block,which is often a one-time task,you can later reuse the block across various circuits. Yet when you want to insert the block some text properties must be edited such as the Block Label,Model or Comment. Conventionally, the solution is to duplicate the block, make edits, and then create a new block.

To address this challenge, the `xlsx2dxf_converter.py` script comes into play, make it easier to modify and reuse blocks. Additionally, a more user-friendly alternative is provided through the `xlsx2dxf_gui.py` script, which offers a GUI interface using the core functionality of `xlsx2dxf_converter.py`.

In many cases, an electrical block can have multiple representations based on the specific circuit type. For instance, in a power circuit, you might only need power-related ports, while in a communication circuit, you may focus solely on communication ports. To change between various represantation you use the `Group Name` property within the associated XLSX file. There is also the `Block Category` optional property which makes it easier to find the target block using the GUI.

For example:

![groups explanation-2](https://github.com/aimilios/xlsx2dxf-converter/assets/7573375/b378f47f-7bfe-4a5a-8762-f868916f0937)




## Usage

To use the xlsx2dxf_converter, follow these steps:

1. **Prerequisites**: Ensure you have Python installed on your system and also the following Python modules installed:

- **ezdxf**
- **openpyxl**
- **PyQt5** (only required if you intend to use the GUI script)

2. **Setup**:
   - Create a folder named `Blocks` in the same directory as the script.
   - Place the XLSX block description files inside the `Blocks` folder. Feel free to use the examples already provided in the folder as references. Alternatively, you can utilize the "block_template.xlsx" file by copying it into the "Blocks" folder.


3. **Run From Terminal**:
To run from the terminal, navigate to the script's directory an run the command:

   ```bash
   python convert_dxf.py "<file_path>" "<block_label>" "<block_comment>" "<group_name>" <dxf_filename>
   ```
* file_path: The xlsx filename 
* block_label: A Unique word for every part in the circuit.
* block_comment: A comment describing the part.
* group_name: The group name of the xlsx file.
* dxf_filename: Output dxf file name.

- Example-1:
   ```bash
   python xlsx2dxf_converter.py "Blocks\Delta DVP12SA2.xlsx" "communication_all" "PLC1" "Delta PLC\nDVP-DVP-SA2" "Main\nPLC" "output2.dxf"
   ```
   
- Example-2:
   ```
    python xlsx2dxf_converter.py "Blocks\Teco L510.xlsx" "power" "VFD23" "Teco L510\n210-SH1F-P 0.75kW" "Main\nInverter\nRollers" "output.dxf"
   ```
3. **Run Using the GUI**:
    Execute the script `xlsx2dxf_gui.py1. The dialog window has the following widgets:
- **Reload Button**: Whenever you modify and save an XLSX file, it's essential to click the reload button. This action ensures that the changes take effect.
- **Categories**: This widget enables you to select a category. Categories within the XLSX file can be any word. To maintain consistency, it is recommended to implement a validation list rule in the `Block Template.xlsx` file. This ensures uniformity when duplicating the template to generate a new block XLSX.
- **Group Name**: Select the desired group name.
- **Block Label**: Add the label for the block.
- **Block Model**: Add the comment for the block . The reason that this is different from the block name,its because multiple part models can be represented by the same block.
- **Block Comment**: A comment that describes the block.
 





