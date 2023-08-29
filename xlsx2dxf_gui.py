import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QListWidget, QLabel,QTextEdit, QLineEdit, QPushButton
import os
import xlsx2dxf_converter

class Category():
    def __init__(self, category_name):
        self.category_name = category_name
        self.blocks = []

class Block():
    def __init__(self, block_name, block_category):
        self.block_name = block_name
        self.block_category = block_category
        self.groups = []

class Group():
    def __init__(self, group_name):
        self.group_name = group_name

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("xlsx2dxf-converter GUI")
        
        # --- Reload Button
        # If there is a change to a xlsx block file then the reload button must be pressed
        self.reload_button = QPushButton("Reload Blocks")
        self.reload_button.clicked.connect(self.load_categories_xlsx)             
        
        # --- Categories List
        self.category_list = QListWidget()
        self.category_list.setFixedHeight(250)
        self.category_list.itemSelectionChanged.connect(self.populate_blocks)
        
        # --- Blocks Names List
        self.block_list = QListWidget()
        self.block_list.itemSelectionChanged.connect(self.populate_groups)
        
        # --- Groups Names List
        self.group_list = QListWidget()
        
        # --- Create a list of all List Widgets 
        self.list_widgets = [self.category_list,self.block_list,self.group_list]
        for list_widget in self.list_widgets:
                    list_widget.itemSelectionChanged.connect(self.checkListWidgets)
            
        # --- Block Label Input
        self.block_label_input = QLineEdit()
        self.block_label_input.setPlaceholderText("Enter Block Label")

        # --- Block Model Input
        self.block_model_input = QTextEdit()
        self.block_model_input.setFixedHeight(75)
        self.block_model_input.setPlaceholderText("Enter Block Model")    
        
        # --- Block Comment Input
        self.block_comment_input = QTextEdit()
        self.block_comment_input.setFixedHeight(75)
        self.block_comment_input.setPlaceholderText("Enter Block Comment")

        # --- Generate Block Button
        self.generate_button = QPushButton("Generate Block")
        self.generate_button.clicked.connect(self.generate_block)
        self.generate_button.setEnabled(False)
                  
        # --- Add Widgets to Main Window
        layout = QVBoxLayout()
        layout.addWidget(self.reload_button)
        layout.addWidget(QLabel("Categories"))
        layout.addWidget(self.category_list)
        layout.addWidget(QLabel("Blocks"))
        layout.addWidget(self.block_list)
        layout.addWidget(QLabel("Groups"))
        layout.addWidget(self.group_list)
        layout.addWidget(QLabel("Block Label"))
        layout.addWidget(self.block_label_input)
        layout.addWidget(QLabel("Block Model"))
        layout.addWidget(self.block_model_input)  
        layout.addWidget(QLabel("Block Comment"))
        layout.addWidget(self.block_comment_input)      
        layout.addWidget(self.generate_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.load_categories_xlsx()

        
    def load_categories_xlsx(self):
        # --- Change Text of the Reload Button
        self.reload_button.setText('Loading Blocks xlsxs')
        self.repaint()
        
        # --- Clear Lists
        self.category_list.clear()
        self.block_list.clear()
        self.group_list.clear()
        
        # --- Import all xlsx File From Blocks Directory
        blocks_folder_name = 'Blocks'
        current_directory = os.getcwd()
        blocks_directory = f'{current_directory}{os.path.sep}{blocks_folder_name}{os.path.sep}'     #Create the directory of Blocks Folder
        # --- Get All xlsx Files in Blocks Folder        
        xlsx_files = [file for file in os.listdir(blocks_directory) if file.endswith('.xlsx')]    
        
        # --- Create Block Objects using xlsx2dxf_converter
        # Iterate through the list of xlsx files and convert them to Block objects.
        blocks = []
        for xlsx_name in xlsx_files:
            xlsx_file_path = f'{blocks_directory}{xlsx_name}'
            blocks.append(xlsx2dxf_converter.parse_xlsx_file(xlsx_file_path))
               
        # --- Create Category Objects
        # Iterate through blocks to create category objects.
        # If a category does not exist yet, a new one is created and associated with the block.
        # If the category already exists, the block is added to the existing category.
        categories = []        
        for block in blocks:
            if (block.block_category not in [cat.category_name for cat in categories]):
                new_category = Category(block.block_category)
                new_category.blocks.append(block)
                categories.append(new_category)
                
            else:
                for existing_category in categories:
                    if (existing_category.category_name == block.block_category):
                        break
                existing_category.blocks.append(block)

        self.categories = sorted(categories, key=lambda category: category.category_name)
        self.reload_button.setText('Reload Blocks')
        
        self.populate_categories()
        
        
    def checkListWidgets(self):
        all_selected = all(list_widget.selectedItems() for list_widget in self.list_widgets)
        self.generate_button.setEnabled(all_selected)

        
    def populate_categories(self):
        self.category_list.clear()
        for category in self.categories:
            self.category_list.addItem(category.category_name)

    def populate_blocks(self):      
        selected_category = self.category_list.currentItem().text()
        selected_blocks = [block.block_name for category in self.categories if category.category_name == selected_category for block in category.blocks]
        self.block_list.clear()
        self.block_list.addItems(selected_blocks)

    def populate_groups(self):    
        selected_category = self.category_list.currentItem().text()
        selected_block = self.block_list.currentItem().text()
        selected_groups = [group.group_name for category in self.categories if category.category_name == selected_category for block in category.blocks if block.block_name == selected_block for group in block.groups]
        self.group_list.clear()
        self.group_list.addItems(selected_groups)

    def generate_block(self):
        selected_category = self.category_list.currentItem().text()
        selected_block_name = self.block_list.currentItem().text()
        selected_group_name = self.group_list.currentItem().text()
        selected_block_label = self.block_label_input.text()
        selected_block_model = self.block_model_input.toPlainText()
        selected_block_comment = self.block_comment_input.toPlainText()
        
        print("Selected Block:", selected_block_name)
        print("Selected Group:", selected_group_name)
        print("Block Label:", selected_block_label)
        print("Block Model:", selected_block_model)
        print("Block Comment:", selected_block_comment)
        
        # --- Get Block Object from Name
        selected_block = None
        for category in self.categories:
            if category.category_name == selected_category:
                for block in category.blocks:
                    if block.block_name == selected_block_name:
                        selected_block = block
                        break
                break
        
        doc = xlsx2dxf_converter.create_dxf_document(selected_block,selected_group_name,selected_block_label,selected_block_model,selected_block_comment)
    
        dxf_filename = 'output.dxf'
        doc.saveas(dxf_filename)
        print(f"DXF file '{dxf_filename}' created.")    
        
    def closeEvent(self, event):
        app.quit()
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
