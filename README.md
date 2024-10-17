# readexcelpython
# Excel Reader with Tkinter GUI

![Example Screenshot](example.png)

## 📋 Description

This is a simple Excel reader application built with Python's `tkinter` for the graphical user interface (GUI) and `pandas` for reading Excel files. The application allows users to select an Excel file from their system and displays the content in a nicely formatted table within the window.

## 🛠 Features

- **Load Excel files**: Users can easily select and load `.xlsx` or `.xls` files.
- **Display Data**: The content of the selected Excel file is displayed in a table (using `ttk.Treeview`) with proper formatting.
- **Customizable UI**: The interface is styled with a modern look, using `ttk.Style` for button and table theming.

## 🚀 Usage

### Prerequisites

Ensure you have Python installed along with the following libraries:
- `pandas`
- `openpyxl`
- `tkinter`

You can install the required dependencies by running:

```bash
pip install pandas openpyxl
Running the Application
Clone the repository:

bash
Copiar código
git clone https://github.com/LeoGidev/readexcelpython
cd excel-reader-tkinter
Run the Python script:

bash
Copiar código
python excel_reader.py
Click the Cargar Excel button, select an Excel file from your system, and the content will be displayed in a table within the window.

📸 Screenshot


This is a screenshot of the application after loading an Excel file:



📁 File Structure
bash
Copiar código
.
├── excel_reader.py        # Main Python script for the app
├── example.png            # Screenshot of the app
├── README.md              # This README file
└── requirements.txt       # Optional: requirements file for dependencies
💡 Customization
You can easily change the style or functionality by modifying the ttk.Style settings in the script, or adding new features such as filtering or editing the Excel data.

🧑‍💻 Contributing
Feel free to fork this repository, open issues, or submit pull requests. All kinds of improvements are welcome!

