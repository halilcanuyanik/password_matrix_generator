# Password Matrix Generator

## Overview
**Password Matrix Generator** is a Python-based desktop application that generates a secure password based on a predefined **letter shape pattern** in a random character matrix.  
It uses **Tkinter** for the graphical user interface and **OpenPyXL** for Excel export with highlighted password paths.

The program allows users to:
- Generate a random matrix of strings.
- Select a letter pattern from a predefined **LETTER_MAP**.
- Extract a password by following the shape of the selected letter within the matrix.
- Save the matrix to an Excel file with the letter's shape highlighted.

---

## Features
- **Custom Letter Patterns:** Each letter from **A–Z** has a unique 5×5 binary shape map.
- **Randomized Matrix:** Generates a matrix of random strings containing letters, digits, and special symbols.
- **GUI Interface:** User-friendly interface with drop-down menus and buttons.
- **Password Extraction:** Retrieves password characters from the letter's shape path.
- **Excel Export:** Saves the generated matrix to `.xlsx` with the selected letter's cells highlighted.

---

## Requirements
To run this project, you need:

- Python 3.8+
- Required Python libraries:
  ```bash
  pip install openpyxl
  ```

## How It Works

1. **Matrix Generation**
   - The program creates a matrix (e.g., 10×10) where each cell contains a random 4-character string.
   
2. **Letter Mapping**
   - `LETTER_MAP` defines the shape of each letter as a binary grid (`1` = part of the letter, `0` = empty space).
   
3. **Password Extraction**
   - When a letter and character index are selected, the program reads the corresponding characters from the matrix cells where the letter shape is present.
   
4. **Excel Export**
   - The program highlights the letter’s shape in yellow when saving to an Excel file.

---

## Usage

1. **Run the Program**
   ```bash
   python password_matrix_generator.py
   ```
2. **Select a Letter**
  - Choose any letter from A to Z in the dropdown menu.

3. **Select Character Index**
  - Choose which character (1–4) from each cell should be used in the password.

4. **Generate Password**
  - Click "Generate Password" to see the result in the output section.

5. **Save to Excel**
  - Click "Save to Excel" to export the matrix with the letter's path highlighted.

---

## File Structure
- password_matrix_generator.py   # Main program file (.py)
- password_matrix_generator.exe  # Main program file (.exe)
- README.md                      # Documentation

---

## License
This project is released under the MIT License.
