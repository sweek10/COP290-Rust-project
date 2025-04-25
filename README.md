# RUST SPREADSHEET PROGRAM #
## Overview ##

This project is a robust spreadsheet application written in Rust, designed to provide efficient and reliable spreadsheet functionality through both a command-line interface (CLI) and an optional web interface.   
The program supports core spreadsheet operations such as formula evaluation, cell dependency tracking, and data visualization, with advanced features like autofill, undo/redo, and file import available in extension mode.    
The project is modular, leveraging Rust's safety and performance features to ensure maintainability and scalability.   
This project was developed as part of the Rust Lab for the COP290 course by:              
Shreya Bhaskar (2023CS10941)   
Sweety Naveen Kumar (2023CS10172)   
Sambhav Singhwi (2023CS10722)     

## Running the Program
**Accessing the Web Interface (alternative method)** : When running the program with the --extension flag, the web interface is available at http://localhost:8000.

# Few Followups #
## 1) ## 
If any .csv or .xlsx file is to be loaded in the spreasheet while running the extension then the following command should be used - **make ext1-some_file.csv** in case of a .csv file and similarly **make ext1-some_file.xlsx** for loading a .xlsx file.
## 2) ##
After the presentation of our extensions on 22nd April, we have improved our web interface to support toggling between dark and light themes and also correct the **UNDO-REDO EXTENSION** which now works appropriately  
for all the file input extension along with other extensions.
