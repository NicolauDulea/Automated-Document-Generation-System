# Automated-Document-Generation-System

This project automates the creation of contracts by filling in a Word document template with data from an Excel file. It provides a simple graphical user interface (GUI) for selecting the necessary files and generating contracts, making the process efficient and user-friendly.

# Features
Excel and Word Integration: Automatically reads data from an Excel sheet and fills out predefined fields in a Word document template.
Dynamic Date Handling: Handles multiple date formats in the Excel file, choosing the most recent date if more than one is provided.
Personalized Contracts: Creates individual contract documents for each entry in the Excel file and saves them to a designated folder.
User-Friendly Interface: Provides an easy-to-use interface for selecting the Excel and Word files, and generating the documents.

# Requirements
To run this project, you need to have Python installed, along with several libraries, including PySimpleGUI, pandas, and python-docx.

# How to Use

1. Clone the Repository: Download or clone the project from GitHub to your local machine.

2. Prepare Your Files:
   - Ensure the Excel file contains the relevant columns, including client name, civil status, address, ID type and number, expiry date, and NIF.
   - Prepare a Word template that includes placeholders where the data will be inserted, such as name, address, and identification details.
   
3. Run the Application:
   - Open the program and use the interface to select your Excel file and Word document template.
   - Click the button to generate the contracts.
   
4. View Results:
   - The generated Word documents will be saved in a folder on your desktop, organized by the current date.
# Example Workflow
 Select your Excel file containing the client information.
 Select the Word template with placeholders.
 Generate the contracts—each client's information will be inserted into the template, and each contract will be saved as a separate document.
