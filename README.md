# Auto-Inspection - Vehicle Inspection Software
# Overview
Auto-Inspection is a desktop software application designed to streamline the vehicle inspection process. Developed using Python with the Tkinter GUI toolkit, it provides a user-friendly interface for inspectors to record vehicle details, perform various inspection checks, print comprehensive results, and securely save data to a database. The software is built to be compatible with both 32-bit and 64-bit Windows operating systems and includes a secure login system to protect sensitive inspection data.

This project was developed as a junior-year endeavor, focusing on desktop application development, database integration, and basic security features.

# Features
Auto-Inspection offers the following key functionalities:

Secure Login: A robust login system ensures that only authorized personnel can access the application.

Vehicle Data Entry: Input and manage essential vehicle information (e.g., make, model, VIN, license plate).

Inspection Checklist: A guided interface for performing various inspection checks (e.g., engine, brakes, lights, tires, body).

Result Recording: Record inspection findings, including pass/fail statuses, notes, and severity.

Printable Reports: Generate and print detailed inspection reports for clients or records.

Database Storage: Persistently save all inspection data to a Microsoft Access database.

Cross-Architecture Compatibility: Designed to run on both 32-bit and 64-bit Windows systems.

# Technologies Used
The Auto-Inspection software is built with the following technologies:

Programming Language: Python

Graphical User Interface (GUI): Tkinter (Python's standard GUI toolkit)

Database: Microsoft Access (.mdb or .accdb files)

Database Connectivity: pyodbc (for connecting Python to MS Access)

Report Generation: Python's built-in capabilities or external libraries for simple text/PDF output.

# Getting Started
To get Auto-Inspection running on your local machine, follow these steps.

# Prequisites
Python: Python 3.x installed (compatible with both 32-bit and 64-bit versions).

Download Python

Microsoft Access Database Engine: Even if you don't have Microsoft Access installed, you might need the Microsoft Access Database Engine (ACE Redistributable) to allow pyodbc to connect to .mdb or .accdb files. Ensure you download the correct 32-bit or 64-bit version matching your Python installation.

Microsoft Access Database Engine 2010 Redistributable (for .mdb)

Microsoft Access Database Engine 2016 Redistributable (for .accdb)

# Required Python Libraries:

pip install pyodbc

Installation Steps
Clone the Repository (if applicable): If your project is in a Git repository, clone it:

git clone https://github.com/your-username/Auto-Inspection.git # Replace with your actual repo URL
cd Auto-Inspection

Otherwise, simply navigate to the directory where your project files are located.

Install Python Dependencies: Open your command prompt or terminal in the project directory and install the necessary Python library:

pip install pyodbc

# Database Setup:

Ensure your Microsoft Access database file (e.g., inspection_data.accdb or inspection_data.mdb) is located in the expected path relative to your application's main script.

Verify that the database file contains the necessary tables (e.g., Users, Vehicles, Inspections, InspectionDetails) with the correct schema.

# Running the Application
Once all prerequisites are met and dependencies are installed, you can run the main application script:

python main_app.py # Replace main_app.py with your actual main script name

The application's login screen should appear, prompting for credentials.

# Usage
Login: Upon launching the application, enter your username and password to gain access.

Navigate: Use the intuitive GUI to navigate between vehicle entry, inspection forms, and report generation.

Record Data: Fill in vehicle details and complete the inspection checklist.

Save & Print: Save the inspection record to the database and generate a printable report.
