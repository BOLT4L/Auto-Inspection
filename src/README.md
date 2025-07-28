
# Auto Vault Records Keeper (Desktop)

A comprehensive desktop application for managing car annual inspection records using Python, Tkinter, and MS Access database.

## Features

- Local vehicle inspection management
- PDF report generation
- Microsoft Access database integration

## Requirements

- Python 3.7+
- Microsoft Access Database Engine (ODBC driver)
- pyodbc library

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Ensure Microsoft Access Database Engine is installed on your system
   - For 64-bit systems: Download and install Microsoft Access Database Engine 2016 Redistributable
   - The ODBC driver for Access must be available

## Usage

1. Run the application:
```bash
python main.py
```

2. **First Time Setup**:
   - Click "Select Database" to create a new Access database file or select existing one
   - The application will automatically create the required table structure

3. **Adding Inspections**:
   - Use the "Add New Inspection" tab
   - Fill in vehicle information, inspection details, and checklist
   - Click "Save Inspection"

4. **Viewing Records**:
   - Use "View Inspections" tab to see all records
   - Edit or delete records as needed

5. **Searching**:
   - Use "Search" tab to find specific inspections
   - Search by registration number, owner name, make, or model

6. **Reports**:
   - Export data to CSV
   - Generate expiring certificate reports
   - View inspection statistics

## Database Structure

The application creates an "Inspections" table with the following fields:
- Vehicle Information: Registration, Make, Model, Year, Engine/Chassis numbers
- Owner Details: Name, Address, Phone
- Inspection Data: Date, Inspector, Center, Certificate details
- Results: Overall result, individual checklist items, defects/notes
- System: Creation date, unique ID

## Features Detail

### Inspection Checklist
- Brakes, Lights, Steering, Suspension
- Tyres, Exhaust, Seat Belts, Mirrors
- Horn, Windscreen
- Each item can be marked as PASS/FAIL/ADVISORY

### Reports
- **CSV Export**: Export all inspection data
- **Monthly Reports**: Summary of monthly inspection activities
- **Expiring Certificates**: Alerts for certificates expiring in next 3 months
- **Statistics**: Overall system statistics and trends

### Data Validation
- Required fields validation
- Date format checking
- Numeric field validation
- Duplicate registration checking

## Troubleshooting

### Common Issues:

1. **Database Connection Error**:
   - Ensure Microsoft Access Database Engine is installed
   - Check if the database file path is accessible
   - Verify ODBC driver availability

2. **Permission Issues**:
   - Run application as administrator if needed
   - Check file/folder permissions for database location

3. **Missing Dependencies**:
   - Install all required Python packages: `pip install -r requirements.txt`

## Technical Notes

- Uses Tkinter for cross-platform GUI
- ODBC connection via pyodbc for database operations
- JSON storage for complex data (checklist results)
- Scrollable forms for better UX on smaller screens
- Comprehensive error handling and user feedback

## Future Enhancements

- Photo attachment support for inspection evidence
- Digital signature capability
- Automatic backup functionality
- Multi-user support with login system
- Integration with external APIs for vehicle data
- Print functionality for inspection certificates
- Advanced reporting with charts and graphs

## Support

For issues or questions, please check the error messages and ensure all requirements are properly installed.
