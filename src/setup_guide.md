
# Setup Guide for Car Inspection Management System

## Step-by-Step Installation

### 1. System Requirements
- Windows 10/11 (recommended) or Windows 7+
- Python 3.8+
- Microsoft Access (for database)
- At least 100MB free disk space

### 2. Install Python
If Python is not installed:
1. Download Python from https://python.org
2. During installation, check "Add Python to PATH"
3. Verify installation: Open Command Prompt and type `python --version`

### 3. Install Microsoft Access Database Engine
This is required for ODBC connectivity to Access databases:

**Option A: Microsoft Access Database Engine 2016**
1. Download from Microsoft's official website
2. Choose the version matching your system (32-bit or 64-bit)
3. Install with default settings

**Option B: If you have Microsoft Office installed**
- The ODBC driver should already be available
- Test by checking "ODBC Data Sources" in Windows Control Panel

### 4. Download and Setup Application
1. Extract the application files to a folder (e.g., `C:\CarInspection\`)
2. Open Command Prompt as Administrator
3. Navigate to the application folder: `cd C:\CarInspection\`
4. Install Python dependencies: `pip install pyodbc`

### 5. First Run
1. Double-click `main.py` or run `python main.py` from command prompt
2. Click "Select Database" button
3. Choose location and name for your database file (e.g., `CarInspections.accdb`)
4. The application will create the database structure automatically

### 6. Test the Application
1. Go to "Add New Inspection" tab
2. Enter some test data
3. Click "Save Inspection"
4. Check "View Inspections" tab to see the saved record

## Troubleshooting Common Setup Issues

### Issue: "No module named 'pyodbc'"
**Solution**: Install pyodbc
```bash
pip install pyodbc
```

### Issue: "Data source name not found"
**Solution**: Install Microsoft Access Database Engine
- Download and install Access Database Engine 2016 Redistributable
- Restart your computer after installation

### Issue: "Permission denied" when creating database
**Solution**: 
- Run the application as Administrator
- Or choose a different location for the database (e.g., Documents folder)

### Issue: Application window appears too small/large
**Solution**: 
- The application is designed to be resizable
- Drag window corners to adjust size
- Window size and position are not saved between sessions

### Issue: "tkinter not found"
**Solution**: 
- Tkinter comes with Python by default
- If missing, reinstall Python with "tcl/tk and IDLE" option checked

## Advanced Configuration

### Custom Database Location
- You can place the database file anywhere accessible
- Network drives are supported but may have slower performance
- Consider regular backups of the database file

### Multiple Users
- The application supports multiple users accessing the same database
- Place the database on a shared network location
- Each user needs the application installed locally

### Performance Optimization
- For large datasets (>10,000 records), consider:
  - Regular database maintenance
  - Archiving old records
  - Using indexed searches

## Security Considerations

### Database Security
- The Access database is not password protected by default
- Consider adding database-level password protection for sensitive data
- Regular backups are recommended

### Data Privacy
- The application stores all data locally in the Access database
- No data is transmitted over the internet
- Ensure proper file permissions on the database file

## Backup and Maintenance

### Regular Backups
1. Close the application
2. Copy the `.accdb` database file to a backup location
3. Consider automating this process with Windows Task Scheduler

### Database Maintenance
- Access databases can benefit from periodic compacting
- Use Access application to compact and repair if available
- Monitor database file size growth

## Getting Help

### Log Files
- The application displays errors in popup messages
- For technical support, note the exact error message
- Check Windows Event Viewer for system-level issues

### Common Support Questions
1. **Q**: Can I import data from Excel?
   **A**: Currently not supported directly. Consider manual entry or CSV import features.

2. **Q**: Can I run this on Mac or Linux?
   **A**: The application uses Windows-specific ODBC drivers. Linux/Mac would require modifications.

3. **Q**: How many records can the database hold?
   **A**: Access databases can handle millions of records, but performance may degrade with very large datasets.

4. **Q**: Can I customize the inspection checklist?
   **A**: Currently, the checklist is fixed in the code. Customization would require code modifications.

## Deployment for Organizations

### Single Computer Setup
- Install on one computer for single-user access
- Database stored locally

### Multi-User Network Setup
1. Install the application on each user's computer
2. Place the database file on a shared network drive
3. All users point to the same database file
4. Ensure proper network permissions

### IT Administrator Notes
- No special Windows services required
- Application runs in user space
- Database file requires read/write access for all users
- Consider Group Policy for mass deployment if needed
