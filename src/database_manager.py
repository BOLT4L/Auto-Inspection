import pyodbc
from datetime import datetime

class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.connection = None
    
    def connect(self):
        """Establish database connection"""
        try:
            conn_str = f'Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_path};'
            self.connection = pyodbc.connect(conn_str)
            return True
        except Exception as e:
            print(f"Database connection error: {e}")
            return False
    
    def create_tables(self):
        """Create database tables with all inspection columns"""
        if not self.connection:
            return False
        
        try:
            cursor = self.connection.cursor()
            
            # Check if table exists
            try:
                cursor.execute("SELECT COUNT(*) FROM RESULT")
                print("Inspections table already exists")
                return True
            except:
                pass  # Table doesn't exist, create it
            
            # Create Inspections table with all columns
            create_table_sql = """
            CREATE TABLE RESULT (
    ID VARCHAR(10) ,
    InspectionDate DATETIME,
    PlateNo VARCHAR(20),
    CustName VARCHAR(100),
    VehModel VARCHAR(50),
    VehicleMade VARCHAR(50),
    ChassisNo VARCHAR(50),
    EngineNo VARCHAR(50),
    Status VARCHAR(20),
    Inspector VARCHAR(50),
    FrontAxileSideSlip DOUBLE,
    RearAxileSideSlip DOUBLE,
    SuspenstionFrontLeftEfficieny DOUBLE,
    SuspenstionFrontRightEfficieny DOUBLE,
    SuspensionFrontRandLDifference DOUBLE,
    FrontLeftWeight DOUBLE,
    FrontRightWeight DOUBLE,
    SuspensionRareLeftEfficency DOUBLE,
    SuspenstionRareRightEfficency DOUBLE,
    SuspensionrareRandLDifference DOUBLE,
    RareLeftWeight DOUBLE,
    RareRightWeight DOUBLE,
    RollingFrictionFrontLeftAxle DOUBLE,
    RollingFrictionFrontRightAxle DOUBLE,
    RollingFrictionFrontDifference DOUBLE,
    RollingFrictionRareLeftaxle DOUBLE,
    RollingFrictionRareRightaxle DOUBLE,
    RollingFrictionRareDifference DOUBLE,
    OutOfRoundnessOrOvalityFrontLeft DOUBLE,
    OutOfRoundnessOrOvalityFrontRight DOUBLE,
    OutOfRoundnessOrOvalityRearLeft DOUBLE,
    OutofRoundnessOrOvalityRearRight DOUBLE,
    MaximumServiceBrakeForceFrontLeft DOUBLE,
    MaximumServiceBrakeForceFrontRight DOUBLE,
    MaximumServiceBrakeForceFrontDifference DOUBLE,
    MaximumServiceBrakeForceRareLeft DOUBLE,
    MaximumServiceBrakeForceRareRight DOUBLE,
    MaximumServiceBrakeForceRareDifference DOUBLE,
    ServiceBrakeForceFrontLeft DOUBLE,
    ServiceBrakeForceFrontRight DOUBLE,
    ServiceBrakeForceFrontDifference DOUBLE,
    ServiceBrakeForceRareLeft DOUBLE,
    ServiceBrakeForceRareRight DOUBLE,
    ServiceBrakeForceRareDifference DOUBLE,
    Front1ServiceBrakeEfficiency DOUBLE,
    Front2ServiceBrakeEfficiency DOUBLE,
    Rear1ServiceBrakeEfficiency DOUBLE,
    Rear2ServiceBrakeEfficiency DOUBLE,
    Rear3ServiceBrakeEfficiency DOUBLE,
    Rear4ServiceBrakeEfficiency DOUBLE,
    Rear5ServiceBrakeEfficiency DOUBLE,
    Rear6ServiceBrakeEfficiency DOUBLE,
    Rear7ServiceBrakeEfficiency DOUBLE,
    Rear8ServiceBrakeEfficiency DOUBLE,
    TotalServiceBrakeEfficiency DOUBLE,
    FrontAxleWeight DOUBLE,
    RearAxleWeight DOUBLE,
    TotalVehicleweight DOUBLE,
    ParkingBrakeLeftForce DOUBLE,
    ParkingBrakeRightForce DOUBLE,
    ParkingBrakeLeftRightdifference DOUBLE,
    ParkingbrakeOutOfRoundnessLEft DOUBLE,
    ParkingBrakeOutOfRoundnessRight DOUBLE,
    TotalParkingBrakeEfficeny DOUBLE,
    HeadLightHighBeamIntensityLeft VARCHAR(50), -- Changed to VARCHAR
    HeadLightHighBeamIntensityRight VARCHAR(50), -- Changed to VARCHAR
    HeadLightHighBeamHorizontalDeviationLeft VARCHAR(50), -- Changed to VARCHAR
    HeadLightHighBeamHorizontalDeviationRight VARCHAR(50), -- Changed to VARCHAR
    HeadLightHighBeamVerticalDeviationLeft VARCHAR(50), -- Changed to VARCHAR
    HeadLightHighBeamVerticalDeviationRight VARCHAR(50), -- Changed to VARCHAR
    HeadLightLowBeamIntensityLeft VARCHAR(50), -- Changed to VARCHAR
    HeadLightLowBeamIntensityRight VARCHAR(50), -- Changed to VARCHAR
    HeadLightLowBeamHorizontalDeviationLeft VARCHAR(50), -- Changed to VARCHAR
    HeadLightLowBeamHorizontalDeviationRight VARCHAR(50), -- Changed to VARCHAR
    HeadLightLowBeamVerticalDeviationLeft VARCHAR(50), -- Changed to VARCHAR
    HeadLightLowBeamVerticalDeviationRight VARCHAR(50), -- Changed to VARCHAR
    FogLightIntensityLeft DOUBLE,
    FogLightIntensityRight DOUBLE,
    FogLightVerticalDeviationLeft DOUBLE,
    FogLightVerticalDeviationRight DOUBLE,
    FogLightHorizontalDeviationLeft DOUBLE,
    FogLightHorizontalDeviationRight DOUBLE,
    HC DOUBLE,
    CO DOUBLE,
    CO2 DOUBLE,
    O2 DOUBLE,
    COCOrr DOUBLE,
    Nox DOUBLE,
    Lamda DOUBLE,
    OilTemp DOUBLE,
    RPM DOUBLE,
    Opacimeter DOUBLE,
    PHOTO VARCHAR(255),
    FilePlace VARCHAR(100),
    LIBRENO VARCHAR(50),
    CARTYPE VARCHAR(50),
    MADEYEAR INTEGER,
    FUELTYPE VARCHAR(20),
    FLAGG VARCHAR(10),
    MACHINERESULT VARCHAR(20),
    VISUALRESULT VARCHAR(20),
    ISTRUCK BOOLEAN,
    LampHeight DOUBLE,
    VEHICLETYPE VARCHAR(30),
    PlateCodeN VARCHAR(10),
    PlateCodeT VARCHAR(10),
    RPTID VARCHAR(20),
    CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    LastModified DATETIME DEFAULT CURRENT_TIMESTAMP
);
            """
            
            cursor.execute(create_table_sql)
            
            # Create indexes for better performance
            cursor.execute("CREATE INDEX idx_PlateNo ON Inspections (PlateNo)")
            cursor.execute("CREATE INDEX idx_InspectionDate ON Inspections (InspectionDate)")
            cursor.execute("CREATE INDEX idx_Status ON Inspections (Status)")
            
            self.connection.commit()
            print("Inspections table created successfully")
            return True
            
        except Exception as e:
            print(f"Error creating tables: {e}")
            return False
    
    def save_inspection(self, inspection_data):
        """Save complete inspection data to database"""
        if not self.connection:
            return False
        
        try:
            # Set default values for required fields if not provided
            if 'InspectionDate' not in inspection_data:
                inspection_data['InspectionDate'] = datetime.now()
            if 'CreatedDate' not in inspection_data:
                inspection_data['CreatedDate'] = datetime.now()
            if 'LastModified' not in inspection_data:
                inspection_data['LastModified'] = datetime.now()
            
            
            # Remove 'CreatedDate' and 'LastModified' so the DB defaults are used
            if 'CreatedDate' in inspection_data:
                inspection_data.pop('CreatedDate')
            if 'LastModified' in inspection_data:
                inspection_data.pop('LastModified')
            # Convert booleans to int for YESNO fields
            if 'ISTRUCK' in inspection_data:
                inspection_data['ISTRUCK'] = int(bool(inspection_data['ISTRUCK']))
            # Ensure FLAGG is always 'M' or 'N', default to 'M' if missing or invalid
            if 'FLAGG' not in inspection_data or inspection_data['FLAGG'] not in ("M", "N"):
                inspection_data['FLAGG'] = "M"
            # Ensure VISUALRESULT and MACHINERESULT are set to required values (all uppercase)
            inspection_data['VISUALRESULT'] = 'PASS'
            inspection_data['MACHINERESULT'] = 'PASS'
            # Convert datetime objects to string
            for dt_col in ['InspectionDate']:
                if dt_col in inspection_data and isinstance(inspection_data[dt_col], datetime):
                    inspection_data[dt_col] = inspection_data[dt_col].strftime('%Y-%m-%d %H:%M:%S')
            columns = ', '.join(inspection_data.keys())
            placeholders = ', '.join(['?'] * len(inspection_data))
            values = tuple(inspection_data.values())
            # Use the correct table name 'RESULT' instead of 'Inspections'
            sql = f"INSERT INTO RESULT ({columns}) VALUES ({placeholders})"
            # Debug print statements
            print("\n[DEBUG] Attempting to save inspection:")
            print("Columns:", columns)
            print("Values:", values)
            print("SQL:", sql)
            
            cursor = self.connection.cursor()
            cursor.execute(sql, values)
            self.connection.commit()
            print("Inspection data saved successfully")
            return True
            
        except Exception as e:
            print(f"Error saving inspection: {e}")
            return False
    
    def get_inspection(self, plate_no=None, inspection_id=None):
        """Retrieve inspection data by plate number or ID"""
        if not self.connection:
            return None
        
        try:
            cursor = self.connection.cursor()
            
            if plate_no:
                cursor.execute("SELECT * FROM Inspections WHERE PlateNo = ?", plate_no)
            elif inspection_id:
                cursor.execute("SELECT * FROM Inspections WHERE ID = ?", inspection_id)
            else:
                return None
                
            columns = [column[0] for column in cursor.description]
            row = cursor.fetchone()
            
            if row:
                return dict(zip(columns, row))
            return None
            
        except Exception as e:
            print(f"Error retrieving inspection: {e}")
            return None
    
    def backup_database(self, backup_path):
        """Create a backup of the database"""
        try:
            import shutil
            shutil.copy2(self.db_path, backup_path)
            return True
        except Exception as e:
            print(f"Backup error: {e}")
            return False
    
    def close(self):
        """Close database connection"""
        if self.connection:
            self.connection.close()
            self.connection = None 