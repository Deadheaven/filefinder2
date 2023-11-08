import os
import pandas as pd
import psutil
import mysql.connector
import csv
from datetime import datetime, timedelta
import logging
import xlrd
from dotenv import load_dotenv
load_dotenv()
import time
import socket


# Define the list of file extensions to search for
file_extensions = os.getenv("FILE_EXTENSIONS").split(",")  # Add more extensions as needed

# Define patterns to identify sensitive data in file names
sensitive_patterns = os.getenv("SENSITIVE_PATTERNS").split(",")

# Define your MySQL database connection details
host = os.getenv("MYSQL_HOST")  # Replace with the MySQL server address
port = os.getenv("MYSQL_PORT")  # Replace with the MySQL server port
database_name = os.getenv("MYSQL_DATABASE")
username = os.getenv("MYSQL_USERNAME")
password = os.getenv("MYSQL_PASSWORD")
n_days = int(os.getenv("N_DAYS"))

# Configure logging to a file
log_file = "error.log"
logging.basicConfig(filename=log_file, level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

def get_ip_address():
    hostname = socket.gethostname()    
    IPAddr = socket.gethostbyname(hostname)    
    return IPAddr

def get_drives():
    drives = []
    try:
        partitions = psutil.disk_partitions(all=True)  # Include all drives
        for partition in partitions:
            if partition.device:
                drives.append(partition.device)
    except Exception as e:
        # Log the error to the log file
        logging.error(f"Error retrieving drive information: {str(e)}")
    return drives

# Define a custom exception class for file-related errors
class FileError(Exception):
    pass

def is_recently_accessed_or_modified(file_path, n_days):
    try:
        now = datetime.now()
        file_info = os.stat(file_path)
        file_mtime = datetime.fromtimestamp(file_info.st_mtime)
        file_atime = datetime.fromtimestamp(file_info.st_atime)
        delta_mtime = now - file_mtime
        delta_atime = now - file_atime
        return delta_mtime.days <= n_days or delta_atime.days <= n_days
    except Exception as e:
        # Log the error to the log file
        logging.error(f"Error checking file modification/access time: {str(e)}")
        return False

def is_sensitive_file(file_path, sensitive_patterns):
    try:
        file_name = os.path.basename(file_path).lower()
        with open(file_path, "r", encoding="utf-8", errors="ignore") as file:
            file_content = file.read().lower()
        for pattern in sensitive_patterns:
            if pattern in file_name or pattern in file_content:
                return True
    except Exception as e:
        # Log the error to the log file
        logging.error(f"Error checking file for sensitive data: {str(e)}")
    return False

def search_files(root_dir, extensions, n_days, sensitive_patterns):
    found_assets = []
    try:
        for foldername, subfolders, filenames in os.walk(root_dir):
            for filename in filenames:
                if any(filename.lower().endswith(ext) for ext in extensions):
                    file_path = os.path.join(foldername, filename)
                    if is_recently_accessed_or_modified(file_path, n_days) and not is_sensitive_file(file_path, sensitive_patterns):
                        found_assets.append(file_path)
    except Exception as e:
        # Log the error to the log file
        logging.error(f"Error scanning files: {str(e)}")
    return found_assets

def upsert_to_database(file_path, connection):
    cursor = connection.cursor()
    file_size = os.path.getsize(file_path)
    file_name = os.path.basename(file_path)
    file_extension = os.path.splitext(file_name)[1]
    modification_time = datetime.fromtimestamp(os.path.getmtime(file_path))
    access_time = datetime.fromtimestamp(os.path.getatime(file_path))
    creation_time = datetime.fromtimestamp(os.path.getctime(file_path))

    # Perform an upsert based on file_path
    cursor.execute('''
        INSERT INTO file_name_info (file_path, file_size, file_name, file_extension, modification_time, access_time, creation_time)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE
        file_size = %s, file_name = %s, file_extension = %s, modification_time = %s, access_time = %s, creation_time = %s;
    ''', (file_path, file_size, file_name, file_extension, modification_time, access_time, creation_time,
       file_size, file_name, file_extension, modification_time, access_time, creation_time))
    connection.commit()

# Define the SQL statement for table creation
create_table_sql = '''
CREATE TABLE IF NOT EXISTS file_name_info (
    file_name_info_pk INT AUTO_INCREMENT PRIMARY KEY,
    file_path VARCHAR(255) UNIQUE,
    file_size BIGINT,
    file_name VARCHAR(255),
    file_extension VARCHAR(10),
    modification_time DATETIME,
    access_time DATETIME,
    creation_time DATETIME
);
'''

def create_dataassets_table(connection):
    try:
        cursor = connection.cursor()
        cursor.execute(create_table_sql)
        connection.commit()
        print("file_name_info table created or already exists.")
    except mysql.connector.Error as err:
        logging.error(f"Error creating file_name_info table: {err}")

# Function to create a new table for .xls files
def create_xls_file_sheet_table(connection, xls_files):
    try:
        cursor = connection.cursor()
        for xls_file in xls_files:
            workbook = xlrd.open_workbook(xls_file)
            sheets = workbook.sheet_names()
            for sheet_name in sheets:
                sheet = workbook.sheet_by_name(sheet_name)
                num_rows = sheet.nrows
                num_cols = sheet.ncols

                # Create the xls_file_sheet table
                cursor.execute(f'''
                    CREATE TABLE IF NOT EXISTS xls_file_sheet (
                        xls_file_sheet_pk INT AUTO_INCREMENT PRIMARY KEY,
                        file_name_info_fk INT ,
                        sheet_name VARCHAR(255),
                        total_cols INT,
                        total_rows INT,
                        UNIQUE KEY unique_file_sheet (xls_file_sheet_pk, sheet_name),
                        FOREIGN KEY (file_name_info_fk) REFERENCES file_name_info(file_name_info_pk)
                    );
                ''')
                connection.commit()
                cursor.execute('''
                INSERT INTO xls_file_sheet (file_name_info_fk, sheet_name, total_cols, total_rows)
                VALUES (
                    (SELECT file_name_info_pk FROM file_name_info WHERE file_path = %s),
                    %s, %s, %s
                )ON DUPLICATE KEY UPDATE
                               total_cols=VALUES(total_cols),
                               total_rows=VALUES(total_rows);
                
                ''', (xls_file, sheet_name, num_cols, num_rows))
                connection.commit()
        print("Tables for .xls files created and data inserted.")
    except Exception as e:
        logging.error(f"Error creating .xls file tables and inserting data: {str(e)}")

# Function to create a new table for .xls file rows
def create_xls_file_sheet_row_table(connection, xls_files):
    try:
        cursor = connection.cursor()
        for xls_file in xls_files:
            xls_data = pd.read_excel(xls_file, sheet_name=None)  # Read all sheets

            for sheet_name, sheet in xls_data.items():
                num_rows, num_cols = sheet.shape

                # Create the xls_file_sheet_row table
                cursor.execute(f'''
                    CREATE TABLE IF NOT EXISTS xls_file_sheet_row (
                        xls_file_sheet_row_pk INT AUTO_INCREMENT PRIMARY KEY,
                        xls_file_sheet_fk INT,
                        sheet_name VARCHAR(255),
                        col_no INT,
                        row_no INT,
                        is_row VARCHAR(3),
                        col_data_1 VARCHAR(255),
                        col_data_2 VARCHAR(255),
                        col_data_3 VARCHAR(255),
                        col_data_4 VARCHAR(255),
                        col_data_5 VARCHAR(255),
                        col_data_6 VARCHAR(255),
                        col_data_7 VARCHAR(255),
                        col_data_8 VARCHAR(255),
                        col_data_9 VARCHAR(255),
                        col_data_10 VARCHAR(255),
                        is_truncate VARCHAR(3),
                        UNIQUE KEY unique_file_sheet (xls_file_sheet_fk, sheet_name,row_no),
                        FOREIGN KEY (xls_file_sheet_fk) REFERENCES xls_file_sheet(xls_file_sheet_pk)
                    );
                ''')
                connection.commit()

                # Insert the first 10 columns of data into the table, or all if there are fewer than 10 columns
                for row_idx in range(min(int(os.getenv("MIN_ROW")), num_rows)):  # Read up to the first 3 rows
                    is_row = "no" if row_idx == 0 else "yes"  # First row is a heading, the rest are data
                    col_data = sheet.iloc[row_idx, :10].tolist()  # Take the first 10 columns
                    col_data.extend(["NULL"] * (10 - len(col_data)))  # Fill the remaining columns with "NULL"
                    col_data = [str(data)[:255] for data in col_data]  # Truncate data if necessary
                    # Check for truncation if there are more than 10 columns
                    is_truncate = "yes" if num_cols > 10 else "no"

                    cursor.execute(f'''
                    INSERT INTO xls_file_sheet_row (xls_file_sheet_fk, sheet_name, col_no, row_no, is_row,
                    col_data_1, col_data_2, col_data_3, col_data_4, col_data_5,
                    col_data_6, col_data_7, col_data_8, col_data_9, col_data_10, is_truncate)
                    VALUES (
                    (SELECT xls_file_sheet_pk FROM xls_file_sheet WHERE sheet_name = %s LIMIT 1),
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                    );
                    ''', (sheet_name, sheet_name, num_cols, row_idx + 1, is_row, *col_data, is_truncate))
        print("Tables for .xls file rows created and data inserted.")
    except Exception as e:
        logging.error(f"Error creating .xls file row tables and inserting data: {str(e)}")
#function for audit table        
def create_audit_table(connection, ip, start_time, end_time, elapsed_time):
    try:
        cursor = connection.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS audit_info (
                pc_ip_address TEXT,
                start_time TIMESTAMP,
                end_time TIMESTAMP,
                duration TEXT
            );
        ''')
        connection.commit()

        cursor.execute('''
            INSERT INTO audit_info (pc_ip_address, start_time,end_time, duration)
            VALUES (%s, FROM_UNIXTIME(%s), FROM_UNIXTIME(%s), %s);
        ''', (ip, start_time, end_time, end_time-start_time))
        connection.commit()
        print("Table for audit created and data inserted.")
    except Exception as e:
        logging.error(f"Error creating audit table and inserting data: {str(e)}")


if __name__ == "__main__":
    start_time = time.time()
    drives = get_drives()
    if not drives:
        print("No drives found.")
    else:
        print("Available drives:")
        for i, drive in enumerate(drives, start=1):
            print(f"{i}. {drive}")

        scan_option = input("Choose an option:\n1. Full Scan\n2. Drive-specific Scan\nEnter the option number (1 or 2): ")

        try:
            if scan_option == '1':
                print(f"Performing a full scan for data assets modified or accessed in the last {n_days} days:")
                found_assets = []
                for drive in drives:
                    found_assets.extend(search_files(drive, file_extensions, n_days, sensitive_patterns))
            elif scan_option == '2':
                drive_choice = input("Enter the drive letter to scan (e.g., C, D, E, ...): ").upper()
                if drive_choice in [d[0] for d in drives]:
                    selected_drive = [d for d in drives if d[0] == drive_choice][0]
                    print(f"Scanning {selected_drive} for data assets modified or accessed in the last {n_days} days:")
                    found_assets = search_files(selected_drive, file_extensions, n_days, sensitive_patterns)
                else:
                    print("Invalid drive choice.")
                    found_assets = []
            else:
                print("Invalid option. Please choose 1 for Full Scan or 2 for Drive-specific Scan.")
                found_assets = []
        except ValueError:
            print("Invalid input. Please enter a valid option or drive letter.")

        connection = None
        try:
            connection = mysql.connector.connect(
                host=host,
                port=port,
                database=database_name,
                user=username,
                password=password
            )
            # Create the file_name_info table if it doesn't exist
            create_dataassets_table(connection)

            if found_assets:
                for asset in found_assets:
                    upsert_to_database(asset, connection)
                print(f"Scan results for the last {n_days} days saved to the MySQL database.")
            else:
                print("No data assets found.")
        except Exception as e:
            # Log the error to the log file
            logging.error(f"Error connecting to the database: {str(e)}")
        finally:
            if connection:
                connection.close()
        if ".xls" in file_extensions:
            xls_files = [file for file in found_assets if file.lower().endswith(".xls")]
            if xls_files:
                connection = mysql.connector.connect(
                    host=host,
                    port=port,
                    database=database_name,
                    user=username,
                    password=password
                )
                create_xls_file_sheet_table(connection, xls_files)
                create_xls_file_sheet_row_table(connection, xls_files)
                connection.close()
            else:
                print("No .xls files found.")
    end_time = time.time()
    elapsed_time = end_time - start_time
    ip=get_ip_address()
    connection = mysql.connector.connect(
            host=host,
            port=port,
            database=database_name,
            user=username,
            password=password
        )
    create_audit_table(connection,ip,start_time,end_time,elapsed_time)
    connection.close()
