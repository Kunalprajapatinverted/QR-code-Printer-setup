from datetime import datetime,timezone
import mysql.connector
from mysql.connector import Error
import pytz
import qrmain

# Database connection details
db_config = {
    'host': 'testdb.cbk6wukoo9oo.ap-south-1.rds.amazonaws.com',         # Replace with your DB host 
    'user': 'kunalprajapat',        # Replace with your DB user 
    'password': 'Kunal_inverted',    # Replace with your DB password 
    'database': 'BMS_barcode'       # Replace with your DB name 
}

# Fetch available product types (CellType), BatterySpec, and BatteryName from the BatterySpecification table
def fetch_product_types():
    try:
        # Establish a connection to the database
        connection = mysql.connector.connect(**db_config)

        if connection.is_connected():
            cursor = connection.cursor()

            # Query to fetch CellType, BatterySpec, and BatteryName from BatterySpecification table
            query = "SELECT DISTINCT CellType, BatterySpec, BatteryName FROM BatterySpecification"
            # print(query)
            cursor.execute(query)

            # Fetch all results
            result = cursor.fetchall()

            # Return the result as a list of tuples (CellType, BatterySpec, BatteryName)
            return result

    except Error as e:
        print(f"Error fetching product types: {e}")
        return []
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()


# Fetch the specification for a given product type (CellType) from the database
def get_specification_for_product_type(product_type):
    try:
        # Establish a connection to the database
        connection = mysql.connector.connect(**db_config)

        if connection.is_connected():
            cursor = connection.cursor()

            # Query to fetch the BatterySpec based on CellType
            query = "SELECT BatterySpec FROM BatterySpecification WHERE CellType = %s"
            cursor.execute(query, (product_type,))

            # Fetch the result (assuming only one specification per CellType)
            result = cursor.fetchone()

            if result:
                # Return the specification found
                return result[0]
            else:
                print(f"No specification found for product type {product_type}.")
                return "Unknown"

    except Error as e:
        print(f"Error fetching specification for {product_type}: {e}")
        return "Unknown"
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

# Function to select product type by number and map to specification and battery name
def select_product_and_specification():
    product_details = fetch_product_types()
    
    if not product_details:
        print("No product types found in the database.")
        return None, None, None  # Return None for all values
    
    print("Select product type:")
    for idx, (type_name, spec, name) in enumerate(product_details, start=1):
        print(f"{idx}. {type_name}")
        # print(f"{idx}. {type_name} - Spec: {spec}, Name: {name}")
    
    try:
        # User selects product type by number
        choice = int(input(f"Enter the number for product type (1-{len(product_details)}): "))
        if choice < 1 or choice > len(product_details):
            print("Invalid choice. Please select a valid number from the list.")
            return select_product_and_specification()
    except ValueError:
        print("Invalid input. Please enter a number.")
        return select_product_and_specification()

    # Retrieve selected product details
    product_type, specification, battery_name = product_details[choice - 1]
    
    return product_type, specification, battery_name

# Function to get IST time (Indian Standard Time)
def get_ist_time():
    """Get the current time in IST."""
    ist = pytz.timezone('Asia/Kolkata')  # IST timezone
    utc_time = datetime.now(pytz.utc)  # Get UTC time in a timezone-aware format
    ist_time = utc_time.astimezone(ist)  # Convert UTC to IST
    return ist_time
# print(get_ist_time())

# Generate date format with year and month as alphabetic characters (Year and Month only, without serial number)
def generate_date_format():
    """Generate the date format with year and month in uppercase alphabetic format."""
    today = get_ist_time()  # Use IST time
    
    # Convert year and month to uppercase alphabetic format
    year = chr(64 + (today.year % 100))  # Convert year to uppercase letter (2024 -> 'X' for '24')
    month = chr(64 + today.month)  # Convert month to uppercase letter (1 -> 'A', ..., 12 -> 'L')
    
    # Return the formatted date as "YearMonth" (e.g., "MX")
    formatted_date = f"{year}{month}"
    # print(formatted_date)
    return formatted_date

# Function to check if a barcode already exists for the current month and get the last serial number
def get_last_serial_number(product_type, specification):
    try:
        # Establish a connection to the database
        connection = mysql.connector.connect(**db_config)

        if connection.is_connected():
            cursor = connection.cursor()

            # Get current date in 'YYYY-MM' format
            today = get_ist_time()
            year_month = f"{today.year}-{today.month:02d}"
            # print(year_month)

            # Query to check if the barcode for this year and month already exists
            query = """
                SELECT MAX(SUBSTRING(BatteryBarcode, LENGTH(BatteryBarcode) - 3, 4))
                FROM BatteryCode 
                WHERE BatteryBarcode LIKE %s
            """
            barcode_prefix = f"IN{product_type}{generate_date_format()}{specification[:4]}"
            # print(barcode_prefix)
            # print(query)
            cursor.execute(query, (f"{barcode_prefix}%",))  # Using % as a wildcard for the serial part
            
            result = cursor.fetchone()
            print(f"Query Result: {result}")  # Debugging output

            if result and result[0]:
                # Extract last serial number and return it as an integer
                last_serial_str = result[0]
                print(f"Last Serial Number: {last_serial_str}")  # Debugging output
                return int(last_serial_str)  # Convert the serial number part to an integer
            else:
                # If no matching barcode is found, return 0 to indicate start at 0001
                return 0

    except Error as e:
        print(f"Error checking barcode in database: {e}")
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

    return 0

# Generate the barcode with BatterySpec and BatteryName included in the generated barcode
def generate_barcode():
    # Fixed part of the barcode
    brand_name = "IN"

    # Get product details (product_type, specification, battery_name)
    product_type, specification, battery_name = select_product_and_specification()

    if not product_type or not specification or not battery_name:
        print("Invalid product type selected. Exiting.")
        return None

    # Get the current year and month for barcode generation
    today = get_ist_time()
    year_month = f"{today.year}-{today.month:02d}"

    # Get the last serial number for the given product type and specification
    last_serial_number = get_last_serial_number(product_type, specification)

    # Increment the serial number
    serial_number = last_serial_number + 1

    # Format the serial number with 4 digits
    formatted_serial_number = f"{serial_number:04d}"

    # Generate formatted date (Year + Month)
    formatted_date = generate_date_format()

    # Format the barcode as per the new requirements
    barcode = f"{brand_name}{product_type}{formatted_date}{specification}{formatted_serial_number}"

    return barcode, product_type, specification, battery_name  # Return the barcode with additional details


# Upload the barcode to MySQL database
def upload_barcode_to_db(barcode, product_type, specification, battery_name):                       
    try:
        # Establish a connection to the database
        connection = mysql.connector.connect(**db_config)

        if connection.is_connected():
            cursor = connection.cursor()
            # Get IST time for the 'created_at' field
            ist_time = get_ist_time()
            # Format the IST time for database insertion
            ist_time_str = ist_time.strftime('%Y-%m-%d %H:%M:%S')

            # Prepare the SQL query to insert the barcode, specification, and battery name
            query = """
                INSERT INTO BatteryCode (BatteryBarcode, BatterySpec, BatteryName, CreatedDate)
                VALUES (%s, %s, %s, %s)
            """
            cursor.execute(query, (barcode, specification, battery_name, ist_time_str))

            # Commit the transaction
            connection.commit()

            print(f"Barcode {barcode} has been successfully uploaded to the database.")

    except Error as e:
        print(f"Error uploading barcode to database: {e}")
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()


# Function to print and upload the barcode
def print_and_upload_barcode():
    # Generate barcode and fetch additional details
    barcode, product_type, specification, battery_name = generate_barcode()

    if barcode:
        print(f"\nGenerated Barcode: {barcode}")
        # Upload the barcode along with specification and battery name
        upload_barcode_to_db(barcode, product_type, specification, battery_name)


if __name__ == "__main__":
    print_and_upload_barcode()