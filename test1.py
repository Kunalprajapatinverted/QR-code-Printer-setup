import mysql.connector
from datetime import datetime, timezone
import pytz  # Import pytz to work with time zones

# Function to connect to MySQL database and return the connection object
def connect_to_database():
    try:
        mydb = mysql.connector.connect(
            host="vidurindia.cbkgp37sk7e2.ap-south-1.rds.amazonaws.com", # Replace with your actual host, e.g., "localhost" or an IP address
            user="banku",  # Your MySQL username
            password="BankuInverted1994",      # Your MySQL password
            database="kunalTesting" # The name of your database
        )
        return mydb
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return None

def get_ist_time():
    """Get the current time in IST."""
    ist = pytz.timezone('Asia/Kolkata')  # IST timezone
    utc_time = datetime.now(timezone.utc)  # Get UTC time in a timezone-aware format
    ist_time = utc_time.astimezone(ist)  # Convert UTC to IST
    return ist_time

def generate_date_format(serial_number):
    """Generate the date format with serial number appended."""
    today = get_ist_time()  # Use IST time

    day = today.day  # Extract the day
    year = chr(96 + (today.year % 100))  # Convert year to alphabetic format (2024 -> 'x' for '24')
    month = chr(96 + today.month)  # Convert month to alphabetic format (1 -> 'a', ..., 12 -> 'l')
    
    formatted_date = f"{day}{year}{month}{serial_number:03d}"  # Format the date with the serial number
    return formatted_date

def display_bms_options():
    """Display available BMS specifications from the database."""
    # Connect to the MySQL database
    mydb = connect_to_database()
    if mydb is None:
        print("Failed to connect to the database.")
        return []

    cursor = mydb.cursor()
    try:
        # Query to get the distinct BMS types from the BmsSpecification table 
        query = "SELECT bmstype FROM BmsSpecification"
        cursor.execute(query)
        bms_specs = cursor.fetchall()  # Fetch all results

        # Check if any BMS types were retrieved
        if not bms_specs:
            print("No BMS specifications found in the database.")
            return []

        # Display the available BMS specifications
        print("\nPlease select a BMS specification from the following list:")
        for index, (spec,) in enumerate(bms_specs, 1):  # Unpack the tuple returned by fetchall
            print(f"{index}. {spec}")

        # Return the list of BMS specifications for further use
        return [spec for (spec,) in bms_specs]  # List of BMS types as a simple list of strings

    except mysql.connector.Error as err:
        print(f"Error fetching BMS specifications: {err}")
        return []
    finally:
        cursor.close()
        mydb.close()

def upload_barcode_to_db(bmsBarcode, product_description):
    """Upload the generated bmsBarcode to the MySQL database."""
    # Connect to the MySQL database
    mydb = connect_to_database()
    if mydb is None:
        print("Failed to connect to the database.")
        return

    cursor = mydb.cursor()
    try:
        # Get IST time for the 'created_at' field
        ist_time = get_ist_time()
        # Format the IST time for database insertion
        ist_time_str = ist_time.strftime('%Y-%m-%d %H:%M:%S')
        
        # Product name is always 'Dally BMS'
        # product_name = 'Dally BMS'

        # Insert the generated bmsBarcode into the bmsBarcode table with the IST timestamp and other fields
        query = """INSERT INTO bmsBarcode (barcode_value, created_at, product_name, product_description)
                   VALUES (%s, %s, %s, %s)"""
        # cursor.execute(query, (bmsBarcode, ist_time_str, product_name, product_description))
        cursor.execute(query, (bmsBarcode, ist_time_str, product_description))
        mydb.commit()  # Commit the transaction to save changes
        print(f"bmsBarcode {bmsBarcode} has been uploaded to the database.")
    except mysql.connector.Error as err:
        print(f"Error uploading bmsBarcode to the database: {err}")
    finally:
        cursor.close()
        mydb.close()

def check_and_generate_barcode(bms_spec):
    """Check if a bmsBarcode for the current date exists and generate the next bmsBarcode."""
    # Connect to the MySQL database
    mydb = connect_to_database()
    if mydb is None:
        return None, None
    
    cursor = mydb.cursor()
    today_date = get_ist_time().strftime('%Y-%m-%d')  # Get today's date in 'YYYY-MM-DD' format in IST
    
    # Query to check the last bmsBarcode for the current date
    query = "SELECT barcode_value FROM bmsBarcode WHERE DATE(created_at) = %s AND barcode_value LIKE %s ORDER BY barcode_value DESC LIMIT 1"
    cursor.execute(query, (today_date, f"{bms_spec}%"))
    
    result = cursor.fetchone()
    cursor.close()
    mydb.close()

    if result:
        # bmsBarcode for today exists, extract the serial number from the last bmsBarcode
        last_barcode = result[0]
        # Extract the serial number from the last bmsBarcode (assuming it's always 3 digits at the end)
        last_serial_number = int(last_barcode[-3:])
        new_serial_number = last_serial_number + 1
        return new_serial_number, last_barcode
    else:
        # No bmsBarcode exists for today, start from 001
        return 1, None

def get_product_description(bms_spec):
    """Return the product description based on the BMS specification from the database."""
    # Connect to the MySQL database
    mydb = connect_to_database()
    if mydb is None:
        print("Failed to connect to the database.")
        return 'Unknown BMS Spec'  # Return a default value if no connection is made

    cursor = mydb.cursor()

    try:
        # Query to fetch the product description based on the bms_spec
        query = "SELECT product_description FROM BmsSpecification WHERE bmstype = %s"
        cursor.execute(query, (bms_spec,))
        result = cursor.fetchone()

        # Check if a result was returned
        if result:
            return result[0]  # The product description is in the first column of the result
        else:
            return 'Unknown BMS Spec'  # Return default if no matching specification is found

    except mysql.connector.Error as err:
        print(f"Error fetching product description: {err}")
        return 'Unknown BMS Spec'
    finally:
        cursor.close()
        mydb.close()

def main():
    """Main function to run the bmsBarcode generation process."""
    while True:
            bms_specs = display_bms_options()

            if not bms_specs:
                print("No BMS specifications available. Exiting process.")
                break  # Exit if no BMS specs are found in the database

            try:
                user_choice = int(input("\nEnter the number corresponding to your selection (1-{}) or '0' to quit: ".format(len(bms_specs))))
                if user_choice == 0:
                    print("Exiting bmsBarcode generation process.")
                    break  # Exit the loop if the user chooses to stop
                elif 1 <= user_choice <= len(bms_specs):
                    selected_spec = bms_specs[user_choice - 1]
                    print(f"\nYou selected the BMS specification: {selected_spec}")
                    
                    # Get the product description for the selected BMS specification
                    product_description = get_product_description(selected_spec)

                    # Check if a bmsBarcode exists for today, and get the next serial number
                    new_serial_number, last_barcode = check_and_generate_barcode(selected_spec)
                    
                    # If no bmsBarcode exists, start from 001
                    if new_serial_number == 1:
                        print(f"Starting new serial numbers for today.")
                    else:
                        print(f"Last bmsBarcode for today: {last_barcode}, generating new bmsBarcode with serial number {new_serial_number:03d}")
                    
                    # Generate the bmsBarcode with the new serial number
                    date_string = generate_date_format(new_serial_number)
                    bmsBarcode = f"{selected_spec}{date_string}"  # Concatenate the BMS spec with the formatted date
                    print(f"Generated bmsBarcode: {bmsBarcode}")
                    
                    # Upload the bmsBarcode to the database with the product description
                    upload_barcode_to_db(bmsBarcode, product_description)
                else:
                    print(f"Invalid choice. Please select a number between 1 and {len(bms_specs)}.")
            
            except ValueError:
                print("Invalid input. Please enter a valid number.")

if __name__ == "__main__":
    main()