import time
import logging
from sheet_reader import get_sheet_data
from bscscan_filler import fill_bscscan_contract

# Set up logging
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger(__name__)

def auto_submit():
    # Get user inputs
    spreadsheet_url = input("Enter Google Sheets URL: ")
    start_row = int(input("Enter start row (e.g., 2): "))
    end_row = int(input("Enter end row (e.g., 10): "))
    
    # Fixed columns (first and second columns)
    columns = [1, 2]
    
    # Fixed URL for BscScan contract
    website_url = "https://bscscan.com/token/0xBD576D184f5843881e471f9292036a076CB532b0#writeContract"
    
    logger.info(f"Starting process with spreadsheet: {spreadsheet_url}")
    logger.info(f"Reading rows {start_row} to {end_row}, columns {columns}")
    
    # Step 1: Get the data from the Google Sheet
    try:
        data = get_sheet_data(spreadsheet_url, start_row, end_row, columns)
        logger.info(f"Successfully retrieved {len(data)} rows of data")
        
        if not data:
            logger.warning("No data found in spreadsheet. Please check your inputs.")
            return
    except Exception as e:
        logger.error(f"Error retrieving data from spreadsheet: {e}")
        return
    
    # Print preview of data
    logger.info("Data preview:")
    for i, row in enumerate(data[:3]):  # Show first 3 rows as preview
        logger.info(f"Row {i+start_row}: {row}")
    
    if len(data) > 3:
        logger.info(f"... and {len(data)-3} more rows")
    
    confirm = input("Data retrieved. Proceed with form filling? (y/n): ")
    if confirm.lower() != 'y':
        logger.info("Process cancelled by user")
        return

    # Step 2: Fill the BSCScan contract form with each row of data
    fill_bscscan_contract(website_url, data)

    logger.info("Process completed!")

# Run the script
if __name__ == "__main__":
    auto_submit()