import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchFrameException
import time
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger(__name__)

def fill_bscscan_contract(website_url, data_rows):
    """
    Fills BSCScan contract form with spreadsheet data.
    Handles iframe and potential CAPTCHA/verification challenges.
    
    Args:
        website_url: The BSCScan contract URL
        data_rows: List of rows, each containing [address, value] for the contract
    """
    # Use undetected_chromedriver instead of standard selenium
    options = uc.ChromeOptions()
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--window-size=1920,1080")
    
    # Launch browser
    driver = uc.Chrome(options=options)
    driver.get(website_url)
    
    # Wait for initial page load
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        logger.info("Page loaded initially")
    except TimeoutException:
        logger.error("Initial page load timed out")
        driver.quit()
        return
    
    # Check if there's any verification needed
    captcha_check = input("Is there a CAPTCHA or verification prompt? (y/n): ")
    if captcha_check.lower() == 'y':
        logger.info("Waiting for manual CAPTCHA/verification completion...")
        input("Complete the verification manually, then press Enter to continue...")
    
    # Additional wait after verification
    time.sleep(5)
    
    # *** CRITICAL: Switch to the iframe containing contract elements ***
    try:
        # First switch back to main content just to be safe
        driver.switch_to.default_content()
        logger.info("Switched to default content")
        
        # Wait for the iframe to be present
        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "writecontractiframe"))
        )
        logger.info("Found contract iframe")
        
        # Switch to the iframe
        driver.switch_to.frame("writecontractiframe")
        logger.info("Switched to contract iframe")
        
        # Quick check to ensure we're in the correct frame
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.card-header"))
            )
            logger.info("Successfully switched to iframe with contract elements")
        except TimeoutException:
            logger.warning("Couldn't find contract elements after switching to iframe")
            input("Please check if the interface looks correct and press Enter to continue...")
    
    except (TimeoutException, NoSuchFrameException) as e:
        logger.error(f"Error switching to iframe: {e}")
        input("Failed to find contract iframe. Please check the page and press Enter to continue or Ctrl+C to exit...")
    
    # Process each row from the spreadsheet
    for row_index, row_data in enumerate(data_rows):
        if len(row_data) < 2 or not row_data[0] or not row_data[1]:
            logger.warning(f"Skipping row {row_index+1}: Invalid data format or empty fields")
            continue
            
        logger.info(f"Processing row {row_index+1} of {len(data_rows)}: {row_data}")
        
        try:
            # 1. Find and click the sendLockByAdmin dropdown
            try:
                # First check if we're still in the iframe
                try:
                    # Quick test if we're still in the iframe
                    driver.find_element(By.CSS_SELECTOR, "div.card-header")
                except:
                    logger.warning("No longer in iframe, switching back...")
                    driver.switch_to.default_content()
                    driver.switch_to.frame("writecontractiframe")
                    logger.info("Switched back to contract iframe")
                
                # Look for the dropdown by ID first (more reliable)
                dropdown_selector = "a[href='#collapse6'][data-bs-toggle='collapse']"
                
                dropdown = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, dropdown_selector))
                )
                
                # Scroll to the element to ensure it's visible
                driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", dropdown)
                time.sleep(1)  # Allow time for scrolling
                
                # Check if dropdown is already expanded
                is_expanded = "collapsed" not in dropdown.get_attribute("class")
                
                if not is_expanded:
                    # Try multiple approaches to click the element
                    try:
                        # First try: standard click
                        dropdown.click()
                    except Exception:
                        try:
                            # Second try: JavaScript click
                            driver.execute_script("arguments[0].click();", dropdown)
                        except Exception:
                            # Third try: click with ActionChains
                            from selenium.webdriver.common.action_chains import ActionChains
                            ActionChains(driver).move_to_element(dropdown).click().perform()
                    
                    logger.info("Clicked on sendLockByAdmin dropdown")
                    # Wait for dropdown to fully expand
                    time.sleep(2)
                    
                    # Verify dropdown expanded properly
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.visibility_of_element_located((By.ID, "input_6_1"))
                        )
                        logger.info("Dropdown expanded successfully")
                    except TimeoutException:
                        logger.warning("Dropdown might not have expanded properly. Trying again...")
                        driver.execute_script("arguments[0].click();", dropdown)
                        time.sleep(2)
                else:
                    logger.info("sendLockByAdmin dropdown already expanded")
                    
            except Exception as e:
                logger.error(f"Error finding or clicking dropdown: {e}")
                reload_option = input("Error accessing dropdown. Reload page? (y/n): ")
                if reload_option.lower() == 'y':
                    driver.refresh()
                    time.sleep(5)
                    
                    # Switch to iframe again after refresh
                    driver.switch_to.default_content()
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "writecontractiframe"))
                    )
                    driver.switch_to.frame("writecontractiframe")
                    
                    # Allow manual handling of any verification
                    captcha_check = input("Is there a CAPTCHA or verification prompt? (y/n): ")
                    if captcha_check.lower() == 'y':
                        input("Complete the verification manually, then press Enter to continue...")
                    
                    row_index -= 1  # Retry this row
                    continue
                else:
                    raise
                
            # 2. Fill in the form fields
            # Set types field to 2
            try:
                types_field = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "input_6_1"))
                )
                types_field.clear()
                types_field.send_keys("2")
                logger.info("Entered '2' in types field")
            except Exception as e:
                logger.error(f"Error setting types field: {e}")
                raise
                
            # Set address field (first column from spreadsheet)
            try:
                address_field = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "input_6_2"))
                )
                address_field.clear()
                address_field.send_keys(row_data[0])
                logger.info(f"Entered address: {row_data[0]}")
            except Exception as e:
                logger.error(f"Error setting address field: {e}")
                raise
                
            # Set value field (second column from spreadsheet)
            try:
                value_field = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "input_6_3"))
                )
                value_field.clear()
                value_in_wei = int(float(row_data[1]) * 1e18)  # Convert to integer wei
                value_field.send_keys(str(value_in_wei))       # Send as string
                logger.info(f"Entered value: {row_data[1]}")
            except Exception as e:
                logger.error(f"Error setting value field: {e}")
                raise
                
            # 3. Pause for user to review before proceeding
            proceed = input(f"Row {row_index+1} filled. Press Enter to continue to next row or 'q' to quit: ")
            if proceed.lower() == 'q':
                logger.info("User requested to quit")
                break
                
            # 4. Clear the fields for the next row
            try:
                types_field.clear()
                address_field.clear()
                value_field.clear()
                logger.info("Cleared fields for next row")
            except Exception as e:
                logger.warning(f"Error clearing fields: {e}")
                # If clearing fails, we'll continue and the new values will overwrite
            
            # Wait briefly before proceeding to next row
            time.sleep(2)
            
        except Exception as e:
            logger.error(f"Error processing row {row_index+1}: {e}")
            retry_option = input("Error occurred. Retry this row? (y/n): ")
            if retry_option.lower() == 'y':
                row_index -= 1  # Reprocess this row
                continue
            
            continue_option = input("Continue to next row? (y/n): ")
            if continue_option.lower() != 'y':
                break
    
    # Ask before closing browser
    logger.info("All rows processed")
    close_option = input("Close browser? (y/n): ")
    if close_option.lower() == 'y':
        driver.quit()
        logger.info("Browser closed")
    else:
        logger.info("Browser left open. Close it manually when finished.")