import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchFrameException
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import logging
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import queue
import platform

# Set up logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S',
                    filename='bscscan_automation.log')
logger = logging.getLogger(__name__)
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter('%(message)s')
console.setFormatter(formatter)
logger.addHandler(console)

# Resource path for finding files when using PyInstaller
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_sheet_data(spreadsheet_url, start_row, end_row):
    """ Get data from Google Sheet """
    logger.info("Connecting to Google Sheets...")
    try:
        # Define scope
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

        # Authenticate using service account
        creds_path = resource_path('service_account.json')
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
        client = gspread.authorize(creds)

        # Open the sheet by URL
        sheet = client.open_by_url(spreadsheet_url).sheet1

        # Read the desired rows (columns 1 and 2 only)
        data = []
        rows = sheet.get_all_values()[start_row-1:end_row]
        for row in rows:
            # Check if row has at least 2 values
            if len(row) >= 2 and row[0] and row[1]:
                data.append([row[0], row[1]])  # Only take first two columns

        logger.info(f"Successfully retrieved {len(data)} rows of data")
        return data
    except Exception as e:
        logger.error(f"Error retrieving data from spreadsheet: {e}")
        return None

class BscscanApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BSCScan Form Automation")
        self.root.geometry("650x650")
        self.root.resizable(True, True)
        
        # For communication between threads
        self.continue_event = threading.Event()
        self.driver = None
        self.automation_thread = None
        self.current_row_index = 0
        
        # Set icon if available - use different approach on macOS
        self.set_app_icon()
        
        self.create_widgets()
    
    def set_app_icon(self):
        """Set application icon with cross-platform support"""
        if platform.system() == 'Darwin':  # macOS
            try:
                # On macOS, the icon is set in the app bundle
                pass
            except:
                pass
        else:  # Windows/Linux
            try:
                self.root.iconbitmap(resource_path("icon.ico"))
            except:
                pass
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="BSCScan Contract Automation", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)
        
        # Input fields frame
        input_frame = ttk.LabelFrame(main_frame, text="Input Settings", padding="10")
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Google Sheet URL
        ttk.Label(input_frame, text="Google Sheets URL:").grid(column=0, row=0, sticky=tk.W, pady=5)
        self.sheet_url = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.sheet_url, width=50).grid(column=1, row=0, padx=5, pady=5, sticky=tk.W)
        
        # Row range
        ttk.Label(input_frame, text="Start Row:").grid(column=0, row=1, sticky=tk.W, pady=5)
        self.start_row = tk.StringVar(value="2")
        ttk.Entry(input_frame, textvariable=self.start_row, width=10).grid(column=1, row=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(input_frame, text="End Row:").grid(column=0, row=2, sticky=tk.W, pady=5)
        self.end_row = tk.StringVar(value="10")
        ttk.Entry(input_frame, textvariable=self.end_row, width=10).grid(column=1, row=2, padx=5, pady=5, sticky=tk.W)
        
        # Contract URL (fixed but visible)
        ttk.Label(input_frame, text="Contract URL:").grid(column=0, row=3, sticky=tk.W, pady=5)
        self.contract_url = tk.StringVar(value="https://bscscan.com/token/0xBD576D184f5843881e471f9292036a076CB532b0#writeContract")
        url_entry = ttk.Entry(input_frame, textvariable=self.contract_url, width=50)
        url_entry.grid(column=1, row=3, padx=5, pady=5, sticky=tk.W)
        url_entry.config(state="readonly")
        
        # Service account file
        ttk.Label(input_frame, text="Service Account JSON:").grid(column=0, row=4, sticky=tk.W, pady=5)
        self.service_account_path = tk.StringVar()
        service_frame = ttk.Frame(input_frame)
        service_frame.grid(column=1, row=4, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(service_frame, textvariable=self.service_account_path, width=40).pack(side=tk.LEFT)
        ttk.Button(service_frame, text="Browse", command=self.browse_service_account).pack(side=tk.LEFT, padx=5)
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.preview_btn = ttk.Button(button_frame, text="Preview Data", command=self.preview_data)
        self.preview_btn.pack(side=tk.LEFT, padx=5)
        
        self.start_btn = ttk.Button(button_frame, text="Start Automation", command=self.start_automation)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.continue_btn = ttk.Button(button_frame, text="Continue", command=self.continue_action, state=tk.DISABLED)
        self.continue_btn.pack(side=tk.LEFT, padx=5)
        
        # Status area
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Status log with scrollbar
        self.status_text = tk.Text(status_frame, height=15, wrap=tk.WORD)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(status_frame, command=self.status_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, 
                                       length=100, mode='determinate', 
                                       variable=self.progress_var)
        self.progress.pack(fill=tk.X, padx=5, pady=5)
        
        # Initial status message
        self.update_status("Ready. Please enter Google Sheet URL and row range.")
    
    def browse_service_account(self):
        filename = filedialog.askopenfilename(
            title="Select Service Account JSON file",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.service_account_path.set(filename)
    
    def update_status(self, message):
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def preview_data(self):
        try:
            # Save the service account file if it's provided
            if self.service_account_path.get():
                import shutil
                shutil.copy(self.service_account_path.get(), resource_path('service_account.json'))
            
            # Validate input
            if not self.sheet_url.get():
                messagebox.showerror("Error", "Please enter a Google Sheet URL")
                return
            
            try:
                start = int(self.start_row.get())
                end = int(self.end_row.get())
                if start < 1 or end < start:
                    raise ValueError()
            except ValueError:
                messagebox.showerror("Error", "Invalid row numbers")
                return
            
            self.update_status("Fetching data from Google Sheet...")
            self.preview_btn.config(state=tk.DISABLED)
            
            # Get data in a separate thread
            def fetch_data():
                data = get_sheet_data(self.sheet_url.get(), int(self.start_row.get()), int(self.end_row.get()))
                if data:
                    self.root.after(0, lambda: self.show_preview(data))
                else:
                    self.root.after(0, lambda: self.update_status("Error: Could not retrieve data from sheet"))
                    self.root.after(0, lambda: self.preview_btn.config(state=tk.NORMAL))
            
            threading.Thread(target=fetch_data).start()
        except Exception as e:
            self.update_status(f"Error previewing data: {e}")
            self.preview_btn.config(state=tk.NORMAL)
    
    def show_preview(self, data):
        self.preview_btn.config(state=tk.NORMAL)
        
        # Create preview window
        preview = tk.Toplevel(self.root)
        preview.title("Data Preview")
        preview.geometry("500x400")
        
        # Create a Treeview widget
        columns = ('Address', 'Value')
        tree = ttk.Treeview(preview, columns=columns, show='headings')
        tree.heading('Address', text='Address')
        tree.heading('Value', text='Value')
        
        # Configure column widths
        tree.column('Address', width=300)
        tree.column('Value', width=150)
        
        # Add data to the treeview
        for i, row in enumerate(data):
            tree.insert('', tk.END, values=row, tags=('even' if i % 2 == 0 else 'odd',))
        
        # Add alternating row colors
        tree.tag_configure('odd', background='#f0f0f0')
        tree.tag_configure('even', background='white')
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(preview, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        
        # Pack the widgets
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        
        # Add a label with the total
        ttk.Label(preview, text=f"Total rows: {len(data)}").pack(pady=5)
        
        # Add a close button
        ttk.Button(preview, text="Close", command=preview.destroy).pack(pady=10)
        
        self.update_status(f"Retrieved {len(data)} rows of data")
    
    def start_automation(self):
        try:
            # Save the service account file if it's provided
            if self.service_account_path.get():
                import shutil
                shutil.copy(self.service_account_path.get(), resource_path('service_account.json'))
            
            # Validate input
            if not self.sheet_url.get():
                messagebox.showerror("Error", "Please enter a Google Sheet URL")
                return
            
            try:
                start = int(self.start_row.get())
                end = int(self.end_row.get())
                if start < 1 or end < start:
                    raise ValueError()
            except ValueError:
                messagebox.showerror("Error", "Invalid row numbers")
                return
            
            # Disable buttons and enable continue button
            self.start_btn.config(state=tk.DISABLED)
            self.preview_btn.config(state=tk.DISABLED)
            
            self.update_status("Starting automation process...")
            
            # Get data
            data = get_sheet_data(self.sheet_url.get(), int(self.start_row.get()), int(self.end_row.get()))
            if not data:
                self.update_status("Error: Could not retrieve data from sheet")
                self.start_btn.config(state=tk.NORMAL)
                self.preview_btn.config(state=tk.NORMAL)
                return
            
            self.data = data
            self.total_rows = len(data)
            self.current_row_index = 0
            
            # Reset the continue event
            self.continue_event.clear()
            
            # Start the automation in a separate thread
            self.automation_thread = threading.Thread(target=self.run_automation)
            self.automation_thread.daemon = True
            self.automation_thread.start()
        except Exception as e:
            self.update_status(f"Error starting automation: {e}")
            self.start_btn.config(state=tk.NORMAL)
            self.preview_btn.config(state=tk.NORMAL)
            self.continue_btn.config(state=tk.DISABLED)
    
    def run_automation(self):
        try:
            url = self.contract_url.get()
            self.update_status(f"Opening browser to {url}")
            
            # Launch browser with macOS-specific options if needed
            options = uc.ChromeOptions()
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--window-size=1920,1080")
            
            # Add macOS-specific options
            if platform.system() == 'Darwin':
                # Check if running from app bundle and adjust paths accordingly
                if getattr(sys, 'frozen', False):
                    # For macOS, we might need to specify the Chrome binary location
                    # if it's having trouble finding it
                    chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
                    if os.path.exists(chrome_path):
                        options.binary_location = chrome_path
            
            self.driver = uc.Chrome(options=options)
            self.driver.get(url)
            
            # Wait for initial page load
            try:
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                self.update_status("Page loaded initially")
            except TimeoutException:
                self.update_status("Error: Initial page load timed out")
                self.driver.quit()
                return
            
            # Enable the continue button to let user signal completion of any verification
            self.root.after(0, lambda: self.update_status("IMPORTANT: Complete any CAPTCHA or verification if present"))
            self.root.after(0, lambda: self.update_status("Click 'Continue' in the app once the page is fully loaded"))
            self.root.after(0, lambda: self.continue_btn.config(state=tk.NORMAL))
            
            # Wait for user to click Continue
            self.update_status("Waiting for you to click Continue...")
            self.continue_event.wait()  # This blocks until continue_action() sets the event
            self.continue_event.clear()  # Reset the event for the next wait
            
            self.update_status("Continuing with form filling...")
            
            # Process data rows one by one, waiting for user confirmation between rows
            self.process_next_row()
            
        except Exception as e:
            self.update_status(f"Error during automation: {e}")
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.preview_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.continue_btn.config(state=tk.DISABLED))
    
    def process_next_row(self):
        """Process the next row and then wait for user confirmation"""
        try:
            # Check if we're done
            if self.current_row_index >= len(self.data):
                self.update_status("All rows processed!")
                messagebox.showinfo("Complete", "All rows have been processed.")
                
                # Ask before closing browser
                if messagebox.askyesno("Close Browser", "Close the browser?"):
                    if self.driver:
                        self.driver.quit()
                    self.update_status("Browser closed")
                else:
                    self.update_status("Browser left open. Close it manually when finished.")
                
                # Re-enable buttons when done
                self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.preview_btn.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.continue_btn.config(state=tk.DISABLED))
                return
            
            # Get current row data
            row_data = self.data[self.current_row_index]
            self.update_status(f"Processing row {self.current_row_index+1} of {len(self.data)}: {row_data}")
            
            try:
                # Switch to the iframe containing contract elements
                try:
                    # First switch back to main content just to be safe
                    self.driver.switch_to.default_content()
                    
                    # Wait for the iframe to be present
                    iframe = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, "writecontractiframe"))
                    )
                    self.update_status("Found contract iframe")
                    
                    # Switch to the iframe
                    self.driver.switch_to.frame("writecontractiframe")
                    self.update_status("Switched to contract iframe")
                except (TimeoutException, NoSuchFrameException) as e:
                    self.update_status(f"Error switching to iframe: {e}")
                    if messagebox.askretry("Error", "Failed to find contract iframe. Retry?"):
                        # Try again
                        self.process_next_row()
                        return
                    else:
                        # Skip to next row
                        self.current_row_index += 1
                        self.process_next_row()
                        return
                
                # Find and click the sendLockByAdmin dropdown
                try:
                    dropdown_selector = "a[href='#collapse6'][data-bs-toggle='collapse']"
                    dropdown = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, dropdown_selector))
                    )
                    
                    # Scroll to the element to ensure it's visible
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", dropdown)
                    time.sleep(1)  # Allow time for scrolling
                    
                    # Check if dropdown is already expanded
                    is_expanded = "collapsed" not in dropdown.get_attribute("class")
                    
                    if not is_expanded:
                        # Try multiple approaches to click the element
                        try:
                            dropdown.click()
                        except Exception:
                            try:
                                self.driver.execute_script("arguments[0].click();", dropdown)
                            except Exception:
                                from selenium.webdriver.common.action_chains import ActionChains
                                ActionChains(self.driver).move_to_element(dropdown).click().perform()
                        
                        self.update_status("Clicked on sendLockByAdmin dropdown")
                        time.sleep(2)
                    else:
                        self.update_status("sendLockByAdmin dropdown already expanded")
                except Exception as e:
                    self.update_status(f"Error finding or clicking dropdown: {e}")
                    if messagebox.askretry("Error", f"Error accessing dropdown. Retry row {self.current_row_index+1}?"):
                        # Try again with same row
                        self.process_next_row()
                        return
                    else:
                        # Skip to next row
                        self.current_row_index += 1
                        self.process_next_row()
                        return
                
                # Fill in the form fields
                try:
                    # Set types field to 2
                    types_field = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "input_6_1"))
                    )
                    types_field.clear()
                    types_field.send_keys("2")
                    self.update_status("Entered '2' in types field")
                    
                    # Set address field (first column from spreadsheet)
                    address_field = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "input_6_2"))
                    )
                    address_field.clear()
                    address_field.send_keys(row_data[0])
                    self.update_status(f"Entered address: {row_data[0]}")
                    
                    # Set value field (second column from spreadsheet)
                    value_field = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "input_6_3"))
                    )
                    value_field.clear()
                    value_in_wei = int(float(row_data[1]) * 1e18)  # Convert to integer wei
                    value_field.send_keys(str(value_in_wei))       # Send as string
                    self.update_status(f"Entered value: {row_data[1]}")
                    
                    self.update_status(f"✓ Row {self.current_row_index+1} filled successfully")
                except Exception as e:
                    self.update_status(f"Error filling fields: {e}")
                    if messagebox.askretry("Error", f"Error filling fields. Retry row {self.current_row_index+1}?"):
                        # Try again with same row
                        self.process_next_row()
                        return
                    else:
                        # Skip to next row
                        self.current_row_index += 1
                        self.process_next_row()
                        return
                
                # Save references to fields for clearing after user continues
                self.types_field = types_field
                self.address_field = address_field
                self.value_field = value_field
                
                # Increment row counter for next iteration
                self.current_row_index += 1
                
                # Update progress bar
                progress_value = (self.current_row_index / self.total_rows) * 100
                self.root.after(0, lambda: self.progress_var.set(progress_value))
                
                # Ask user to continue to next row
                if self.current_row_index < len(self.data):
                    self.update_status(f"Row {self.current_row_index} complete. Click Continue to process next row.")
                    self.root.after(0, lambda: self.continue_btn.config(state=tk.NORMAL))
                    self.continue_event.wait()
                    self.continue_event.clear()
                    self.root.after(0, lambda: self.continue_btn.config(state=tk.DISABLED))
                    
                    # Clear fields AFTER user clicks continue button
                    try:
                        self.types_field.clear()
                        self.address_field.clear()
                        self.value_field.clear()
                        self.update_status("Cleared fields for next row")
                    except Exception as e:
                        self.update_status(f"Error clearing fields: {e}")
                    
                    # Process next row
                    self.process_next_row()
                else:
                    # All done
                    self.update_status("All rows processed!")
                    messagebox.showinfo("Complete", "All rows have been processed.")
                    
                    # Ask before closing browser
                    if messagebox.askyesno("Close Browser", "Close the browser?"):
                        if self.driver:
                            self.driver.quit()
                        self.update_status("Browser closed")
                    else:
                        self.update_status("Browser left open. Close it manually when finished.")
                    
                    # Re-enable buttons when done
                    self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.preview_btn.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.continue_btn.config(state=tk.DISABLED))
                
            except Exception as e:
                self.update_status(f"Unexpected error: {e}")
                if messagebox.askretry("Error", f"Unexpected error. Retry row {self.current_row_index+1}?"):
                    # Try again with same row
                    self.process_next_row()
                else:
                    # Skip to next row
                    self.current_row_index += 1
                    self.process_next_row()
            
        except Exception as e:
            self.update_status(f"Error processing row: {e}")
            # Re-enable buttons
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.preview_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.continue_btn.config(state=tk.DISABLED))
    
    def continue_action(self):
        """User clicked continue button - signal the automation thread to continue"""
        self.update_status("Continue button clicked. Proceeding...")
        self.continue_btn.config(state=tk.DISABLED)
        self.continue_event.set()  # Signal the automation thread to continue

def main():
    # Configure macOS app name in menu bar (macOS specific)
    if platform.system() == 'Darwin':
        # This makes the app name appear in the menu bar
        os.environ['PYTHONAPPSKEEPPATH'] = '1'
        
    root = tk.Tk()
    app = BscscanApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()