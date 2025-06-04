import os
import time
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage


def setup_driver():
    """Setup Edge driver"""
    edge_options = Options()
    # Keep browser visible for manual login
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument("--disable-dev-shm-usage")
    edge_options.add_argument("--window-size=1920,1080")
    
    service = Service(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=edge_options)
    return driver

def manual_authentication(driver):
    """Let user authenticate manually and confirm when ready"""
    print("=== MANUAL AUTHENTICATION ===")
    print("1. A browser window will open")
    print("2. Please log in to St Andrews manually")
    print("3. Navigate to any module page to confirm you're logged in")
    print("4. Come back here and press Enter when authentication is complete")
    print()
    
    # Open the login page
    test_url = "https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/GG3214/Final+grade/GraphPage"
    print(f"Opening: {test_url}")
    driver.get(test_url)
    
    # Wait for user to complete authentication
    input("Press Enter after you have successfully logged in and can see the module page...")
    
    # Verify we're on the right page
    try:
        # Check if we can find elements that indicate we're logged in
        wait = WebDriverWait(driver, 5)
        
        # Look for page elements that confirm we're authenticated
        if "login" in driver.current_url.lower() or "auth" in driver.current_url.lower():
            print("Warning: Still appears to be on login page")
            input("Please complete login and press Enter again...")
        
        print("Authentication verified! Starting downloads...")
        return True
        
    except Exception as e:
        print(f"Note: {e}")
        print("Proceeding anyway...")
        return True

def save_charts_as_png(driver, url, module_code):
    """Navigate to URL and save scatterChart and barChart as PNG"""
    try:
        print(f"Loading {module_code}: {url}")
        driver.get(url)
        time.sleep(5)  # Wait for page to load
        
        # Check if we got redirected to login (shouldn't happen if session is valid)
        if "login" in driver.current_url.lower() or "auth" in driver.current_url.lower():
            print(f"Session expired! Please re-authenticate.")
            input("Complete authentication and press Enter...")
            driver.get(url)
            time.sleep(5)
        
        # Look specifically for scatterChart and barChart
        try:
            wait = WebDriverWait(driver, 15)
            
            # Find scatter chart
            scatter_chart = None
            bar_chart = None
            
            try:
                # scatter_chart = wait.until(EC.presence_of_element_located((By.ID, "scatterChart")))
                scatter_chart = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#scatterChart .user-select-none.svg-container")))
                print(f"  Found scatterChart for {module_code}")
            except:
                print(f"  scatterChart not found for {module_code}")
            
            try:
                bar_chart = driver.find_element(By.ID, "barChart")
                print(f"  Found barChart for {module_code}")
            except:
                print(f"  barChart not found for {module_code}")
            
            # Give charts extra time to fully render
            time.sleep(5)
            
            charts_found = []
            if scatter_chart:
                charts_found.append(("scatterChart", scatter_chart))
            if bar_chart:
                charts_found.append(("barChart", bar_chart))
            
            if not charts_found:
                print(f"No charts found for {module_code}")
                return 0
            
            print(f"Found {len(charts_found)} charts for {module_code}")
            
        except Exception as e:
            print(f"Error finding charts for {module_code}: {e}")
            return 0
        
        # Create folder for this module
        folder_path = os.path.join("charts", module_code)
        os.makedirs(folder_path, exist_ok=True)
        
        saved_count = 0
        
        # Save each chart with specific naming
        for i, (chart_type, chart_element) in enumerate(charts_found):
            try:
                # Scroll to chart to ensure it's visible
                driver.execute_script("arguments[0].scrollIntoView(true);", chart_element)
                time.sleep(2)
                
                # Take screenshot of the chart
                filename = os.path.join(folder_path, f"Chart_{i+1}.png")
                chart_element.screenshot(filename)
                
                print(f"  ✓ Saved Chart_{i+1}.png ({chart_type})")
                saved_count += 1
                
            except Exception as e:
                print(f"  ✗ Failed to save {chart_type}: {e}")
        
        print(f"Saved {saved_count} charts for {module_code}")
        return saved_count
        
    except Exception as e:
        print(f"Error processing {module_code}: {e}")
        return 0

def download_all_charts():
    """Main function - manual auth then download all charts"""
    
    print("St Andrews Module Charts Downloader")
    print("=" * 50)
    
    # Module codes to process
   
    module_codes = ['GG4258', 'GG3281', 'GG1002', 'GG2014', 'GG4248', 'GG4247', 'SS5103',
    'GG4254', 'GG4257', 'GG3205', 'GG3213', 'GG3214', 'GG5005', 'GG4399',
    'SD4126', 'SD4129', 'SD4133', 'SD1004', 'SD4225', 'SD2006', 'SD2100',
    'SD4110', 'SD3102', 'SD3101', 'SD4120', 'SD4125', 'SD4297', 'SD5801',
    'SD5802', 'SD5805', 'SD5806', 'SD5807', 'SD5810', 'SD5820', 'SD5821',
    'SD5811', 'SD5813', 'SD5812']
    
    base_url = "https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/{}/Final+grade/GraphPage"
    
    # Create main charts folder
    os.makedirs("charts", exist_ok=True)
    
    # Setup browser
    print("Setting up browser...")
    driver = setup_driver()
    
    try:
        # Step 1: Manual authentication
        if not manual_authentication(driver):
            print("Authentication failed. Exiting...")
            return
        
        # Step 2: Download charts from all modules
        print(f"\nProcessing {len(module_codes)} modules...")
        print("=" * 30)
        
        total_saved = 0
        successful_modules = 0
        
        for i, module_code in enumerate(module_codes, 1):
            print(f"\n[{i}/{len(module_codes)}] Processing {module_code}...")
            
            url = base_url.format(module_code)
            saved = save_charts_as_png(driver, url, module_code)
            
            if saved > 0:
                successful_modules += 1
                total_saved += saved
            
            # Small delay between modules
            time.sleep(2)
        
        # Final summary
        print("\n" + "=" * 50)
        print("DOWNLOAD COMPLETE!")
        print(f"Processed: {len(module_codes)} modules")
        print(f"Successful: {successful_modules} modules")
        print(f"Total charts saved: {total_saved}")
        print(f"Charts saved in: ./charts/")
        print("=" * 50)
        
        
        # List what was downloaded
        print("\nDownloaded charts by module:")
        for module_code in module_codes:
            folder_path = os.path.join("charts", module_code)
            if os.path.exists(folder_path):
                files = [f for f in os.listdir(folder_path) if f.endswith('.png')]
                if files:
                    print(f"  {module_code}: {len(files)} charts")
        
        generate_excel_from_charts("charts", "ModuleCharts.xlsx")
        
    except KeyboardInterrupt:
        print("\nDownload interrupted by user")
    except Exception as e:
        print(f"Unexpected error: {e}")
    finally:
        input("\nPress Enter to close the browser...")
        driver.quit()
        

def generate_excel_from_charts(charts_dir="charts", output_file="ModuleCharts.xlsx"):
    """Creates an Excel file with one sheet per module, embedding saved PNG charts."""
    print("\nGenerating Excel file with charts...")
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    for module_code in os.listdir(charts_dir):
        module_path = os.path.join(charts_dir, module_code)
        if not os.path.isdir(module_path):
            continue

        sheet = wb.create_sheet(title=module_code)
        row_pos = 1

        # Sort to maintain order (e.g., Chart_1.png, Chart_2.png)
        for chart_file in sorted(os.listdir(module_path)):
            if not chart_file.lower().endswith('.png'):
                continue

            chart_path = os.path.join(module_path, chart_file)

            try:
                # Resize to avoid huge scaling in Excel
                img = PILImage.open(chart_path)
                img.thumbnail((600, 400))  # Resize to max dimensions
                temp_path = chart_path.replace(".png", "_resized.png")
                img.save(temp_path)

                xl_img = XLImage(temp_path)
                cell_location = f"A{row_pos}"
                sheet.add_image(xl_img, cell_location)
                row_pos += 20  # space between charts

                print(f"  Added {chart_file} to sheet {module_code}")
            except Exception as e:
                print(f"  ✗ Failed to add {chart_file} to sheet {module_code}: {e}")

    wb.save(output_file)
    print(f"\n✓ Excel file created: {output_file}")

if __name__ == "__main__":
    download_all_charts()
