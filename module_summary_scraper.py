import time
import os
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from openpyxl import Workbook


def setup_driver():
    """Setup Microsoft Edge WebDriver with visible browser window."""
    edge_options = Options()
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument("--disable-dev-shm-usage")
    edge_options.add_argument("--window-size=1920,1080")
    service = Service(EdgeChromiumDriverManager().install())
    return webdriver.Edge(service=service, options=edge_options)

def manual_authentication(driver):
    """Prompt user to login manually and verify authentication."""
    print("=== MANUAL AUTHENTICATION ===")
    print("1. A browser window will open.")
    print("2. Log in to MMS manually.")
    print("3. Navigate to any module page to verify login.")
    input("Press Enter once you're logged in...")

    if "login" in driver.current_url.lower() or "auth" in driver.current_url.lower():
        print("Still on login page — make sure you're logged in.")
        input("Press Enter again after logging in...")

    print("Authentication complete.\n")
    return True

def extract_summary_stats(driver, url, module_code):
    """Extract summary stats (Count, Mean, Std. Dev.) from the table footer of a module."""
    print(f"Processing module {module_code}...")
    driver.get(url)
    time.sleep(3)

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "gradesTable"))
        )
        tfoot = driver.find_element(By.CSS_SELECTOR, "#gradesTable tfoot")
        cells = tfoot.find_elements(By.TAG_NAME, "td")
        values = [cell.text.strip() for cell in cells if cell.text.strip()]
        return [module_code] + values
    except Exception as e:
        print(f"Failed to extract summary for {module_code}: {e}")
        return [module_code, "ERROR"]

def save_to_excel(data, filename="ModuleSummaries_2023_4.xlsx"):
    """Save summary statistics to a single Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    for row in data:
        ws.append(row)

    wb.save(filename)
    print(f"\n✓ Summary saved to Excel: {filename}")

def run_summary_scraper():
    """Main function to authenticate, visit each module, extract stats and write to Excel."""
    
    module_codes = ['GG4258', 'GG3281', 'GG1002', 'GG2014', 'GG4248', 'GG4247', 'SS5103',
    'GG4254', 'GG4257', 'GG3205', 'GG3213', 'GG3214', 'GG5005', 'GG4399',
    'SD4126', 'SD4129', 'SD4133', 'SD1004', 'SD4225', 'SD2006', 'SD2100',
    'SD4110', 'SD3102', 'SD3101', 'SD4120', 'SD4125', 'SD4297', 'SD5801',
    'SD5802', 'SD5805', 'SD5806', 'SD5807', 'SD5810', 'SD5820', 'SD5821',
    'SD5811', 'SD5813', 'SD5812']
    
    
    base_url = "https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/{}/Final+grade/" URL for 2025-2 semestre (current year)
    #base_url = "https://mms.st-andrews.ac.uk/mms/module/2023_4/S2/{}/Final+grade/" #URL for 2023-4 (previous year)
    
    summary_data = [["Module", "Count", "Mean", "Std. Dev."]]

    driver = setup_driver()

    try:
        # Authenticate manually
        test_url = base_url.format("GG1002")
        driver.get(test_url)
        if not manual_authentication(driver):
            print("Authentication failed.")
            return

        for code in module_codes:
            url = base_url.format(code)
            row = extract_summary_stats(driver, url, code)
            summary_data.append(row)
            time.sleep(1)

        save_to_excel(summary_data)

    finally:
        input("Press Enter to close browser...")
        driver.quit()

if __name__ == "__main__":
    run_summary_scraper()