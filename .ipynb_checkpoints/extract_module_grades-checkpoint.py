import time
import os
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from openpyxl import Workbook


# === Step 1: Setup Edge WebDriver ===
def setup_driver():
    edge_options = Options()
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument("--disable-dev-shm-usage")
    edge_options.add_argument("--window-size=1920,1080")
    service = Service(EdgeChromiumDriverManager().install())
    return webdriver.Edge(service=service, options=edge_options)


# === Step 2: Manual login via browser ===
def manual_authentication(driver):
    print("=== MANUAL AUTHENTICATION ===")
    print("1. A browser will open.")
    print("2. Please login to MMS manually.")
    print("3. Navigate to any module page to verify.")
    print("4. Come back here and press Enter to continue.")
    
    test_url = "https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/GG1002/Final+grade/"
    driver.get(test_url)
    input("Press Enter once you have logged in and see the module page...")
    print("Authentication complete.\n")


# === Step 3: Extract grades from a module ===
def extract_grades_from_module(driver, module_code):
    url = f"https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/{module_code}/Final+grade/"
    print(f"Processing module: {module_code}")
    driver.get(url)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "gradesTable"))
    )

    table = driver.find_element(By.ID, "gradesTable")
    
    # Extract headers from second header row
    headers = table.find_elements(By.CSS_SELECTOR, "thead tr:nth-of-type(2) th")
    header_texts = [h.text.strip() for h in headers]

    # Determine the indices of required columns
    try:
        id_index = header_texts.index("matric")
        grade_index = header_texts.index("calc_grade")
    except ValueError as e:
        print(f"Required columns not found in {module_code}: {e}")
        return []

    # Extract rows from tbody
    rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
    records = []
    for row in rows:
        cols = row.find_elements(By.TAG_NAME, "td")
        if len(cols) > max(id_index, grade_index):
            matric_number = cols[id_index].text.strip()
            calc_grade = cols[grade_index].text.strip()
            records.append((matric_number, calc_grade))

    print(f"  ✓ {len(records)} records extracted.")
    return records


# === Step 4: Save all module data to Excel ===
def save_to_excel(data_dict, filename="ModuleGrades.xlsx"):
    wb = Workbook()
    wb.remove(wb.active)

    for module, data in data_dict.items():
        ws = wb.create_sheet(title=module)
        ws.append(["Matric Number", "Calc Grade"])
        for row in data:
            ws.append(list(row))

    wb.save(filename)
    print(f"\n✓ Excel file saved: {filename}")


# === Step 5: Main script logic ===
def main():
    """
    module_codes = [
        'GG4258', 'GG3281', 'GG1002', 'GG2014', 'GG4248', 'GG4247', 'SS5103',
        'GG4254', 'GG4257', 'GG3205', 'GG3213', 'GG3214', 'GG5005', 'GG4399',
        'SD4126', 'SD4129', 'SD4133', 'SD1004', 'SD4225', 'SD2006', 'SD2100',
        'SD4110', 'SD3102', 'SD3101', 'SD4120', 'SD4125', 'SD4297', 'SD5801',
        'SD5802', 'SD5805', 'SD5806', 'SD5807', 'SD5810', 'SD5820', 'SD5821',
        'SD5811', 'SD5813', 'SD5812'
    ]"""
    module_codes = [
        'GG4258', 'GG3281'
    ]
    

    print("=== St Andrews Final Grades Extractor ===\n")
    driver = setup_driver()

    try:
        manual_authentication(driver)
        all_data = {}
        for code in module_codes:
            try:
                records = extract_grades_from_module(driver, code)
                if records:
                    all_data[code] = records
            except Exception as e:
                print(f"  ✗ Failed to extract {code}: {e}")
            time.sleep(1.5)

        if all_data:
            save_to_excel(all_data)
        else:
            print("No data was extracted.")

    finally:
        input("\nPress Enter to close the browser...")
        driver.quit()


if __name__ == "__main__":
    main()
