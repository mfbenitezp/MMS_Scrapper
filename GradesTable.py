import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager


def setup_driver():
    edge_options = Options()
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument("--disable-dev-shm-usage")
    edge_options.add_argument("--window-size=1920,1080")
    service = Service(EdgeChromiumDriverManager().install())
    return webdriver.Edge(service=service, options=edge_options)


def manual_login(driver, test_url):
    print("=== Manual Authentication ===")
    print("1. A browser will open.")
    print("2. Login to MMS and navigate to the test module page.")
    print("3. Once you see the full table, come back and press Enter here.")
    driver.get(test_url)
    input("Press Enter when you're on the module page and see the grades table...")


def extract_table_html(driver):
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "gradesTable"))
    )
    table_element = driver.find_element(By.ID, "gradesTable")
    return table_element.get_attribute("outerHTML")


def parse_html_table_to_dataframe(table_html):
    soup = BeautifulSoup(table_html, "html.parser")
    table = soup.find("table")
    df = pd.read_html(str(table))[0]
    return df


def main():
    module_url = "https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/GG1002/Final+grade/"
    driver = setup_driver()

    try:
        manual_login(driver, module_url)
        table_html = extract_table_html(driver)
        df = parse_html_table_to_dataframe(table_html)

        print("\n=== DataFrame Preview ===")
        print(df.head())
        print("\n=== Column Names ===")
        print(df.columns.tolist())

    finally:
        input("\nPress Enter to close the browser...")
        driver.quit()


if __name__ == "__main__":
    main()
