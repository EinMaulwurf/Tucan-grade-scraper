import csv
import os
import re
from time import sleep
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd

# Set up pandas dataframe
columns = ["Module", "1.0", "1.3", "1.7", "2.0", "2.3",
           "2.7", "3.0", "3.3", "3.7", "4.0", "5.0", "Grade", "CP"]
df_grades_total = pd.DataFrame(columns=columns)

def initialize_driver():
    options = Options()
    options.headless = True
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def login(driver, username, password):
    url = "https://www.tucan.tu-darmstadt.de/"
    driver.get(url)
    wait = WebDriverWait(driver, 10)
    
    wait.until(EC.presence_of_element_located((By.NAME, "usrname"))).send_keys(username)
    driver.find_element(By.NAME, "pass").send_keys(password)
    driver.find_element(By.ID, "logIn_btn").click()

def navigate_to_module_results(driver):
    wait = WebDriverWait(driver, 10)
    wait.until(EC.element_to_be_clickable((By.ID, "link000280"))).click()
    wait.until(EC.element_to_be_clickable((By.ID, "link000323"))).click()
    wait.until(EC.element_to_be_clickable((By.ID, "link000324"))).click()

def get_grade_cp(row_source):
    soup = BeautifulSoup(row_source, "html.parser")
    module_name = soup.find_all("td", {"class": "tbdata"})[1].text.strip()

    grade_cp = soup.find_all("td", {"class": "tbdata_numeric"})

    grade_text = grade_cp[0].text.strip()
    cp_text = grade_cp[1].text.strip()
    
    own_grade = 1.0 if grade_text == "b" else float(re.sub('\s+', ' ', grade_text).replace(',', '.'))
    cp = float(re.sub('\s+', ' ', cp_text).replace(',', '.'))

    return own_grade, cp, module_name

def scrape_grade_sheet(driver, row):
    # Find the link with the title "Notenspiegel" within the row
    grade_sheet_link = row.find_element(By.XPATH, ".//a[@title='Notenspiegel']")
    grade_sheet_link.click()
    
    # Optional: Wait for the popup or handle the new window
    wait = WebDriverWait(driver, 10)
    wait.until(EC.number_of_windows_to_be(2))
    
    # Switch to the new window
    windows = driver.window_handles
    driver.switch_to.window(windows[-1])
    
    # Perform scraping logic for the Notenspiegel popup
    soup = BeautifulSoup(driver.page_source, "html.parser")
    module_name = re.sub('\s+', ' ', soup.find("h2").text.strip())
    grades_raw = soup.find_all("td", {"class": "tbdata"})
    grades_raw.pop(0)  # Remove first element "Count"
    grades = []
    for grade in grades_raw:
        if grade.text.strip() == "---":
            grades.append(0)
        else:
            grades.append(int(grade.text.strip()))
    
    # Close the popup and switch back
    driver.close()
    driver.switch_to.window(windows[0])

    return module_name, grades

def scrape_semester_data(driver):
    global df_grades_total
    wait = WebDriverWait(driver, 10)
    semester_select = Select(wait.until(EC.presence_of_element_located((By.ID, "semester"))))
    semesters = [option.text for option in semester_select.options]
    
    # Loop through all semesters
    for index in range(len(semesters)):
        semester_select = Select(driver.find_element(By.ID, "semester"))
        semester_select.select_by_index(index)
        
        wait.until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr")))
        all_rows = driver.find_elements(By.XPATH, "//table/tbody/tr")
        all_rows.pop(-1)
        
        # Loop through all modules in a semester
        for row in all_rows:
            row_source = row.get_attribute("innerHTML")
            # Some modules do not provide a Notenspiegel. For these, 0 is entered as grades.
            if "Notenspiegel" in row_source:
                module_name_old, grades = scrape_grade_sheet(driver, row)
            else:
                grades = [0] * 11
            
            own_grade, cp, module_name = get_grade_cp(row_source)

            print(f"Module: {module_name}, Grades: {grades}, Grade: {own_grade}, CP: {cp}")

            # Some modules only have "Passed/Failed" as grade. These are ignored.
            if len(grades) == 11:
                df_grades_total.loc[len(df_grades_total)] = [module_name] + grades + [own_grade, cp]

def main():
    load_dotenv()
    username = os.getenv('username')
    password = os.getenv('password')
    
    driver = initialize_driver()
    login(driver, username, password)
    navigate_to_module_results(driver)
    scrape_semester_data(driver)
    driver.quit()

    df_grades_total.to_csv("grades.csv", index=False)

if __name__ == "__main__":
    main()