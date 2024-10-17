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

# Laden der Umgebungsvariablen
load_dotenv()

HEADER = ["Modul", "1,0", "1,3", "1,7", "2,0", "2,3",
          "2,7", "3,0", "3,3", "3,7", "4,0", "5,0", "Note", "CP"]
noten_gesamt = []
eigene_noten_cp_gesamt = []

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

def navigate_to_modulergebnisse(driver):
    wait = WebDriverWait(driver, 10)
    wait.until(EC.element_to_be_clickable((By.ID, "link000280"))).click()
    wait.until(EC.element_to_be_clickable((By.ID, "link000323"))).click()
    wait.until(EC.element_to_be_clickable((By.ID, "link000324"))).click()

def scrape_semester_data(driver):
    global noten_gesamt, eigene_noten_cp_gesamt
    wait = WebDriverWait(driver, 10)
    semester_select = Select(wait.until(EC.presence_of_element_located((By.ID, "semester"))))
    semesters = [option.text for option in semester_select.options]
    
    for index in range(len(semesters)):
        semester_select = Select(driver.find_element(By.ID, "semester"))
        semester_select.select_by_index(index)
        
        wait.until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr")))
        alle_zeilen = driver.find_elements(By.XPATH, "//table/tbody/tr")
        
        for zeile in alle_zeilen:
            source_zeile = zeile.get_attribute("innerHTML")
            if "Notenspiegel" not in source_zeile:
                continue
            soup = BeautifulSoup(source_zeile, "html.parser")
            note_cp = soup.find_all("td", {"class": "tbdata_numeric"})
            
            for i in range(0, len(note_cp), 2):
                note_text = note_cp[i].text.strip()
                cp_text = note_cp[i+1].text.strip()
                
                note = 1.0 if note_text == "b" else float(re.sub('\s+', ' ', note_text).replace(',', '.'))
                cp = float(re.sub('\s+', ' ', cp_text).replace(',', '.'))
                
                eigene_noten_cp_gesamt.append([note, cp])
        
        open_notenspiegel(driver)
    
def open_notenspiegel(driver):
    wait = WebDriverWait(driver, 10)
    original_handle = driver.current_window_handle
    
    links = driver.find_elements(By.XPATH, "//*[@title='Notenspiegel']")
    for link in links:
        link.click()
        wait.until(EC.number_of_windows_to_be(2))
        
        new_handle = [handle for handle in driver.window_handles if handle != original_handle][0]
        driver.switch_to.window(new_handle)
        
        soup = BeautifulSoup(driver.page_source, "html.parser")
        modul_name = re.sub('\s+', ' ', soup.find("h2").text.strip())
        
        noten = soup.find_all("td", {"class": "tbdata"})
        if noten:
            noten.pop(0)  # Entferne erstes Element "Anzahl"
        
        noten_liste = [modul_name]
        for note in noten:
            noten_liste.append(0 if note.text.strip() == "---" else int(note.text.strip()))
        
        noten_gesamt.append(noten_liste)
        
        driver.close()
        driver.switch_to.window(original_handle)

def merge_data():
    for i in range(len(noten_gesamt)):
        if i < len(eigene_noten_cp_gesamt):
            noten_gesamt[i].extend(eigene_noten_cp_gesamt[i])

def filter_final_list():
    global noten_gesamt
    noten_gesamt = [zeile for zeile in noten_gesamt if len(zeile) > 6]

def export_to_csv():
    with open("noten.csv", "w", encoding="UTF8", newline='') as f:
        writer = csv.writer(f)
        writer.writerow(HEADER)
        for zeile in noten_gesamt:
            if len(zeile) >= 5:
                writer.writerow(zeile)

def main():
    username = os.getenv('username')
    password = os.getenv('password')
    
    driver = initialize_driver()
    try:
        login(driver, username, password)
        navigate_to_modulergebnisse(driver)
        scrape_semester_data(driver)
        merge_data()
        filter_final_list()
        export_to_csv()
    finally:
        driver.quit()

if __name__ == "__main__":
    main()