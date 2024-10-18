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
columns = ["Modul", "1,0", "1,3", "1,7", "2,0", "2,3",
           "2,7", "3,0", "3,3", "3,7", "4,0", "5,0", "Note", "CP"]
pd_noten_gesamt = pd.DataFrame(columns=columns)

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

def get_note_cp(source_zeile):
    soup = BeautifulSoup(source_zeile, "html.parser")
    modul_name = soup.find_all("td", {"class": "tbdata"})[1].text.strip()

    note_cp = soup.find_all("td", {"class": "tbdata_numeric"})

    note_text = note_cp[0].text.strip()
    cp_text = note_cp[1].text.strip()
    
    eigene_note = 1.0 if note_text == "b" else float(re.sub('\s+', ' ', note_text).replace(',', '.'))
    cp = float(re.sub('\s+', ' ', cp_text).replace(',', '.'))

    return eigene_note, cp, modul_name

def scrape_notenspiegel(driver, zeile):
    # Finde den Link mit dem Titel "Notenspiegel" innerhalb der Zeile
    notenspiegel_link = zeile.find_element(By.XPATH, ".//a[@title='Notenspiegel']")
    notenspiegel_link.click()
    
    # Optional: Warte auf das Popup oder handle das neue Fenster
    wait = WebDriverWait(driver, 10)
    wait.until(EC.number_of_windows_to_be(2))
    
    # Wechsle zum neuen Fenster
    windows = driver.window_handles
    driver.switch_to.window(windows[-1])
    
    # Führe hier deine Scraping-Logik für das Notenspiegel-Popup aus
    soup = BeautifulSoup(driver.page_source, "html.parser")
    modul_name = re.sub('\s+', ' ', soup.find("h2").text.strip())
    noten_raw = soup.find_all("td", {"class": "tbdata"})
    noten_raw.pop(0)  # Entferne erstes Element "Anzahl"
    noten = []
    for note in noten_raw:
        if note.text.strip() == "---":
            noten.append(0)
        else:
            noten.append(int(note.text.strip()))
    
    # Schließe das Popup und wechsle zurück
    driver.close()
    driver.switch_to.window(windows[0])

    return modul_name, noten

def scrape_semester_data(driver):
    global noten_gesamt, eigene_noten_cp_gesamt
    wait = WebDriverWait(driver, 10)
    semester_select = Select(wait.until(EC.presence_of_element_located((By.ID, "semester"))))
    semesters = [option.text for option in semester_select.options]
    
    # Loop über alle Semester
    for index in range(len(semesters)):
        semester_select = Select(driver.find_element(By.ID, "semester"))
        semester_select.select_by_index(index)
        
        wait.until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr")))
        alle_zeilen = driver.find_elements(By.XPATH, "//table/tbody/tr")
        alle_zeilen.pop(-1)
        
        # Loop über alle Module in einem Semester
        for zeile in alle_zeilen:
            source_zeile = zeile.get_attribute("innerHTML")
            # Manche Module geben keinen Notenspiegel. Für diese werden 0 als Noten eingetragen.
            if "Notenspiegel" in source_zeile:
                modul_name_alt, noten = scrape_notenspiegel(driver, zeile)
            else:
                noten = [0] * 11
            
            eigene_note, cp, modul_name = get_note_cp(source_zeile)

            print(f"Modul: {modul_name}, Noten: {noten}, Note: {eigene_note}, CP: {cp}")

            # Manche Module haben nur "Bestanden/Nicht bestanden" als Note. Diese werden ignoriert.
            if len(noten) == 11:
                pd_noten_gesamt.loc[len(pd_noten_gesamt)] = [modul_name] + noten + [eigene_note, cp]

def main():
    load_dotenv()
    username = os.getenv('username')
    password = os.getenv('password')
    
    driver = initialize_driver()
    login(driver, username, password)
    navigate_to_modulergebnisse(driver)
    scrape_semester_data(driver)
    driver.quit()

    pd_noten_gesamt.to_csv("noten.csv", index=False)

if __name__ == "__main__":
    main()