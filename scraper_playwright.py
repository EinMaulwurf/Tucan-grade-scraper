import csv
import os
import re
from time import sleep
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

# Laden der Umgebungsvariablen
load_dotenv()

HEADER = ["Modul", "1,0", "1,3", "1,7", "2,0", "2,3",
          "2,7", "3,0", "3,3", "3,7", "4,0", "5,0", "Note", "CP"]
noten_gesamt = []
eigene_noten_cp_gesamt = []

def initialize_driver():
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    return playwright, browser, context, page

def login(page, username, password):
    url = "https://www.tucan.tu-darmstadt.de/"
    page.goto(url)
    page.fill('input[name="usrname"]', username)
    page.fill('input[name="pass"]', password)
    page.click('#logIn_btn')

def navigate_to_modulergebnisse(page):
    page.click('#link000280')
    page.click('#link000323')
    page.click('#link000324')

def scrape_semester_data(page):
    global noten_gesamt, eigene_noten_cp_gesamt
    semesters = page.query_selector_all('#semester option')
    for index in range(len(semesters)):
        page.select_option('#semester', index)
        page.wait_for_selector("table tbody tr")
        rows = page.query_selector_all("table tbody tr")
        for row in rows:
            source_zeile = row.inner_html()
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
        open_notenspiegel(page)

def open_notenspiegel(page):
    original_handle = page.context.pages[0]
    links = page.query_selector_all("//*[@title='Notenspiegel']")
    for link in links:
        link.click()
        page.wait_for_event('popup')
        new_page = page.context.pages[-1]
        soup = BeautifulSoup(new_page.content(), "html.parser")
        modul_name = re.sub('\s+', ' ', soup.find("h2").text.strip())
        noten = soup.find_all("td", {"class": "tbdata"})
        if noten:
            noten.pop(0)
        noten_liste = [modul_name]
        for note in noten:
            noten_liste.append(0 if note.text.strip() == "---" else int(note.text.strip()))
        noten_gesamt.append(noten_liste)
        new_page.close()

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

    playwright, browser, context, page = initialize_driver()
    try:
        login(page, username, password)
        navigate_to_modulergebnisse(page)
        scrape_semester_data(page)
        merge_data()
        filter_final_list()
        export_to_csv()
    finally:
        browser.close()
        playwright.stop()

if __name__ == "__main__":
    main()