import sys
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
from time import sleep
from random import betavariate, randint
import re
import csv

# Für CSV Export
header = ["Modul", "1,0", "1,3", "1,7", "2,0", "2,3",
          "2,7", "3,0", "3,3", "3,7", "4,0", "5,0", "Note", "CP"]
noten_gesamt = []
eigeneNotenUndCpGesamt = []

# webdriver intitialisieren
driver = webdriver.Chrome(
    executable_path="/Users/sebastiangeis/Documents/Visual Studio Code/chromedriver")

# Tucan aufrufen
url = "https://www.tucan.tu-darmstadt.de/"
driver.get(url)
sleep(2)

# Login, tu id und passwort müssen eingetragen werden
username = "tu-id"
password = "pwd"

driver.find_element("name", "usrname").send_keys(username)
driver.find_element("name", "pass").send_keys(password)
driver.find_element(By.ID, "logIn_btn").click()

# Zu Modulergebnissen navigieren
sleep(2)
driver.find_element(By.ID, "link000280").click()
sleep(2)
driver.find_element(By.ID, "link000323").click()
sleep(2)
driver.find_element(By.ID, "link000324").click()

# Durch Semester iterieren -------------------------------------
sleep(2)
semesterAuswahl = Select(driver.find_element(By.ID, "semester"))

anzahl_options = 0
list_options = []
for option in semesterAuswahl.options:
    list_options.append(option.text)
    anzahl_options += 1

# semesterAuswahl muss jedes mal neu erstellt werden, da es sich durch "select" aktualisiert
semesterIndex = 0
while semesterIndex < anzahl_options:
    print("SemesterIndex" + str(semesterIndex))
    semesterAuswahlNeu = Select(driver.find_element(By.ID, "semester"))
    semesterAuswahlNeu.select_by_index(semesterIndex)

    # CP und eigene Note auslesen
    alleZeilen = driver.find_elements(By.XPATH, "//table/tbody/tr")
    for zeile in alleZeilen:
        sourceZeile = zeile.get_attribute("innerHTML")
        if "Notenspiegel" not in sourceZeile:
            continue
        else:
            soup = BeautifulSoup(sourceZeile, "html.parser")

            noteUndCP = soup.find_all("td", {"class": "tbdata_numeric"})

            noteUndCPIndex = 0
            while noteUndCPIndex < len(noteUndCP):
                eigeneNoteUndCpListe = []

                if (noteUndCP[noteUndCPIndex].text).strip() == "b":
                    value1 = 1
                else:
                    value1 = re.sub(
                        '\s+', ' ', (noteUndCP[noteUndCPIndex].text).strip())
                    value1 = value1.replace(',', '.')
                    value1 = float(value1)

                value2 = re.sub(
                    '\s+', ' ', (noteUndCP[noteUndCPIndex+1].text).strip())
                value2 = value2.replace(',', '.')
                value2 = float(value2)

                eigeneNoteUndCpListe.append(value1)
                eigeneNoteUndCpListe.append(value2)
                eigeneNotenUndCpGesamt.append(eigeneNoteUndCpListe)
                noteUndCPIndex += 2

            for zeile in noten_gesamt:
                print(zeile)

    # Prüfungsdurchschnitt öffnen -------------------------------------
    sleep(2)
    originalHandle = driver.current_window_handle
    for link in driver.find_elements(By.XPATH, "//*[@title='Notenspiegel']"):

        # neues Fenster öffnen
        link.click()

        sleep(1)
        windowHandles = driver.window_handles
        for handle in windowHandles:
            if handle != originalHandle:
                driver.switch_to.window(handle)
                break

        # neues Fenster auslesen -------------------------------------
        sleep(1)
        source = driver.page_source

        soup = BeautifulSoup(source, "html.parser")

        # Modulname in Liste einfügen
        modulName = soup.find("h2").text
        modulName = modulName.strip()
        modulName = re.sub('\s+', ' ', modulName)

        noten = soup.find_all("td", {"class": "tbdata"})
        noten.pop(0)  # Erstes Element "Anzahl" entfernen

        notenListe = []
        notenListe.append(modulName)

        # Noten zur Liste hinzufügen
        for i in noten:
            if i.text == "---":
                notenListe.append(0)
            else:
                notenListe.append(int((i.text).strip()))

        for i in notenListe:
            print(i)

        noten_gesamt.append(notenListe)

        # Fenster schließen
        driver.close()
        driver.switch_to.window(originalHandle)

    semesterIndex += 1
    continue

# Listen zusammenführen
i = 0
while i < len(noten_gesamt):
    eigeneNote = eigeneNotenUndCpGesamt[i][0]
    cp = eigeneNotenUndCpGesamt[i][1]

    noten_gesamt[i].append(eigeneNote)
    noten_gesamt[i].append(cp)

    i += 1

# Um Einträge mit "bestanden/nicht bestanden" rauszuwerfen
listeFinal = []
for zeile in noten_gesamt:
    if len(zeile) > 6:
        listeFinal.append(zeile)


# CSV export
with open("noten.csv", "w", encoding="UTF8") as f:
    writer = csv.writer(f)

    # write header
    writer.writerow(header)

    for zeile in listeFinal:
        if (len(zeile) < 5):
            continue
        writer.writerow(zeile)

driver.close()
