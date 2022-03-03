import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from time import sleep
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from datetime import date
import PySimpleGUI as gui

def startup():
    window = gui.Window(title="Magento AutoDriver", layout =[
        [gui.Text("Name?"), gui.Input(key="-Name-")],
        [gui.Text("Username?"), gui.Input(key="-Uname-")],
        [gui.Text("Password?"), gui.Input(key="-Pword-")],
        [gui.Text("Selection?"), gui.Combo(values=("by_mpn_folder", "Not Implemented."), default_value = "by_mpn_folder", key = "-Selection-")],
        [gui.Button("Confirm", key="-Confirm-First-")]
        ], margins=(10, 50))

    while True:
        event, values = window.read()
        if event == "Exit" or event == gui.WIN_CLOSED:
            break
        if event == "-Confirm-First-":
            name = values["-Name-"]
            username = values["-Uname-"]
            password = values["-Pword-"]
            selection = values["-Selection-"]
            print(name)
            print(selection)
            window.close()
            layout = [
                [gui.Text("Folder Path? (Folder containing your MPNs.)"), gui.Input(key="-Folder-")] if selection == "by_mpn_folder" else [],
                [gui.Text("-- Choose Actions --")],
                [gui.Text("Update last_modified?"), gui.Checkbox(text="", key="edit:last_modified")],
                [gui.Text("Edit Internal Comments?"), gui.Checkbox(text ="",key="edit:internal_comments"), gui.Input(key="value:internal_comments")],
                [gui.Text("            Mode: "), gui.Combo(values=("Add", "Replace"), key="mode:internal_comments")],
                [gui.Text("Edit Selection for Backend?"), gui.Checkbox(text="", key="edit:selection_for_backend"), gui.Input(key="value:selection_for_backend")],
                [gui.Text("            Mode: "), gui.Combo(values=("Add", "Replace"), key="mode:selection_for_backend")],
                [gui.Button("Save and Continue", key="-Confirm-Page1")]
            ]
            window = gui.Window(title="Magento AutoDriver - Actions Page 1", layout = layout)
        if event == "-Confirm-Page1":
            path_to_folder = values["-Folder-"]
            actions = values
            total_start_time = time.time()
            if selection == "by_mpn_folder":
                filelist = os.listdir(path_to_folder)
                mpnlist = []
                for entry in filelist:
                    mpn = entry.split(".")[0]
                    mpnlist.append(mpn)
                mpnlist.remove("")
                elementcount = len(mpnlist)
                print(mpnlist)
                ##### Generate Progress Window
                window.close()
                window = gui.Window(title="Magento AutoDriver - Running Process", layout=[
                [gui.Text("Total Elements to do: " + str(elementcount))],
                [gui.Text("   ")],
                [gui.Text("Progress:")],
                [gui.ProgressBar(elementcount, orientation="h", key="-Progress-Bar-")],        
                [gui.Text("Finished: 0/" + str(elementcount), key="-Progress-Text-")],
                [gui.Text("   ")],
                [gui.Text("   ")],
                [gui.Text("   ")],
                [gui.Text("Estimated Time: ---- ", key="-Time-Text-")]
                ])
                driver = load()
                window.read(timeout=100)
                wait = WebDriverWait(driver,10)
                login(username, password, driver)
                elements_finished = 0
                for mpn_entry in mpnlist:
                    if elements_finished == 0:
                        starttime = time.time()
                    mpnfield = driver.find_element(By.ID, "productGrid_product_filter_mpn")
                    mpnfield.clear()
                    mpnfield.send_keys
                    mpnfield.send_keys(mpn_entry)
                    mpnfield.send_keys(Keys.RETURN)
                    sleep(2)
                    wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Action")))
                    actionfield = driver.find_element(By.PARTIAL_LINK_TEXT, "Action")
                    actionfield.click()
                    wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(@title, 'Save')]")))         
                    if values["edit:last_modified"] == True:
                        set_last_modified(driver, name)
                    save(driver)
                    elements_finished += 1
                    if elements_finished == 1:
                        endtime = time.time()
                        total_time = endtime - starttime
                        window["-Time-Text-"].update("Estimated time until finished: " + str(int(total_time * elementcount / 60)) + "min")
                    window["-Progress-Text-"].update("Finished: " + str(elements_finished) + "/" + str(elementcount))
                    window["-Progress-Bar-"].update(current_count = elements_finished)
                    window.read(timeout= 1)
            else:
                driver.quit()
                window.close()
                raise Exception

            total_end_time = time.time()
            spent_time = (total_end_time - total_start_time) / 60
            gui.popup("Finished " + str(elementcount) + " elements in " + str(spent_time) + " minutes.", keep_on_top=True)
            quit()
            


def load():
    oslink = os.getcwd() + "/chromedriver"
    link = "https://www.promusictools.com/index.php/b4ckend/catalog_product/"
    driver = webdriver.Chrome(service = Service(oslink))
    driver.get(link)
    return driver

def login(username, password, drive):
    uname_field = drive.find_element(By.ID, "username")
    uname_field.send_keys(username)
    pword_field = drive.find_element(By.NAME, "login[password]")
    pword_field.send_keys(password)
    pword_field.send_keys(Keys.RETURN)

def set_last_modified(drive, name):
    last_modified_by = Select(drive.find_element(By.ID, "internal_lastmodified_by"))
    try:
        last_modified_by.select_by_visible_text(name)
    except:
        print(">>Name does not fit to product_lastmodified_by")
        drive.quit()
        raise Exception
    last_modified_field = drive.find_element(By.ID,"internal_lastmodified_date")
    last_modified_field.clear()
    today = date.today()
    todaysplit = str(today).split("-")
    today = todaysplit[2] + "/" + todaysplit[1] + "/" + todaysplit[0][2:]
    last_modified_field.send_keys(today)
    last_modified_field.send_keys(Keys.RETURN)

def save(drive):

    savebutton = drive.find_element(By.XPATH, "//button[contains(@title, 'Save')]")
    savebutton.click()
    sleep(5)

startup()
