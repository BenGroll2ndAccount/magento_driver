import os
import re
from unicodedata import name
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
from pynput import keyboard
import shutil
import subprocess 

global has_pressed_space
has_pressed_space = False

global current_page
current_page = "General"


def onpressspace(key):
    print("onpressspace triggered")
    global has_pressed_space
    if key == keyboard.Key.space:
        has_pressed_space = True
    print("finished")


def startup():
    global has_pressed_space
    global current_page
    window = gui.Window(title="Magento AutoDriver", layout =[
        [gui.Text("Name?"), gui.Input(key="-Name-", default_text="Groll, Ben")],
        [gui.Text("Username?"), gui.Input(key="-Uname-")],
        [gui.Text("Password?"), gui.Input(key="-Pword-")],
        [gui.Text("Selection?"), gui.Combo(values=("by_mpn_folder", "by_mpn_folders(diff)"), default_value = "by_mpn_folder", key = "-Selection-")],
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
            window.close()
            layout = [
                [gui.Text("Folder Path? (Folder containing your MPNs.)"), gui.Input(key="-Folder-", default_text="/Users/bengrollpmt/Desktop/Images 22.2. neu")] if selection == "by_mpn_folder" or selection == "by_mpn_folders(diff)" else [],
                [gui.Text("Folder Path? (Second folder containing left over MPNs."), gui.Input(key="-SFolder-")] if selection == "by_mpn_folders(diff)" else [],
                [gui.Text("MPN Seperator? (, . ' ' etc) "), gui.Combo(values = (" ", ",", "."),key="-Seperator-", default_value=" ")],
                [gui.Text("-- Choose Actions --")],
                [gui.Text("Update last_modified?"), gui.Checkbox(text="", key="edit:last_modified", default = True)],
                [gui.Text("Edit Internal Comments?"), gui.Checkbox(text ="",key="edit:internal_comments"), gui.Input(key="value:internal_comments")],
                [gui.Text("            Mode: "), gui.Combo(values=("Add", "Replace"), key="mode:internal_comments")],
                [gui.Text("Edit Selection for Backend?"), gui.Checkbox(text="", key="edit:selection_for_backend"), gui.Input(key="value:selection_for_backend")],
                [gui.Text("            Mode: "), gui.Combo(values=("Add", "Replace"), key="mode:selection_for_backend")],
                [gui.Text("Edit Status?"), gui.Checkbox(text="", key = "edit:status", default=True), gui.Combo(values = ("Disabled", "Enabled"), key="value:status", default_value = "Enabled")], 
                [gui.Text("Edit Product Advertising Status?"), gui.Checkbox(text="", key="edit:product_advertising_status"), gui.Combo(values = ("Item is in stock", "Not in stock: Full advertisement (Shopping cart ON, discounts off", "Not in stock: No advertisement (but shown in catalog & search)", "Not in stock: No advertisement (archive only)", "-- not in use -- Not in stock: Full advertisement (discounts active)", "-- not in use -- Not in stock: 'Open box' (full advertisement, discounts active)"), key="value:product_advertising_status")],
                [gui.Button("Save and Continue", key="-Confirm-Actions")]
            ]
            window = gui.Window(title="Magento AutoDriver - Actions Page 1", layout = layout)
        if event == "-Confirm-Actions":
            path_to_folder = values["-Folder-"]
            path_to_second_folder = values["-SFolder-"] if selection == "by_mpn_folders(diff)" else None
            seperator = values["-Seperator-"]
            actions = values            
            window.close()
            layout = [
                [gui.Text("Add Filters to your selection?")],
                [gui.Text("Selection for Backend?"), gui.Checkbox(text="", key="filter:selection_for_backend"), gui.Input(key="value:selection_for_backend")],
                [gui.Text(" ")],
                [gui.Text(" ")],
                [gui.Text(" ")],
                [gui.Text(" ")],
                [gui.Text("Misc.")],
                [gui.Text("What page to stop waiting for input?"), gui.Combo(values = ("General", "Prices & shipping", "Images", "Internal / database work", "Meta Information", "Inventory", "Websites", "Categories", "Accessories", "Up-sells", "Cross-sells", "Product Alerts", "Product Reviews", "Product Tags", "Customers Tagged Product", "Custom Options", "Videos", "Product Attachments"), key = "-Page-To-Wait-For-User-Input-", default_value="Images")],
                [gui.Text("Wait for user input after actions?"), gui.Checkbox(text="", key ="-Wait-For-User-Input-", default=True)], 
                [gui.Button("Start AutoDriver", key="start")]
            
            ]
            window = gui.Window(title="Magento AutoDriver - Filters Page 1", layout = layout)

        if event == "start":
            total_start_time = time.time()
            if selection == "by_mpn_folder" or "by_mpn_folders(diff)":
                mpn_to_folder_map = {}
                filelist = os.listdir(path_to_folder)
                mpnlist = []
                for entry in filelist:
                    mpn = entry.split(seperator)[0]
                    mpnlist.append(mpn)
                    mpn_to_folder_map[str(mpn)] = str(entry)

                if "" in mpnlist: 
                    mpnlist.remove("")
                if selection == "by_mpn_folders(diff)":
                    second_filelist = os.listdir(path_to_second_folder)
                    mpnlist_two = []
                    for entry in second_filelist:
                        mpn = entry.split(seperator)[0]
                        mpnlist_two.append(mpn)
                    if "" in mpnlist:
                        mpnlist.remove("")
                    ##Select bigger(unedited) mpnlist
                    if len(mpnlist) > len(mpnlist_two):
                        remaining_list = []
                        for entry in mpnlist:
                            if entry not in mpnlist_two:
                                remaining_list.append(entry)
                    elif len(mpnlist) < len(mpnlist_two):
                        remaining_list = []
                        for entry in mpnlist_two:
                            if entry not in mpnlist:
                                remaining_list.append(entry)
                    else:
                        print(">>> Folders contain same amount of products")
                        quit()
                elif selection == "by_mpn_folder":
                    remaining_list = mpnlist
                elementcount = len(remaining_list)
                print(remaining_list)
                ##### Generate Progress Window
                window.close()
                window = gui.Window(title="Magento AutoDriver - Running Process", layout=[
                [gui.Text("Total Elements to do: " + str(elementcount))],
                [gui.Text("   ")],
                [gui.Text("Progress:")],
                [gui.ProgressBar(elementcount, orientation="h", key="-Progress-Bar-", size=(20, 20))],        
                [gui.Text("Finished: 0/" + str(elementcount), key="-Progress-Text-")],
                [gui.Text("   ")],
                [gui.Text("   ")],
                [gui.Text("MPN:", key="-current-mpn-")],
                [gui.Text("Estimated Time: ---- ", key="-Time-Text-")]
                ])
                driver = load()
                window.read(timeout=100)
                wait = WebDriverWait(driver,1)
                login(username, password, driver)
                elements_finished = 0
                enter_filters(values = values, drive = driver)
                estimated_finalization_time = 0
                for mpn_entry in remaining_list:
                    if elements_finished == 0:
                        starttime = time.time()
                    subprocess.run("pbcopy", universal_newlines=True, input = str(mpn_entry))
                    mpnfield = driver.find_element(By.ID, "productGrid_product_filter_mpn")
                    mpnfield.clear()
                    mpnfield.send_keys
                    mpnfield.send_keys(mpn_entry)
                    mpnfield.send_keys(Keys.RETURN)
                    sleep(3)
                    try:
                        try:
                            wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Action")))
                        except:
                            print("Presence not ass")
                        try:
                            wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Action")))
                        except:
                            print("Not Lickable")
                        try:
                            actionfield = driver.find_element(By.PARTIAL_LINK_TEXT, "Action")
                            actionfield.click()
                        except:
                            print("Misclick")
                        wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(@title, 'Save and Continue Edit')]"))) 
                        ####### What to actually do to the entry     
                        if actions["edit:last_modified"]:
                            set_last_modified(driver, name)
                            print("Set last Modified")
                        if actions["edit:internal_comments"]:
                            set_internal_comments(drive = driver, values = actions, name=name)
                            print("Set internal Comments")
                        if actions["edit:selection_for_backend"]:
                            set_selection_for_backend(drive = driver, values = actions, name=name)
                            print("Set selection_for_backend")
                        if actions["edit:status"]:
                            set_status(drive = driver, values = actions)
                            print("Set Status")
                        if actions["edit:product_advertising_status"]:
                            print("Set product advertising status")
                            set_product_advertising_status(drive = driver, values = actions)
                        
                                
                        print("Start Wait for User Input")
                        ####### Wait for User Input:
                        print(actions)
                        print(values)
                        if values["-Wait-For-User-Input-"]:
                            if values["-Page-To-Wait-For-User-Input-"] != current_page:
                                switch_to_page(driver, values["-Page-To-Wait-For-User-Input-"])
                            gui.popup()

                        ####### Save and reset
                        try:
                            save(driver)
                        except:
                            print("Save Error " + mpn_entry)
                            #/Users/bengrollpmt/Desktop/Images 22.2. orig
                            sleep(2)
                        
                        shutil.rmtree(mpn_to_folder_map[mpn_entry])
                        print("Deleted " + str(mpn_to_folder_map[mpn_entry]))
                    
                    except:
                        print(">>>Object skipped: " + mpn_entry)
                        
                    
                    # Update Window
                    elements_finished += 1
                    if elements_finished == 1:
                        endtime = time.time()
                        total_time = endtime - starttime
                        window["-Time-Text-"].update("Estimated time until finished: " + str(int(total_time * elementcount / 60)) + "min")
                        estimated_finalization_time = time.time() + int(total_time * elementcount)
                    else:
                        window["-Time-Text-"].update("Estimated time until finished: " + str(int((estimated_finalization_time - time.time())) / 60) + "min")
                    window["-Progress-Text-"].update("Finished: " + str(elements_finished) + "/" + str(elementcount))
                    window["-Progress-Bar-"].update(current_count = elements_finished)
                    window.read(timeout= 1)
                finish(total_start_time=total_start_time, elementcount=elementcount)
                
            else:
                driver.quit()
                window.close()
                raise Exception

            
def finish(total_start_time, elementcount):
    total_end_time = time.time()
    spent_time = (total_end_time - total_start_time) / 60
    gui.popup("Finished " + str(elementcount) + " elements in " + str(spent_time) + " minutes.", keep_on_top=True)
    quit()  

def enter_filters(values, drive):
    if values["filter:selection_for_backend"]:
        selection_for_backend_field = drive.find_element(By.ID, "productGrid_product_filter_internal_selection_for_backend")
        selection_for_backend_field.send_keys(values["value:selection_for_backend"])
        sleep(2)

def switch_to_page(drive : webdriver.Chrome, page):
    print("Starting Switch to page")
    #//button[contains(@title, 'Save and Continue Edit')]
    page_tab = drive.find_element(By.XPATH, "//a[contains(@title, '" + page + "')]")
    #page_tab.find_element(By.XPATH, "./..")
    page_tab.click()

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

def set_status(drive, values):
    selection = Select(drive.find_element(By.ID, "status"))
    selection.select_by_visible_text(values["value:status"])
    sleep(1)

def set_internal_comments(drive, values, name):
    internal_comments_field = drive.find_element(By.ID, "internal_comments")
    if values["mode:internal_comments"] == "Replace":
        internal_comments_field.clear()
    internal_comments_field.send_keys(values["value:internal_comments"] + build_nametag(build_short(name)))
    internal_comments_field.send_keys(Keys.RETURN)

def set_selection_for_backend(drive, values, name):
    selection_for_backend_field = drive.find_element(By.ID, "internal_selection_for_backend")
    if values["mode:selection_for_backend"] == "Replace":
        selection_for_backend_field.clear()
    selection_for_backend_field.send_keys(values["value:selection_for_backend"] + build_nametag(build_short(name)))
    selection_for_backend_field.send_keys(Keys.RETURN)

def set_product_advertising_status(drive, values):
    selection = Select(drive.find_element(By.ID, "artikel_status"))
    selection.select_by_visible_text(values["value:product_advertising_status"])

def save(drive):
    savebutton = drive.find_element(By.XPATH, "//button[contains(@title, 'Save')]")
    savebutton.click()
    sleep(5)

def build_short(name):
    name = name.split(", ")
    name = str(name[1])[0].lower() + str(name[0])[0].lower()
    return name

def build_nametag(kürzel):
    today = str(date.today()).split("-")
    today[0], today[2] = today[2], today[0]
    today = str(today[0] + today[1] + today[2][2:]) + kürzel
    return today

startup()
