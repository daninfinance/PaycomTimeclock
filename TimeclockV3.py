import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller
import PySimpleGUI as sg
import pendulum
from datetime import datetime
import time
import os
script_dir = os.path.dirname(__file__)


### Gets the current week and formats accordingly. ###
today = pendulum.now()
Monday = today.start_of('week').format('MM/DD/YYYY')
Tuesday = today.start_of('week').add(days=1).format('MM/DD/YYYY')
Wednesday = today.start_of('week').add(days=2).format('MM/DD/YYYY')
Thursday = today.start_of('week').add(days=3).format('MM/DD/YYYY')
Friday = today.start_of('week').add(days=4).format('MM/DD/YYYY')
Saturday = today.start_of('week').add(days=5).format('MM/DD/YYYY')
Sunday = today.start_of('week').add(days=6).format('MM/DD/YYYY')
dates = [Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday]

### Reads the tablev5.csv file and assigns to df ###
df = pd.read_csv("tablev5.csv")
df['DATE'][0] = Monday
df['DATE'][1] = Tuesday
df['DATE'][2] = Wednesday
df['DATE'][3] = Thursday
df['DATE'][4] = Friday
df['DATE'][5] = Saturday
df['DATE'][6] = Sunday

### Formats the IN DAY and OUT DAY columns to get rid of NaN values ###
df.fillna('', inplace=True)
print(df)
header_list = list(df.columns)
#print(header_list)

sg.theme('Topanga')

### Formats the menu bar at the top ###
menu_def = [['File', ['Settings', 'Exit']]]

### Selenium stuff is below ###
options = webdriver.ChromeOptions()

### Uncomment this line to run the chromedriver in a hidden window ###
#options.add_argument('--headless')
options.add_argument('window-size=1920x1080')

layout = [
    [sg.Titlebar('Time Clock', background_color='#1c1c18')],
    [sg.MenubarCustom(menu_def, bar_text_color=sg.theme_text_color(), pad=(5,5), text_color=sg.theme_text_color(), bar_background_color=sg.theme_background_color(), background_color=sg.theme_background_color())],
    [sg.Text('    DATE', font=('Verdana', 13), justification='center'), sg.Text('       DAY', font=('Verdana', 13), justification='center'), sg.Text('        IN DAY', font=('Verdana', 13), justification='center'), sg.Text(' LUNCH', font=('Verdana', 13), justification='center'), sg.Text('OUT DAY', font=('Verdana', 13), justification='left')],
    [sg.Text(Monday), sg.Text('Monday', size=(10,1)), sg.InputText(size=(8,1), default_text=df.loc[0].at['IN DAY'], key='In_Monday'), sg.InputCombo(('Yes', 'No'), default_value=df.loc[0].at['LUNCH'], key='combo0', size=(4,1)), sg.InputText(size=(8,1), default_text=df.loc[0].at['OUT DAY'], key='Out_Monday')],
    [sg.Text(Tuesday), sg.Text('Tuesday', size=(10,1)), sg.InputText(size=(8,1), default_text=df.loc[1].at['IN DAY'], key='In_Tuesday'), sg.InputCombo(('Yes', 'No'), default_value=df.loc[1].at['LUNCH'], key='combo1', size=(4,1)), sg.InputText(size=(8,1), default_text=df.loc[1].at['OUT DAY'], key='Out_Tuesday')],
    [sg.Text(Wednesday), sg.Text('Wednesday', size=(10, 1)), sg.InputText(size=(8, 1), default_text=df.loc[2].at['IN DAY'], key='In_Wednesday'), sg.InputCombo(('Yes', 'No'), default_value=df.loc[2].at['LUNCH'], key='combo2', size=(4,1)), sg.InputText(size=(8,1), default_text=df.loc[2].at['OUT DAY'], key='Out_Wednesday')],
    [sg.Text(Thursday), sg.Text('Thursday', size=(10, 1)), sg.InputText(size=(8, 1), default_text=df.loc[3].at['IN DAY'], key='In_Thursday'), sg.InputCombo(('Yes', 'No'), default_value=df.loc[3].at['LUNCH'], key='combo3', size=(4,1)), sg.InputText(size=(8,1), default_text=df.loc[3].at['OUT DAY'], key='Out_Thursday')],
    [sg.Text(Friday), sg.Text('Friday', size=(10, 1)), sg.InputText(size=(8, 1), default_text=df.loc[4].at['IN DAY'], key='In_Friday'), sg.InputCombo(('Yes', 'No'), default_value=df.loc[4].at['LUNCH'], key='combo4', size=(4,1)), sg.InputText(size=(8,1), default_text=df.loc[4].at['OUT DAY'], key='Out_Friday')],
    [sg.Text(Saturday), sg.Text('Saturday', size=(10, 1)), sg.InputText(size=(8, 1), default_text=df.loc[5].at['IN DAY'], key='In_Saturday'), sg.InputCombo(('Yes', 'No'), default_value=df.loc[5].at['LUNCH'], key='combo5', size=(4,1)), sg.InputText(size=(8,1), default_text=df.loc[5].at['OUT DAY'], key='Out_Saturday')],
    [sg.Text(Sunday), sg.Text('Sunday', size=(10, 1)), sg.InputText(size=(8, 1), default_text=df.loc[6].at['IN DAY'], key='In_Sunday'), sg.InputCombo(('Yes', 'No'), default_value=df.loc[6].at['LUNCH'], key='combo6', size=(4,1)), sg.InputText(size=(8,1), default_text=df.loc[6].at['OUT DAY'], key='Out_Sunday')]
]

layout += [[sg.Button('Settings', tooltip='Change login info', font=('Verdana', 8)), sg.Text(' ' * 40), sg.Button('Run', font=('Verdana', 8)), sg.Button('Reset', tooltip="Clear entries", font=('Verdana', 8)), sg.Button('Save & Close', font=('Verdana', 8))]]
window = sg.Window('Time Clock', layout, font=('Verdana', 11), auto_size_text=True, resizable=False, titlebar_background_color=sg.theme_background_color(), return_keyboard_events=True, enable_close_attempted_event=True).finalize()
keys = ['In_Monday', 'Out_Monday', 'In_Tuesday', 'Out_Tuesday', 'In_Wednesday', 'Out_Wednesday', 'In_Thursday', 'Out_Thursday', 'In_Friday', 'Out_Friday', 'In_Saturday', 'Out_Saturday', 'In_Sunday', 'Out_Sunday']

while True:
    event, values = window.read()
    window.refresh()

    import json

    ### Gets credentials from credentials.json ###
    file_path = os.path.join(script_dir)
    with open(file_path + '\\credentials.json', 'r') as f:
        config = json.load(f)

    ### Settings Menu for changing password. The password expires so I was going to add a check ###
    ### to see if the 'Update Password' screen appeared and then add a pop up to enter the old  ###
    ### password and new password. Never got around to doing that so this just updates the cre- ###
    ### dentials file for me.                                                                   ###
    if event == 'Settings':
        print("Settings button pressed")
        settings_title_layout = [sg.Text(' Settings Menu', justification='center', size=(30,1), auto_size_text=True, font=('Verdana', 10, 'underline'))]
        settings_layout = [
            settings_title_layout,
            [sg.Text('Update Paycom Credentials', justification='center', size=(30,1), auto_size_text=False)],
            [sg.Text('Old password:', justification='left', size=(13,1)), sg.InputText(size=(15,1), key='old_password', enable_events=True, text_color='gray', password_char='*'), sg.Text('✓', key='old_password_check', size=(2,1), text_color=sg.theme_background_color())],
            [sg.Text('New password:', justification='left', size=(13,1)), sg.Input(size=(15,1), key='new_password', enable_events=True, text_color='gray'), sg.Text('✓', key='new_password_check1', size=(2,1), text_color=sg.theme_background_color())]
        ]
        settings_window = sg.Window('Settings', settings_layout, font=('Verdana', 15), element_justification='Center', auto_size_text=False, return_keyboard_events=True)

        while True:
            event2, value2 = settings_window.read()
            settings_window.refresh()
            print(event2)
            if event2 == sg.WIN_CLOSED or None:
                print("Closing Settings Window")
                break

            if value2['old_password'] == config['user']['pass']:
                print("Yes")
                settings_window['old_password_check'].update(text_color='white')

            if value2['old_password'] != config['user']['pass']:
                settings_window['old_password_check'].update(text_color=sg.theme_background_color())

            if event2 == 'Update':
                if value2['password'] == '':
                    while True:
                        sg.Popup('You can\'t leave the password blank', title='', keep_on_top=True)
                        break
                # Imports the credentials.json file located in the script's folder for login
                else:
                    answer = sg.PopupYesNo('Are you sure you want to change your password?')
                    if answer == 'Yes':
                        import json
                        with open ('credentials.json', 'r') as f:
                            config = json.load(f)
                            f.close()
                        config['user']['pass'] = value2['password']
                        with open('credentials.json', 'w') as f:
                            json.dump(config, f)
                            f.close()
                            print(config)

    if event == 'Exit':
        break

    # The following makes the GUI focus the next element once current focus reaches 8 characters
    elem = window.FindElementWithFocus()
    if elem is not None:
        key = elem.Key
        value = values[key]
        if len(values[key]) > 8:
            window.Element(key).Update(values[key][:-1])
            next_elem = keys[keys.index(key) + 1]
            window[next_elem].SetFocus()

    if event == 'Run':
        chromedriver_autoinstaller.install(cwd=True)
        sg.PopupNoButtons("Starting script", auto_close=True, auto_close_duration=3, no_titlebar=True,
                          font=('Verdana', 20), non_blocking=True)
        for key in keys:
            df.loc[0].at['IN DAY'] = values['In_Monday']
            df.loc[0].at['OUT DAY'] = values['Out_Monday']
            df.loc[0].at['LUNCH'] = values['combo0']
            df.loc[1].at['IN DAY'] = values['In_Tuesday']
            df.loc[1].at['OUT DAY'] = values['Out_Tuesday']
            df.loc[1].at['LUNCH'] = values['combo1']
            df.loc[2].at['IN DAY'] = values['In_Wednesday']
            df.loc[2].at['OUT DAY'] = values['Out_Wednesday']
            df.loc[2].at['LUNCH'] = values['combo2']
            df.loc[3].at['IN DAY'] = values['In_Thursday']
            df.loc[3].at['OUT DAY'] = values['Out_Thursday']
            df.loc[3].at['LUNCH'] = values['combo3']
            df.loc[4].at['IN DAY'] = values['In_Friday']
            df.loc[4].at['OUT DAY'] = values['Out_Friday']
            df.loc[4].at['LUNCH'] = values['combo4']
            df.loc[5].at['IN DAY'] = values['In_Saturday']
            df.loc[5].at['OUT DAY'] = values['Out_Saturday']
            df.loc[5].at['LUNCH'] = values['combo5']
            df.loc[6].at['IN DAY'] = values['In_Sunday']
            df.loc[6].at['OUT DAY'] = values['Out_Sunday']
            df.loc[6].at['LUNCH'] = values['combo6']
        df.to_csv('tablev5.csv', index=False)

        print("Run button pressed.")
        driver = webdriver.Chrome(options=options)
        driver.get("https://www.paycomonline.net/v4/ee/web.php/app/login")
        wait = WebDriverWait(driver, 30)

        # Imports the credentials.json file located in the script's folder for login
        import json
        file_path = os.path.join(script_dir)
        with open(file_path + '\\credentials.json', 'r') as f:
            config = json.load(f)

        # Assigns the username text box on the website to the 'username' variable.
        # Elements can be found by inspecting the HTML and selecting that element.
        username = driver.find_element(By.NAME, "username")
        username.send_keys(config['user']['name'])
        userpass = driver.find_element(By.NAME, "userpass")
        userpass.send_keys(config['user']['pass'])
        userpin = driver.find_element(By.NAME, "userpin")
        userpin.send_keys(config['user']['pin'] + Keys.ENTER)

        with open(file_path + '\\security_questions.json', 'r') as f:
            security_questions = json.load(f)

        # Checks for the security questions page.
        time.sleep(10)
        if driver.find_element(By.TAG_NAME, 'h5'):
            # Finds element by xpath then ties the question from aria-label to Security1
            first_security_question = "//input[@name='first_security_question']"
            Security1 = driver.find_element(By.XPATH, first_security_question).get_attribute('aria-label')
            for key, value in security_questions.items():
                # Checks if the value returned to Security1 matches one of the security_question dict keys.
                if (Security1 == key):
                    Security1 = driver.find_element(By.XPATH, first_security_question).send_keys(value)

            second_security_question = "//input[@name='second_security_question']"
            Security2 = driver.find_element(By.XPATH, second_security_question).get_attribute('aria-label')
            for key, value in security_questions.items():
                if (Security2 == key):
                    Security2 = driver.find_element(By.XPATH, second_security_question).send_keys(value + Keys.ENTER)

            WebTimeSheet = wait.until(EC.visibility_of_element_located((By.LINK_TEXT, "Web Time Sheet Read Only")))
            WebTimeSheet.click()

        for index, row in df.iterrows():
            ##### IN DAY #####
            if row['IN DAY'] == '':
                print("Error for " + row['DATE'] + ": No time entered for IN DAY.")
                sg.Popup("Error for " + row['DATE'] + ": No time entered for IN DAY.", title='Error')
                continue
            wait.until(EC.invisibility_of_element_located((By.XPATH, "//*[@id='punch-change-request-modal']")))
            try:
                RequestNewPunch = driver.find_element(By.NAME, "timecard-add-punch-change-request")
                RequestNewPunch.click()
            except:
                wait.until(EC.visibility_of_element_located((By.NAME, "timecard-add-punch-change-request")))
                RequestNewPunch = driver.find_element(By.NAME, "timecard-add-punch-change-request")
                RequestNewPunch.click()
            pcrDate = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@name='pcrDate']")))
            pcrDate.clear()
            pcrDate.send_keys(row['DATE'])
            driver.find_element(By.XPATH, "//select[@name='pcrPunchType']/option[text()='IN DAY']").click()
            driver.find_element(By.XPATH, "//input[@name='pcrTime']").clear()
            driver.find_element(By.XPATH, "//input[@name='pcrTime']").send_keys(row['IN DAY'])
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'parsleyCloseButton parsleyShow')))
            driver.find_element(By.NAME, "pcrAdd").click()

            ##### OUT LUNCH #####
            if row['LUNCH'] == 'Yes':
                wait.until(EC.invisibility_of_element_located((By.XPATH, "//*[@id='punch-change-request-modal']")))
                try:
                    RequestNewPunch = driver.find_element(By.NAME, "timecard-add-punch-change-request")
                    RequestNewPunch.click()
                except:
                    RequestNewPunch = driver.find_element(By.NAME, "timecard-add-punch-change-request")
                    RequestNewPunch.click()
                pcrDate = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@name='pcrDate']")))
                pcrDate.clear()
                pcrDate.send_keys(row['DATE'])
                driver.find_element(By.XPATH, "//select[@name='pcrPunchType']/option[text()='OUT LUNCH']").click()
                driver.find_element(By.XPATH, "//input[@name='pcrTime']").clear()
                driver.find_element(By.XPATH, "//input[@name='pcrTime']").send_keys("12 PM")
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'parsleyCloseButton parsleyShow')))
                driver.find_element(By.NAME, "pcrAdd").click()

            ##### IN LUNCH #####
                wait.until(EC.invisibility_of_element_located((By.XPATH, "//*[@id='punch-change-request-modal']")))
                try:
                    RequestNewPunch = driver.find_element(By.NAME, "timecard-add-punch-change-request")
                    RequestNewPunch.click()
                except:
                    RequestNewPunch = driver.find_element(By.NAME, "timecard-add-punch-change-request")
                    RequestNewPunch.click()
                pcrDate = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@name='pcrDate']")))
                pcrDate.clear()
                pcrDate.send_keys(row['DATE'])
                driver.find_element(By.XPATH, "//select[@name='pcrPunchType']/option[text()='IN LUNCH']").click()
                driver.find_element(By.XPATH, "//input[@name='pcrTime']").clear()
                driver.find_element(By.XPATH, "//input[@name='pcrTime']").send_keys("1 PM")
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'parsleyCloseButton parsleyShow')))
                driver.find_element(By.NAME, "pcrAdd").click()

            ##### OUT DAY #####
            if row['OUT DAY'] == '':
                print("Error for " + row['DATE'] + ": No time entered for OUT DAY.")
                sg.Popup("Error for " + row['DATE'] + ": No time entered for OUT DAY.", title='Error',
                         background_color='#AFAFAF', keep_on_top=True)
                continue
            wait.until(EC.invisibility_of_element_located((By.XPATH, "//*[@id='punch-change-request-modal']")))
            try:
                RequestNewPunch = driver.find_element(By.NAME, "timecard-add-punch-change-request")
                RequestNewPunch.click()
            except:
                RequestNewPunch = driver.find_element(By.NAME, "timecard-add-punch-change-request")
                RequestNewPunch.click()
            pcrDate = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@name='pcrDate']")))
            pcrDate.clear()
            pcrDate.send_keys(row['DATE'])
            driver.find_element(By.XPATH, "//select[@name='pcrPunchType']/option[text()='OUT DAY']").click()
            driver.find_element(By.XPATH, "//input[@name='pcrTime']").clear()
            driver.find_element(By.XPATH, "//input[@name='pcrTime']").send_keys(row['OUT DAY'])
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'parsleyCloseButton parsleyShow')))
            driver.find_element(By.NAME, "pcrAdd").click()
        sg.PopupNoButtons("Done!", auto_close=True, auto_close_duration=5, no_titlebar=True,
                          font=('Verdana', 20), non_blocking=True)
        driver.quit()

    if event == 'Save & Close':
        print("Save button pressed.")
        for key in keys:
            df.loc[0].at['IN DAY'] = values['In_Monday']
            df.loc[0].at['OUT DAY'] = values['Out_Monday']
            df.loc[0].at['LUNCH'] = values['combo0']
            df.loc[1].at['IN DAY'] = values['In_Tuesday']
            df.loc[1].at['OUT DAY'] = values['Out_Tuesday']
            df.loc[1].at['LUNCH'] = values['combo1']
            df.loc[2].at['IN DAY'] = values['In_Wednesday']
            df.loc[2].at['OUT DAY'] = values['Out_Wednesday']
            df.loc[2].at['LUNCH'] = values['combo2']
            df.loc[3].at['IN DAY'] = values['In_Thursday']
            df.loc[3].at['OUT DAY'] = values['Out_Thursday']
            df.loc[3].at['LUNCH'] = values['combo3']
            df.loc[4].at['IN DAY'] = values['In_Friday']
            df.loc[4].at['OUT DAY'] = values['Out_Friday']
            df.loc[4].at['LUNCH'] = values['combo4']
            df.loc[5].at['IN DAY'] = values['In_Saturday']
            df.loc[5].at['OUT DAY'] = values['Out_Saturday']
            df.loc[5].at['LUNCH'] = values['combo5']
            df.loc[6].at['IN DAY'] = values['In_Sunday']
            df.loc[6].at['OUT DAY'] = values['Out_Sunday']
            df.loc[6].at['LUNCH'] = values['combo6']
        df.to_csv('tablev5.csv', index=False)
        break

    if event == 'Reset':
        for key in keys:
            window.FindElement(key).Update('')
            window['combo0'].update('Yes')
            window['combo1'].update('Yes')
            window['combo2'].update('Yes')
            window['combo3'].update('Yes')
            window['combo4'].update('Yes')
            window['combo5'].update('No')
            window['combo6'].update('No')

    if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT:
        win_close_alert = sg.PopupYesNo("Do you want to save your progress?")
        if win_close_alert == 'Yes':
            for key in keys:
                df.loc[0].at['IN DAY'] = values['In_Monday']
                df.loc[0].at['OUT DAY'] = values['Out_Monday']
                df.loc[0].at['LUNCH'] = values['combo0']
                df.loc[1].at['IN DAY'] = values['In_Tuesday']
                df.loc[1].at['OUT DAY'] = values['Out_Tuesday']
                df.loc[1].at['LUNCH'] = values['combo1']
                df.loc[2].at['IN DAY'] = values['In_Wednesday']
                df.loc[2].at['OUT DAY'] = values['Out_Wednesday']
                df.loc[2].at['LUNCH'] = values['combo2']
                df.loc[3].at['IN DAY'] = values['In_Thursday']
                df.loc[3].at['OUT DAY'] = values['Out_Thursday']
                df.loc[3].at['LUNCH'] = values['combo3']
                df.loc[4].at['IN DAY'] = values['In_Friday']
                df.loc[4].at['OUT DAY'] = values['Out_Friday']
                df.loc[4].at['LUNCH'] = values['combo4']
                df.loc[5].at['IN DAY'] = values['In_Saturday']
                df.loc[5].at['OUT DAY'] = values['Out_Saturday']
                df.loc[5].at['LUNCH'] = values['combo5']
                df.loc[6].at['IN DAY'] = values['In_Sunday']
                df.loc[6].at['OUT DAY'] = values['Out_Sunday']
                df.loc[6].at['LUNCH'] = values['combo6']
            df.to_csv('tablev5.csv', index=False)
            break
        if win_close_alert == 'No':
            break

#df.loc[0].at['IN DAY'] = "8 am"
print(df)
window.close()


