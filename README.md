# PaycomTimeclock
Python GUI for saving working hours and inputting automatically into Paycom. I created this program to solve the minor annoyance of logging into the Paycom site and manually entering in my time. 

# Overview of Program
When launched, Pendulum gets the current week's dates and formats. These dates then get stored in the "tablev5.csv" file. I didn't code an if file doesn't exist check because laziness 🤷‍♂️ 

```Python
# Gets the current week and formats accordingly.
today = pendulum.now()
Monday = today.start_of('week').format('MM/DD/YYYY')
Tuesday = today.start_of('week').add(days=1).format('MM/DD/YYYY')
Wednesday = today.start_of('week').add(days=2).format('MM/DD/YYYY')
Thursday = today.start_of('week').add(days=3).format('MM/DD/YYYY')
Friday = today.start_of('week').add(days=4).format('MM/DD/YYYY')
Saturday = today.start_of('week').add(days=5).format('MM/DD/YYYY')
Sunday = today.start_of('week').add(days=6).format('MM/DD/YYYY')
dates = [Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday]
```

The GUI displays the dates of the current week, punch in/out fields, and whether or not I took lunch. 
<p align="center">
  <img src="https://user-images.githubusercontent.com/95936691/200417102-3073b28f-90e2-44f0-ae79-953d621d9e0a.png" />
</p>

When Run is clicked the values entered into these fields are saved to the CSV file and chromedriver launches to the Paycom login portal. Selenium then passes the credentials stored in the "credentials.json" file. I tried dabbling with encrypting the password so it wasn't in plain text but I eventually decided it wasn't worth the time.

<p align="center">
  <img src="https://user-images.githubusercontent.com/95936691/200422683-992e4f06-4093-43c2-a9c6-671f65dd750c.png" />
</p>

Selenium then passes the credentials stored in the "security_questions.json" file

<p align="center">
  <img src="https://user-images.githubusercontent.com/95936691/200423932-48e74af2-3622-47c8-9985-853c6ed5fdb3.png" />
</p>

Once logged in, Selenium navigates to the "Web Time Sheet Read Only" page and then clicks on the "Add Punch Change Request" button on the bottom right. 

<p align="center">
  <img src="https://user-images.githubusercontent.com/95936691/200426635-69eac496-d31a-45ae-bb8d-ccb86fb756ef.png" width="500" height="350" />
</p>


<p align="center">
  <img src="https://user-images.githubusercontent.com/95936691/200427712-a0613db5-0dbb-406b-a8db-78eca0a47a8e.png" />
</p>

The script then runs through each of the table's entries and enters in the date, selects the type of punch (IN DAY, OUT LUNCH, IN LUNCH, OUT DAY) and enters in the punch time. Before this, I had to manually enter all of this information and obviously the last thing I'd want to do on a Friday evening is spend 10 minutes punching in my time. 

```Python
        for index, row in df.iterrows():
            ##### IN DAY #####
            if row['IN DAY'] == '':
                print("Error for " + row['DATE'] + ": No time entered for IN DAY.")
                sg.Popup("Error for " + row['DATE'] + ": No time entered for IN DAY.", title='Error')
                continue
            wait.until(EC.invisibility_of_element_located((By.XPATH, "//*[@id='punch-change-request-modal']")))
            try:
                RequestNewPunch = driver.find_element_by_name("timecard-add-punch-change-request")
                RequestNewPunch.click()
            except:
                wait.until(EC.visibility_of_element_located((By.NAME, "timecard-add-punch-change-request")))
                RequestNewPunch = driver.find_element_by_name("timecard-add-punch-change-request")
                RequestNewPunch.click()
            pcrDate = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@name='pcrDate']")))
            pcrDate.clear()
            pcrDate.send_keys(row['DATE'])
            driver.find_element_by_xpath("//select[@name='pcrPunchType']/option[text()='IN DAY']").click()
            driver.find_element_by_xpath("//input[@name='pcrTime']").clear()
            driver.find_element_by_xpath("//input[@name='pcrTime']").send_keys(row['IN DAY'])
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'parsleyCloseButton parsleyShow')))
            driver.find_element_by_name("pcrAdd").click()
```            

