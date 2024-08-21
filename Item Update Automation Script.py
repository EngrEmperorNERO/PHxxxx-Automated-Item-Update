from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from getpass import getpass
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from datetime import date
from datetime import timedelta
import os
import fnmatch
import shutil
import time
from selenium.webdriver.firefox.options import Options
import time
import os
from datetime import date, timedelta
import pandas as pd

firefox_binary_path = 'C:\\Program Files\\Mozilla Firefox\\firefox.exe'  # Update this path to your Firefox executable location

profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", 'Gdrive File Path')
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip,application/octet-stream,application/download,application/msexcel,text/csv,application/csv,application/comma-separated-values,text/plain")
#profile.set_preference("browser.download.dir", 'C:\\Users\\Owner\\Downloads\\')

geckodriver_path = r'C:\Users\Zemo\Documents\Programs\firefoxddriver\geckodriver.exe'  # Update this path to your geckodriver executable location
service = Service(executable_path=geckodriver_path)

#url for SAR report
firefox_options = Options()
firefox_options.binary_location = firefox_binary_path
firefox_options.profile = profile

# Create the WebDriver instance
driver = webdriver.Firefox(service=service, options=firefox_options)
#driver = webdriver.Firefox(options=options, firefox_profile=profile) #for profile setup
driver.get("Downlaoded csv from source file")

driver.maximize_window()

#download csv
driver.find_element(By.XPATH,'Okta Auth Xpath').click()
time.sleep(3)
#username input
username = driver.find_element(By.ID,'username box xpath')  #//*[@id="input28"] #"okta-signin-username"
username.clear()
username.send_keys("email")
#password input
password = driver.find_element(By.ID,'password box xpath') 
password.clear()
password.send_keys("password")

driver.implicitly_wait(5)

driver.find_element(By.XPATH,'/html/body/div[2]/div[2]/main/div[2]/div/div/div[2]/form/div[2]/input').click() 
#send push
driver.find_element(By.XPATH,'/html/body/div[2]/div[2]/main/div[2]/div/div/div[2]/form/div[2]/div/div[2]/div[2]/div[2]/a').click()
#close pop-up
#driver.maximize_window()
driver.implicitly_wait(5)

# xpath for popups //*[@id="ReleaseModalClose"]
try:
    popup=WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ReleaseModalClose"]')))
    popup.click()
except:
    pass

try:
    popup=WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.cssSelector,'#ReleaseModalClose')))
    popup.click()
except:
    pass

driver.implicitly_wait(2)
time.sleep(35)

export = driver.find_element(By.ID,"Element ID")
driver.execute_script("arguments[0].click();",export)
#select report for yesterday
driver.implicitly_wait(5)

driver.implicitly_wait(5)
time.sleep(200)
#change file name
CleanPath = "File Path"
# Get today's date
today = date.today()
yesterday = today - timedelta(days = 1)
today_string =str(today)
yesterday_string = str(yesterday)
year = today_string[0:4]
month = today_string[5:7]
day_today = today_string[8:10]
for i in os.listdir(CleanPath):
    if 'Sample File Name' in i:
        prevName =CleanPath + "\\" +i
        newName = CleanPath + "\\"+ "Pricing_Report_" + month + "." + day_today+ "." + year+ ".csv"
        os.rename(prevName,newName)

#Download File from other Source
service = Service(executable_path=r'C:\Users\Zemo\Documents\Chrome drivers\Version 127\chromedriver-win64 (4)\chromedriver-win64\chromedriver.exe')
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
prefs = {"download.default_directory" : 'Gdrive File Path'}
options.add_experimental_option("prefs",prefs)
driver = webdriver.Chrome(service=service, options=options)

#url for Activity Comments Export
driver.get("Link to the source file")
driver.maximize_window()
#driver.maximize_window()

#username input
driver.implicitly_wait(5)
#username = driver.find_element_by_id("username")
username =driver.find_element(By.ID,"username")
username.clear()
username.send_keys("username")
#password input
password = driver.find_element(By.ID,"password")
password.clear()
password.send_keys("password")
# #select login
driver.implicitly_wait(5)
driver.find_element(By.ID,"Login").click()
driver.implicitly_wait(20)
driver.switch_to.frame(driver.find_element(By.XPATH,'/html/body/div[4]/div[1]/section/div[1]/div[2]/div[2]/div[1]/div/div/div/div/iframe'))
driver.find_element(By.XPATH,'/html/body/div[9]/div/div[1]/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[6]/div/div/button').click()
time.sleep(3)
driver.find_element(By.XPATH,'/html/body/div[9]/div/div[1]/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[6]/div/div/div/ul/li[4]/a').click()
driver.switch_to.default_content()
#detils only
driver.find_element(By.XPATH,'/html/body/div[4]/div[2]/div/div[2]/div/div[2]/div/fieldset/div/div[2]/label/span[2]/span/span[2]').click()
#csv
driver.find_element(By.XPATH,'/html/body/div[4]/div[2]/div/div[2]/div/div[2]/div/div/div[1]/div/div/select/option[3]').click()
time.sleep(15)
#Export
driver.find_element(By.XPATH,'/html/body/div[4]/div[2]/div/div[2]/div/div[3]/button[2]/span').click()
time.sleep(50)

#change file name
CleanPath = "Gdrive File Path"
# Get today's date
today = date.today()
yesterday = today - timedelta(days = 1)
today_string =str(today)
yesterday_string = str(yesterday)
year = today_string[0:4]
month = today_string[5:7]
day_today = today_string[8:10]
for i in os.listdir(CleanPath):
    if 'report' in i:
        prevName =CleanPath + "\\" +i
        newName = CleanPath + "\\"+ "SF_Item_Report_for_update" + month + "." + day_today+ "." + year+ ".csv"
        os.rename(prevName,newName)

#Download File from another Website
CleanPath = "Gdrive File Path"
CleanPath_sfitems = "Gdrive File Path"

# Get today's date
today = date.today()
yesterday = today - timedelta(days=1)
today_string = str(today)
yesterday_string = str(yesterday)
year = today_string[0:4]
month = today_string[5:7]
day_today = today_string[8:10]

for i in os.listdir(CleanPath):
    #if "Pricing_Report_" + month + "." + day_today + "." + year in i:
    if "Pricing_Report_" + month + "." + day_today + "." + year in i:
        print(i)
        df = pd.read_csv(CleanPath + "\\" + i)
        # Remove last row that just has text
        df.drop(df.tail(1).index, inplace=True)
        df = df[(df['Level1'].notna()) & (df['Level1'] != 'Vet')]
        print(df)
        # Save the cleaned DataFrame to a new CSV file
        output_file_path = (
            "Output File Processed"
        )
        df.to_csv(output_file_path, index=False, header=True)

#Clean the 1st downloaded File.
#Next script
CleanPath = "Gdrive File Path"
CleanPath_sfitems = "Gdrive File Path"

# Get today's date
today = date.today()
yesterday = today - timedelta(days=1)
today_string = str(today)
yesterday_string = str(yesterday)
year = today_string[0:4]
month = today_string[5:7]
day_today = today_string[8:10]

for i in os.listdir(CleanPath):
    if "Pricing_Report_" + month + "." + day_today + "." + year in i:
        print(i)
        df = pd.read_csv(CleanPath + "\\" + i)
        df.drop(df.tail(1).index, inplace=True)
        df = df[(df['Level1'].notna()) & (df['Level1'] != 'Vet')]
        print(df)
        output_file_path = (
            "Out file - Processed"
        )
        df.to_csv(output_file_path, index=False, header=True)


#Check if the item is on 1st downloaded file vs. Another File
# Get today's date
today = date.today()
yesterday = today - timedelta(days=1)
today_string = str(today)
yesterday_string = str(yesterday)
year = today_string[0:4]
month = today_string[5:7]
day_today = today_string[8:10]

# Define file paths
cleaned_file_path = "Google Drive File Path"
sf_item_report_path = r"Google Drive File Path"
output_file_path = r"Google Drive File Path"

# Read the cleaned file
cleaned_df = pd.read_csv(cleaned_file_path)

# Read the SF Item Report file
sf_item_df = pd.read_csv(sf_item_report_path,encoding='unicode_escape')

# Create a new column named 'find'
cleaned_df['find'] = 'Not Found'  # Default value is 'Not Found'

# Update the 'find' column based on the VLOOKUP-like operation
for index, row in cleaned_df.iterrows():
    item_number = row['Item Number']
    
    # Check if the item number is present in the SF Item Report
    match_row = sf_item_df[sf_item_df['Item: Item Name'] == item_number]
    
    if not match_row.empty:
        cleaned_df.at[index, 'find'] = match_row.iloc[0]['Item: Item Name']

cleaned_df.to_csv(output_file_path, index=False, encoding='utf-8')

# Display the updated DataFrame
print(cleaned_df)

#Next script
#Save the file that contains the item on 1st downloaded file on the 2nd downloaded file.
# Get today's date
today = date.today()
yesterday = today - timedelta(days=1)
today_string = str(today)
yesterday_string = str(yesterday)
year = today_string[0:4]
month = today_string[5:7]
day_today = today_string[8:10]

# Define file path
sf_item_report_path = r"Google Drive File Path"

# Read the SF Item Report file
sf_item_df = pd.read_csv(sf_item_report_path, encoding='latin1')

# Check the 'find' column and save a copy if value is "Not Found"
if 'find' in sf_item_df.columns:
    not_found_df = sf_item_df[sf_item_df['find'] == 'Not Found']
    
    if not not_found_df.empty:
        # Generate the output file path
        output_file_path = r"file path"

        # Save a copy of the DataFrame with "Not Found" entries
        not_found_df.to_csv(output_file_path, index=False, encoding='utf-8')

        print(f"Copy of the file saved at: {output_file_path}")
    else:
        print("No rows with 'Not Found' in the 'find' column.")
else:
    print("The 'find' column is not present in the DataFrame.")


#Map the Item Category
#Note: Upload only THE PROD_CAT_DESC NOT UNKNOWN CAT
#AND ACTIVE NOT FALSE
# Define file paths

# Get today's date
today = date.today()
yesterday = today - timedelta(days=1)
today_string = str(today)
yesterday_string = str(yesterday)
year = today_string[0:4]
month = today_string[5:7]
day_today = today_string[8:10]
input_file_path = r"file path"
output_file_path = r"file path"

# Read the CSV file
df = pd.read_csv(input_file_path)

# Rename 'Item Number' column to 'Product Code'
df.rename(columns={'Item Number': 'Product Code'}, inplace=True)

# Create a new column 'Item Name' and set its value equal to 'Product Code'
df['Item Name'] = df['Product Code']

# Add a new column 'Active' and set its value to True for all rows
df['Active'] = True

# Select necessary columns
selected_columns = [
    'Product Code',
    'Item Name',
    'Item Description',
    'Level1',
    'Level2',
    'Level3',
    'Level4',
    'Level5',
    'Local Service Price',
    'Stocking Price',
    'Min Qty',
    'Active'  # Include 'Active' column in the selected columns
]

# Rename columns as specified
column_mapping = {
    'Level1': 'V',
    'Level2': 'IV',
    'Level3': 'III',
    'Level4': 'II',
    'Level5': 'I'
}

df = df[selected_columns].rename(columns=column_mapping)
df['V'] = df['V'].replace('Extremities & Trauma', 'Distal Extremities')

def set_prod_cat_desc(row):
    if row['V'] == 'Arthroplasty':
        return 'PLASTY'
    elif row['V'] == 'Biologics':
        return 'BIOLOGICS'
    elif row['V'] == 'Capital Consumables':
        return 'CAPITAL'
    elif row['V'] == 'Knee & Hip':
        return 'KNEE'
    elif row['V'] == 'Loan Fees':
        return 'LOAN FEES'
    elif row['V'] == 'Shoulder':
        return 'SHOULDER'
    elif row['V'] == 'Spine':
        return 'SPINE'
    elif row['V'] == 'Tissue':
        return 'TISSUE'
    elif row['V'] == 'Upper Extremity':
        return 'SHOULDER'
    elif row['V'] == 'Imaging & Resection':
        if row['III'] == 'Repair Charges':
            return 'SERVICE & REPAIR'
        elif row['III'] == 'Video':
            return 'IMAGING & RESECTION'
    elif row['V'] == 'Distal Extremities':
        if row['III'] in ['Long Bone Trauma', 'Small Bone Trauma', 'Trauma']:
            return 'TRAUMA'
        else:
            return 'SMALL JOINT'
    
    return 'UNKNOWN CATEGORY'

df['Prod_Cat_Desc'] = df.apply(set_prod_cat_desc, axis=1)

# Save the modified DataFrame to a new CSV file
df.to_csv(output_file_path, index=False)

print(f"Modified file saved at: {output_file_path}")
