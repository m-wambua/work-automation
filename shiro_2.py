# Find the applicants on the page
applicants = browser.find_elements(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div[2]/div/ul/li')

# Create an empty DataFrame to store the data
df = pd.DataFrame()

# Loop through the applicants
for applicant in applicants:
    # Click on the applicant to activate it
    applicant.click()
    
    # Wait for the content of the applicant to load
    WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div[2]/div/div')))
    
    # Find the tabs for this applicant
    tabs = browser.find_elements(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div[2]/div/ul/li')
    
    # Create an empty dictionary to store the data for this applicant
    applicant_data = {}
    
    # Loop through the tabs for this applicant
    for tab in tabs:
        # Click on the tab to activate it
        tab.click()
        
        # Wait for the content of the tab to load
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div[2]/div/div/div[2]/div/form')))
        
        # Find the div elements within the active tab
        divs = browser.find_elements(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div[2]/div/div/div[2]/div/form//div')
        
        # Create an empty list to store the data for this tab
        tab_data = []
        
        # Loop through the div elements
        for i, div in enumerate(divs):
            # Extract the text from the div element
            text = div.text
            
            # If this is the first div element (i.e., the header), use its text as the column name
            if i == 0:
                column_name = text
            else:
                # Otherwise, append the text to the tab_data list
                tab_data.append(text)
        
        # Store the data for this tab in the applicant_data dictionary using the column name as the key and the data as the value
        applicant_data[column_name] = ' '.join(tab_data)
    
    # Append a new row to the DataFrame using this applicant's data
    df = df.append(applicant_data, ignore_index=True)

# At this point, you can write the DataFrame to an Excel file like this:
df.to_excel('output.xlsx', index=False)


#############################################################################
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd

# Set up the webdriver
browser = webdriver.Chrome()

# Navigate to the desired page
browser.get('https://www.example.com')

# Define the root XPATH
root_xpath = '/html/body/div[2]/div/div/div[2]/div/div[2]/div/div/div[2]/div/form'

# Define a DataFrame to store the extracted data
df_data = pd.DataFrame()

for tab in tabs:
    tab.click()
    WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, root_xpath)))
    
    # Locate all the input elements that are children of div elements under the root XPATH
    input_elements = browser.find_elements_by_xpath(f'{root_xpath}//div/input')
    
    # Define a dictionary to store the extracted data for the current tab
    applicant_data = {}
    
    # Loop through the located input elements
    for i, input_element in enumerate(input_elements):
        # Extract the value of the current input element
        input_value = input_element.get_attribute('value')
        
        # Store the extracted value in the dictionary using a column name based on the index of the input element
        applicant_data[f'Input {i + 1}'] = input_value
    
    # Append the extracted data for the current tab to the DataFrame
    df_data = pd.concat([df_data, pd.DataFrame(applicant_data, index=[0])], ignore_index=True)

# Write the DataFrame to an Excel file
df_data.to_excel('input_values.xlsx', index=False)


#================================================================================================
# Create an empty list to store the extracted data for all applicants
all_applicants_data = []

# Loop through the rows of the column containing the 'review?ref=' strings
for row in df.iloc[2:,0]:
    # Split the string at the equal sign to extract the value
    value = row
    # Do something with the value
    my_string = 'https://portal.erc.go.ke:8446/new-licence-applications/review?ref='
    my_integer = row

    # Use an f-string to insert the integer into the string
    result = f'{my_string}{my_integer}'
    data["Links"].append(result)
    browser.get(result)
    
    # Define a dictionary to store the extracted data for the current applicant
    applicant_data = {}
    
    tabs=browser.find_elements(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/div/ul/li')
      
    # Define the root XPATH
    root_xpath = '/html/body/div[2]/div/div/div[2]/div/div[2]/div/div/div[2]/div/form'

    for tab in tabs:
        tab.click()
        time.sleep(2)
        # Locate all the input elements that are children of div elements under the root XPATH
        input_elements = browser.find_elements(By.XPATH,f'{root_xpath}//div/input')
        
        # Loop through the located input elements
        for i, input_element in enumerate(input_elements):
            # Extract the value of the current input element
            input_value = input_element.get_attribute('value')
            
            # Store the extracted value in the dictionary using a column name based on the index of the input element
            applicant_data[f'Input {i + 1}'] = input_value
    
    # Append the extracted data for the current applicant to the list
    all_applicants_data.append(applicant_data)

# Write the list of dictionaries to a text file
with open('output_values.txt', 'w') as f:
    for applicant_data in all_applicants_data:
        f.write(str(applicant_data) + '\n')



###################################################################
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains

import time
import pandas as pd

options = Options()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\ip 5\mike\work_related\data",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

browser = webdriver.Chrome(executable_path=r'C:\Users\ip 5\mike\work_related\chromedriver.exe',chrome_options=options)
browser.get('https://portal.erc.go.ke:8446/site/login')
            
username = browser.find_element(By.NAME,'LoginForm[username]')
password = browser.find_element(By.NAME,'LoginForm[password]')

username.send_keys("Kiprotich.Bii")
password.send_keys("Sup.Reem23!")
time.sleep(15)
login_button = browser.find_element(By.NAME,'login-button')
login_button.click()

# Navigating to the electricity department menu
time.sleep(10)
actions=ActionChains(browser)
main_menu=browser.find_element(By.XPATH,'/html/body/div[2]/header/nav/div/div/div[2]/ul/li[7]')
actions.move_to_element(main_menu).perform()
sub_menu=browser.find_element(By.XPATH,'/html/body/div[2]/header/nav/div/div/div[2]/ul/li[7]/ul/li[1]')
sub_menu.click()

# Clicking thr Excel Button

excel_button=browser.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[1]/button[1]')
excel_button.click()
time.sleep(5)

# Load the Excel file
df = pd.read_excel('C:/Users/ip 5/mike/work_related/data/ERC Internal Portal.xlsx')
data = {"Links": []}
data_tabs={}
df_data=pd.DataFrame()
applicant_data = {}

# Create an empty list to store the extracted data for all applicants
all_applicants_data = []

# Loop through the rows of the column containing the 'review?ref=' strings
for row in df.iloc[2:,0]:
    # Split the string at the equal sign to extract the value
    value = row
    # Do something with the value
    my_string = 'https://portal.erc.go.ke:8446/new-licence-applications/review?ref='
    my_integer = row

    # Use an f-string to insert the integer into the string
    result = f'{my_string}{my_integer}'
    data["Links"].append(result)
    browser.get(result)
    
    # Define a dictionary to store the extracted data for the current applicant
    applicant_data = {}
    
    tabs=browser.find_elements(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/div/ul/li')
      
    # Define the root XPATH
    root_xpath = '/html/body/div[2]/div/div/div[2]/div/div[2]/div/div/div[2]/div/form'

    for tab in tabs:
        tab.click()
        time.sleep(2)
        # Locate all the input elements that are children of div elements under the root XPATH
        input_elements = browser.find_elements(By.XPATH,f'{root_xpath}//div/input')
        
        # Loop through the located input elements
        for i, input_element in enumerate(input_elements):
            # Extract the value of the current input element
            input_value = input_element.get_attribute('value')
            
            # Store the extracted value in the dictionary using a column name based on the index of the input element
            applicant_data[f'Input {i + 1}'] = input_value
    
    # Append the extracted data for the current applicant to the list
    all_applicants_data.append(applicant_data)

# Create a DataFrame from the list of dictionaries
df_data = pd.DataFrame(all_applicants_data)

# Write the DataFrame to an Excel file
df_data.to_excel('output_values.xlsx', index=False)
