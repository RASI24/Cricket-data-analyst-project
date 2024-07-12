from selenium import webdriver
from bs4 import BeautifulSoup
import time
import openpyxl

# Create a workbook and select the active worksheet
workbook=openpyxl.Workbook()
worksheet=workbook.active



#scraping website url
url = 'https://www.moneycontrol.com/news/cricket/cricket-live-score/chennai-super-kings-vs-royal-challengers-bengaluru/5970/ckbc03222024237771.html'

#Path to the ChromeDriver executable
path = r'C:/Program Files (x86)/chromedriver'

#Initialize the WebDriver
browser = webdriver.Chrome(executable_path=path)
browser.get(url)

#Wait for the page to fully load
time.sleep(5)

#Parse the page source with BeautifulSoup
soup = BeautifulSoup(browser.page_source, 'html.parser')

#Close the WebDriver
browser.close()

# Find all tables with the specified class
tables = soup.find_all('table', class_='score-tbl1')

# Access the second table
table = tables[2]
    
# Initialize an empty list to store bowler data
bowlers_data = []

# Find all rows in the table
rows = table.find_all('tr')

    # Iterate over each row
for row in rows:
    # Find all cells in the row
    cells = row.find_all('th')
    # Extract the text content of each cell and strip whitespace
    row_data = [cell.text.strip() for cell in cells]
    # Append non-empty row data to bowlers_data list
    if row_data:
        bowlers_data.append(row_data)

    
# Find all rows in the table
rows = table.find_all('tr')

    # Iterate over each row
for row in rows:
    # Find all cells in the row
    cells = row.find_all('td')
    # Extract the text content of each cell and strip whitespace
    row_data = [cell.text.strip() for cell in cells]
    # Append non-empty row data to bowlers_data list
    if row_data:
        bowlers_data.append(row_data)

# Print the extracted data in list format
for bowler in bowlers_data:
    worksheet.append(bowler)	
#set the location where the file you want to save
file_path = 'C:/Users/RASI/Documents/data_scraping/CSK_BATTERS_RESULT/csk vs rcb.xlsx'

# Save the workbook to the specified location
workbook.save(file_path)

print(f"Workbook saved to {file_path}")
