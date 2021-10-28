# Selenium requires the use of a webdriver to interact with Chrome. 
from selenium import webdriver

# This is an error exception for when the class located in the excel file 
# could not be located in the LMS. 
from selenium.common.exceptions import NoSuchElementException

# I used Keys to send the enter command. Our security team did not want me to 
# include passwords so I allowed for passwords to be typed manually after the
# website was brought up.
from selenium.webdriver.common.keys import Keys

# xlrd is how we interact with Excel where the .csv is stored with the class
# information
import xlrd

# Time is required for tracking the start and end of the process. 
import time

# start time enables the start of the timer function meant to track how long
# it takes the entire process to run. 
start_time = time.time()

# This is where the chromedriver is stored on my machine. 
PATH = "C:\Program Files (x86)\chromedriver.exe"

# This mentions that the driver is for chrome. 
driver = webdriver.Chrome(PATH)

# This is where the log file is stored on the machine, and it's in the writable
# state
f = open("C:\\Users\\username\\Desktop\\logfile.txt", "w")

# This pulls up the website for the LMS.
driver.get("https://university-lms-website")

# This prompts the user to enter their details in the python console. 
input("Enter Username and Password, then press enter...")
# return to python console and press Enter
print("You're In!")

# Access the Excel workbook named courselist.xls and open the sheet titled Sheet1
workbook = xlrd.open_workbook('\\Users\username\Desktop\courselist.xls')
worksheet = workbook.sheet_by_name('Sheet1')

# Places the "cursor" at cell A1
num_rows = worksheet.nrows - 1
curr_row = -1

# As long as there are more rows, add another row to the total.
# The while loop runs all the code for what I call the "Interactables"
# This is the part of the code that behaves like a human. 
while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    course = worksheet.cell(curr_row, 0).value
    #print(curr_row)
# This prints the current row/class in the python console for monitoring
    print(course)

# The try except loop searches for class. 
    try:
# This searches for the waffle icon and clicks on it. 
        waffle = driver.find_element_by_css_selector("[text='Select a course...']").click()
        # Search for the course
        
# Waits three seconds to search for the search bar. 
	time.sleep(3)
        sb = driver.find_element_by_tag_name("input-search")
        
# This clicks on the search bar. 
	sb.click()
# This enters the course number from the excel sheet and presses Enter to search. 
        sb.send_keys(str(course), Keys.ENTER)

# Click on the course
        time.sleep(3)
# This clicks on the course hyperlink as it is pulled up by the search bar. 
        ClickCourse = driver.find_element_by_class_name("course-selector-item-name").click()

# This except is for when a course is found on the Excel doc, but not in the 
# list from the database management system.
    except NoSuchElementException as err:
# This prints to the console that the course is not found.         
#print("course not found.")

# The page is refreshed, slept for 3 seconds, and the course is recorded to the
# log file. 
        driver.refresh()
        time.sleep(3)
        #print({course}, "not found.")
        f.write("Course number: %s: status - not found\n" % course)

# Here the course tools is selected in the found course. 
    else:
        # Click on Course Tools
        time.sleep(3)
# The course tools is clicked.         
	courseTools = driver.find_element_by_class_name("navigation-s-group-text")
        courseTools.click()

# Here the hyperlink for Edit course is selected, then clicked. 
        # Click on Edit Course
        time.sleep(3)
        editCourse = driver.find_element_by_css_selector("[text='Edit Course']")
        editCourse.click()

# Here the hyperlink for course offereing is selected, then clicked. 
        # Click on Course Offering
        time.sleep(3)
        courseOffering = driver.find_element_by_xpath("/html/body/div[3]/div/div[3]/ul[1]/li[1]/a")
        courseOffering.click()
 
# Scroll to sections checkbox, wait 3 seconds. 
        time.sleep(3)
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")

        time.sleep(3)
        sectionsCheck = driver.find_element_by_xpath(
                "/html/body/div[3]/div/div[3]/div/div/div/form/div/table[13]/tbody/tr[2]/td/div/input")

# This searches for the checkbox and engages the if else statement.
# The if statement checks whether the box is checked, then unchecks it, 
# Or if it is already unchecked, moves on.  
	checkbox = driver.find_element_by_name("hasSections")
        if checkbox.is_selected():
            print("Checkbox is selected. Unchecking Checkbox")
            sectionsCheck.click()
            print("Unchecking Complete! Moving to Groups.")
            saveSections = driver.find_element_by_xpath(
                "/html/body/div[3]/div/div[3]/div/div/floating-buttons/button[1]")
            saveSections.click()
        else:
            print("Checkbox is not selected. Moving to Groups.")

# Click on Communication Tools after a 3 second wait.
        time.sleep(3)
        communicationTools = driver.find_element_by_xpath(
                "/html/body/header/nav/navigation/navigation-main-footer/div/div/div[4]/dropdown/button/span/span")
        communicationTools.click()

# Click on Groups after a 3 second wait. 
        time.sleep(3)
        clickGroups = driver.find_element_by_css_selector("[text='Groups']")
        clickGroups.click()

        try:
# find Groups Checkbox. If no checkbox is found, the program continues.
# If groups is found, delete it. 
            time.sleep(3)
            sectionCheckbox = driver.find_element_by_xpath(
                "/html/body/div[3]/div/div[3]/div/div/div[2]/form/div/div[1]/table-wrapper/table/tbody/tr[2]/td/table/tbody/tr/td[1]/input")
            sectionCheckbox.click()

        except NoSuchElementException as err:
            print("No Groups found. Moving to next course.")
            driver.refresh()
            time.sleep(3)

        else:
# click Delete Button. 
            time.sleep(3)
            sectionDelete = driver.find_element_by_css_selector("[text='Delete']")
            sectionDelete.click()

# confirm delete of sections by clicking delete once again. 
            time.sleep(3)
            sectionDelete = driver.find_element_by_xpath("/html/body/div[5]/div/div[1]/table/tbody/tr/td[1]/button[1]")
            sectionDelete.click()
            #print("Groups deleted. Moving to next course.")
# The course number is noted, and course complete is noted in the log file. 
        finally:
            f.write("Course number: %s: status - complete\n" % course)
# ^ This is where the script ends for the current class and moves to the next one. 

# Prints to python console the conclusion of the project. 
print("Success. No more courses. Operation Checkmate completed! \n")
# Writes to the log file the conclusion of the project. 
f.write("Success. No more courses. Operation Checkmate completed! \n")
# Closes the time function. 
stop_time = time.time()
# Calculates the length of Checkmate. 
timer = stop_time - start_time

# Prints to the python console. 
print("Operation Checkmate completed in %s seconds" % timer)
# Prints to the log file. 
f.write("Operation Checkmate completed in %s seconds" % timer)
