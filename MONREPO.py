from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import csv
import re
import time

driver = webdriver.Chrome()
driver.get("https://billing.cgnet.com.np/h8ssrms/Login.aspx")

# Store the ID of the main login window
main_window = driver.current_window_handle

driver.find_element(By.NAME, "txtUserName").send_keys("laxman.koirala")
driver.find_element(By.NAME, "txtPassword").send_keys("koirala.laxman")
driver.find_element(By.ID, "save").click()

wait = WebDriverWait(driver, 20)
wait.until(lambda d: "Login.aspx" not in d.current_url)

# --- POPUP HANDLING BLOCK ---
# Get all open window handles
all_windows = driver.window_handles

if len(all_windows) > 1:
    for window in all_windows:
        if window != main_window:
            driver.switch_to.window(window)
            driver.close() # Close the popup
    
    # Switch back to the main window to continue
    driver.switch_to.window(main_window)
# -----------------------------

# Navigate directly to the audit page after login
driver.get("https://billing.cgnet.com.np/h8ssrms/Auditpage.aspx")

wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_combomodul")))

# Select module and operation
Select(driver.find_element(By.ID, "ContentPlaceHolder1_combomodul")).select_by_visible_text("Case")
time.sleep(2)  
activity_element = driver.find_element(By.ID, "ContentPlaceHolder1_comboactivity")
Select(activity_element).select_by_visible_text("Update")
time.sleep(2)  

# Select DateWise
datewise_radio = driver.find_element(By.ID, "ContentPlaceHolder1_rdbdat_0")
datewise_radio.click()
time.sleep(1)

# Set dates
from_date = "14 Apr 2026"
to_date = "14 May 2026"
from_date_input = driver.find_element(By.ID, "ContentPlaceHolder1_txtfrmdate")
from_date_input.clear()
from_date_input.send_keys(from_date)
from_date_input.send_keys("\t")
time.sleep(0.5)

to_date_input = driver.find_element(By.ID, "ContentPlaceHolder1_txttodate")
to_date_input.clear()
to_date_input.send_keys(to_date)
to_date_input.send_keys("\t")
time.sleep(0.5)

# Select user and records per page
Select(driver.find_element(By.ID, "ContentPlaceHolder1_combouserddl")).select_by_visible_text("Abhishek Poudyal")
time.sleep(2)  
Select(driver.find_element(By.ID, "ContentPlaceHolder1_ddlPageSize")).select_by_visible_text("200 Records")

# Search
search_button = driver.find_element(By.ID, "ContentPlaceHolder1_btnserch")
search_button.click()

# Allow results to load
wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#tabledata tr, #ContentPlaceHolder1_gvAudit tbody tr')))
time.sleep(2)

rows = []
selectors = [
    "#ContentPlaceHolder1_gvAudit tbody tr",
    "#gvAudit tbody tr",
    "#tabledata tbody tr",
    "#ContentPlaceHolder1_UpdatePanel1 tbody tr",
    "table#ContentPlaceHolder1_gvAudit tbody tr",
    "table[id*='gv'] tbody tr",
    "table tbody tr"
]

row_date_re = re.compile(r'^\d{2} [A-Za-z]{3} \d{4} \d{1,2}:\d{2} (AM|PM)$')

rows = []
working_selector = None
for selector in selectors:
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        print(f"Trying selector: {selector}, found {len(elements)} elements")
        temp_rows = []
        for row in elements:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 6:
                first_cell = cells[0].text.strip()
                if row_date_re.match(first_cell):
                    basic_data = [
                        cells[0].text.strip(),
                        cells[1].text.strip(),
                        cells[2].text.strip(),
                        cells[3].text.strip(),
                        cells[4].text.strip(),
                        cells[6].text.strip(),
                    ]
                    temp_rows.append((row, basic_data))
                    if len(temp_rows) <= 5:
                        print(f"Collected row {len(temp_rows)}: {first_cell} - {basic_data[5]}")
        if temp_rows:
            rows = temp_rows
            working_selector = selector
            print(f"Found {len(rows)} audit rows with selector {selector}")
            break
    except Exception as e:
        print(f"Error with selector {selector}: {e}")
        continue

if not rows:
    print("No audit results found.")
else:
    # Extract popup details for each row
    complete_rows = []
    for i, (row, basic_data) in enumerate(rows):
        popup_details = ""
        try:
            # Re-find the current row element in case the DOM refreshed
            current_rows = driver.find_elements(By.CSS_SELECTOR, working_selector)
            if i < len(current_rows):
                row = current_rows[i]

            view_link = None
            try:
                view_link = row.find_element(By.CSS_SELECTOR, "a[id*=LinkButtonView]")
            except Exception:
                links = row.find_elements(By.TAG_NAME, "a")
                if links:
                    view_link = links[0]

            if view_link:
                driver.execute_script("arguments[0].scrollIntoView(true);", view_link)
                driver.execute_script("arguments[0].click();", view_link)
                print(f"Clicked popup for row {i+1}")

                try:
                    wait.until(EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_ifrm')))
                    iframe = driver.find_element(By.ID, 'ContentPlaceHolder1_ifrm')
                    driver.switch_to.frame(iframe)
                    popup_details = driver.find_element(By.TAG_NAME, 'body').text.strip()
                    driver.switch_to.default_content()
                except Exception:
                    try:
                        modal = driver.find_element(By.CSS_SELECTOR, 'div.ui-dialog-content, div.modal-body, #dialog')
                        popup_details = modal.text.strip()
                    except Exception:
                        popup_details = ''

                try:
                    close_button = driver.find_element(By.ID, 'ContentPlaceHolder1_btnCancelD')
                    driver.execute_script("arguments[0].click();", close_button)
                except Exception:
                    try:
                        close_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Close') or contains(text(), 'close')]")
                        driver.execute_script("arguments[0].click();", close_button)
                    except Exception:
                        try:
                            close_button = driver.find_element(By.CSS_SELECTOR, ".close")
                            driver.execute_script("arguments[0].click();", close_button)
                        except Exception:
                            from selenium.webdriver.common.keys import Keys
                            driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(1)
            else:
                print(f"No popup link found for row {i+1}")
                popup_details = ''
        except Exception as e:
            print(f"Error extracting popup for row {i+1}: {e}")
            popup_details = 'Error extracting details'

        complete_rows.append(basic_data + [popup_details])

    rows = complete_rows

with open("tickets.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Timestamp", "Ticket ID", "User", "Module", "Action", "Details", "Popup Details"])
    writer.writerows(rows)

# Close the session
driver.quit()

# Automatically generate categorized multi-tab Excel report
try:
    from extract_tickets_excel import extract_and_categorize_tickets
    print("\nCategorizing extracted tickets into multi-tab Excel report...")
    extract_and_categorize_tickets()
except Exception as err:
    print(f"Error generating Excel report: {err}")