import time
import random
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import os

# Function to initialize the WebDriver
def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

# Function to save the Excel workbook
def save_excel(workbook, file_path):
    workbook.save(file_path)

# Function to scrape app data
def scrape_app_data(driver, ws):
    # Navigate to the correct URL
    driver.get("https://app.sensortower.com/aso/keyword-research?app_id=482066631&os=ios&country=US&term=")

    # Wait for the page to load
    time.sleep(3 + random.randint(1, 3))

    # Enter email address
    email_input = driver.find_element(By.ID, "email")
    email_input.send_keys("ercinkerem54@gmail.com")

    # Click the "Next" button
    next_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit'][value='Next']")
    next_button.click()

    # Wait for the password screen to load
    time.sleep(3 + random.randint(1, 3))

    # Enter password
    password_input = driver.find_element(By.ID, "password")
    password_input.send_keys("Kerembumbum1")

    # Click the "Sign In" button
    sign_in_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit'][value='Sign In']")
    sign_in_button.click()

    # Wait for the login process to complete
    time.sleep(5 + random.randint(1, 5))

    # Find the search input box
    search_input = driver.find_element(By.ID, "research-keywords-input")

    # Enter the search term
    search_input.send_keys("bmr")

    # Click the "Research" button
    search_button = driver.find_element(By.CSS_SELECTOR, "button.universal-flat-button.universal-flat-button-green.research-keyword-button")
    search_button.click()

    # Wait for the results to load
    time.sleep(3 + random.randint(1, 3))

    # Find all <tr> elements
    rows = driver.find_elements(By.XPATH, "/html/body/div[3]/section[2]/div[3]/table/tbody/tr")

    # Process the first 5 <tr> elements
    for i, row in enumerate(rows[:5]):
        try:
            link = row.find_element(By.TAG_NAME, "a")
            actions = ActionChains(driver)
            actions.move_to_element(link).click().perform()
            time.sleep(3 + random.randint(1, 3))
            driver.switch_to.window(driver.window_handles[1])

            app_name = driver.find_element(By.CSS_SELECTOR, "h2.MuiTypography-root.MuiTypography-h2.AppOverviewSubappDetailsHeader-module__appName--hKwll.css-1fjsqn2").text
            developer_name = driver.find_element(By.CSS_SELECTOR, "a.MuiTypography-root.MuiTypography-inherit.MuiLink-root.MuiLink-underlineHover.BaseLink-module__link--ZB6lH.css-320qng[href^='/ios/publisher/publisher']").text
            app_store_link = driver.find_element(By.CSS_SELECTOR, "a.MuiTypography-root.MuiTypography-inherit.MuiLink-root.MuiLink-underlineHover.BaseLink-module__link--ZB6lH.AppOverviewSubappDetailsHeader-module__viewInStoreLink--u1imV.css-320qng").get_attribute("href")
            support_url = driver.find_element(By.CSS_SELECTOR, "a.MuiTypography-root.MuiTypography-inherit.MuiLink-root.MuiLink-underlineHover.BaseLink-module__link--ZB6lH.css-320qng[href^='http']").get_attribute("href")
            revenue = driver.find_element(By.CSS_SELECTOR, "span.MuiTypography-root.MuiTypography-h1.AppOverviewKpiStatLink-module__link--wxTKU.css-1an6zfe[aria-labelledby='app-overview-unified-kpi-revenue']").text

            release_date = ""
            lis = driver.find_elements(By.CSS_SELECTOR, "li.MuiListItem-root.MuiListItem-gutters.MuiListItem-padding.AppOverviewSubappAboutReleaseDetails-module__listItem--hx2Qp.css-1uvwvgz")
            for li in lis:
                h4_text = li.find_element(By.CSS_SELECTOR, "h4.MuiTypography-root.MuiTypography-h4.css-fmy8y8").text
                if h4_text == "Worldwide Release Date:":
                    release_date = li.find_element(By.CSS_SELECTOR, "p.MuiTypography-root.MuiTypography-body1.AppOverviewSubappAboutReleaseDetails-module__value--PIlcg.css-fcoq8a").text
                    break

            # Write data to the Excel sheet
            ws.append([app_name, developer_name, app_store_link, support_url, revenue, release_date])
            save_excel(wb, excel_file)
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(1 + random.randint(1, 2))
        except Exception as e:
            print(f"Error: {e}")
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            continue

if __name__ == "__main__":
    excel_file = "app_data.xlsx"
    if os.path.exists(excel_file):
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Application Name", "Developer", "App Store Link", "Support URL", "Revenue", "Release Date"])

    driver = init_driver()
    scrape_app_data(driver, ws)
    driver.quit()
