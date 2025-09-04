import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# ===== CSV Setup =====
csv_filename = r"C:\Users\Lenovo\Desktop\SeleniumB\cisco_data\Japan.csv"
with open(csv_filename, mode="w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow([
        "Index", "Headquarters", "Address", "Partner Website", "Phone Number",
        "Integrator Level", "Specializations", "Providers", "Designations"
    ])

# Setup Chrome options and driver path
chrome_options = Options()
prefs = {"profile.default_content_setting_values.geolocation": 2}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--start-maximized")

service = Service("C:/Users/Lenovo/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe")
driver = webdriver.Chrome(service=service, options=chrome_options)

locations=["Japan"]

try:
    # Open page
    driver.get("https://locatr.cloudapps.cisco.com/WWChannels/LOCATR/openBasicSearch.do")

    # Add locations
    for location in locations:
        input_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "searchLocationInput"))
        )
        input_field.click()
        input_field.clear()
        input_field.send_keys(location)

        suggestion = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span.display-name"))
        )
        suggestion.click()
        time.sleep(1)

    

    # Click search
    search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "searchBtn"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
    time.sleep(0.5)
    search_button.click()

    profile_index = 0

    while True:  # Loop through all pages
        # Wait for results
        WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.XPATH, "//button[contains(text(), 'View Profile')]"))
        )
        time.sleep(2)

        visited = set()

        while True:
            try:
                buttons = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//button[contains(text(), 'View Profile')]"))
                )

                unvisited = [(i, btn) for i, btn in enumerate(buttons) if str(i) not in visited]

                if not unvisited:
                    break

                i, current_button = unvisited[0]
                visited.add(str(i))

                driver.execute_script("arguments[0].scrollIntoView(true);", current_button)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", current_button)

                profile_index += 1
                print(f"\n[{profile_index}] Opened profile.")

                # Wait for company details container
                company_details = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "company-details"))
                )

                # Headquarters
                try:
                    headquarters_elem = company_details.find_element(By.CSS_SELECTOR, "h2.headquartersNew")
                    lines = headquarters_elem.get_attribute("innerText").strip().splitlines()
                    headquarters = " ".join(line.strip() for line in lines if line.strip() and line.strip() != 'A')
                except:
                    headquarters = "N/A"

                # Address
                try:
                    address_elem = company_details.find_element(By.CLASS_NAME, "partnerAddressAlign")
                    address = " ".join(line.strip() for line in address_elem.text.strip().splitlines() if line.strip())
                except:
                    address = "N/A"

                # Partner website
                try:
                    website_elem = company_details.find_element(By.XPATH, ".//a[contains(normalize-space(text()), 'Visit Partner Website')]")
                    partner_website = website_elem.get_attribute("href")
                except:
                    partner_website = "N/A"

                # Phone
                try:
                    phone_elem = company_details.find_element(By.XPATH, ".//a[contains(@class, 'partnerPhone')]")
                    phone_href = phone_elem.get_attribute("href")
                    phone_number = phone_href.replace("tel:", "").strip()
                except:
                    phone_number = "N/A"

                # Specializations
                try:
                    qualifications_sections = WebDriverWait(company_details, 5).until(
                        EC.presence_of_all_elements_located((By.CLASS_NAME, "qualificationsSection"))
                    )

                    try:
                        integrator_list = qualifications_sections[0].find_elements(By.CSS_SELECTOR, "ul.qualificationsData li")
                        integrator_level = [li.text.strip() for li in integrator_list if li.text.strip()]
                    except:
                        integrator_level = []

                    try:
                        specialization_list = qualifications_sections[1].find_elements(By.CSS_SELECTOR, "ul.qualificationsData li")
                        specializations = [li.text.strip() for li in specialization_list if li.text.strip()]
                    except:
                        specializations = []

                    try:
                        provider_list = qualifications_sections[2].find_elements(By.CSS_SELECTOR, "ul.qualificationsData li")
                        providers = [li.text.strip() for li in provider_list if li.text.strip()]
                    except:
                        providers = []

                    try:
                        designation_list = qualifications_sections[3].find_elements(By.CSS_SELECTOR, "ul.qualificationsData li")
                        designations = [li.text.strip() for li in designation_list if li.text.strip()]
                    except:
                        designations = []

                except:
                    integrator_level = []
                    specializations = []
                    providers = []
                    designations = []

                print(f"[{profile_index}] Headquarters: {headquarters}")
                print(f"[{profile_index}] Address: {address}")
                print(f"[{profile_index}] Partner Website: {partner_website}")
                print(f"[{profile_index}] Phone Number: {phone_number}")
                print(f"[{profile_index}] Integrator Level: {', '.join(integrator_level) if integrator_level else 'N/A'}")
                print(f"[{profile_index}] Specializations: {', '.join(specializations) if specializations else 'N/A'}")
                print(f"[{profile_index}] Providers: {', '.join(providers) if providers else 'N/A'}")
                print(f"[{profile_index}] Designations: {', '.join(designations) if designations else 'N/A'}")

                # ===== Save to CSV =====
                with open(csv_filename, mode="a", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow([
                        profile_index,
                        headquarters,
                        address,
                        partner_website,
                        phone_number,
                        ", ".join(integrator_level) if integrator_level else "N/A",
                        ", ".join(specializations) if specializations else "N/A",
                        ", ".join(providers) if providers else "N/A",
                        ", ".join(designations) if designations else "N/A"
                    ])

                # Back to results
                back_link = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[@action-type='clickBackToResult']"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", back_link)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", back_link)
                print(f"[{profile_index}] Returned to results.")
                time.sleep(2)

            except Exception as e:
                print(f"[{profile_index}] Error during profile loop: {e}")
                break

        # Try clicking the "Next" button
        try:
            next_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH,
                    "//div[contains(@class, 'col-md-2') and contains(@class, 'hidden-xs')]//span[contains(text(), 'Next') and not(contains(@style, 'display: none'))]"
                ))
            )

            if next_button.is_displayed() and next_button.is_enabled():
                driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", next_button)
                print("\n➡️ Clicked Next page\n")
                time.sleep(3)
            else:
                print("\n⛔ No more pages to scrape.")
                break

        except:
            print("\n✅ Reached last page. Scraping complete.")
            break

finally:
    driver.quit()
