import re
import os
import gspread
import logging
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import Playwright, sync_playwright, expect, TimeoutError, Page
from dotenv import load_dotenv, find_dotenv
from bs4 import BeautifulSoup
import time

# Configure logging
logging.basicConfig(filename='mapping_log.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# locate .env
dotenv_path = find_dotenv()
if not dotenv_path:
    raise FileNotFoundError(".env file not found.")
load_dotenv(dotenv_path)

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('GoogleCloudCredentials.json', scope)
client = gspread.authorize(creds)

# Get the Google Sheet ### Update for new runs
sheet = client.open("Consistent_Google_Sheet_Source").worksheet("Source")

# Set URL for repository
repositories = {
    "Geriatrics": "https://example.com/document_sets/4886",
    "Radiology": "https://example.com/document_sets/4890",
    "Diagnostic": "https://example.com/document_sets/19560",
    "DX": "https://example.com/document_sets/19560",
    "Telemedicine": "https://example.com/document_sets/39733",
    "High": "https://example.com/document_sets/19561",
    "Palliative": "https://example.com/document_sets/39732",
    "Social": "https://example.com/document_sets/25437",
    "Trauma-Informed": "https://example.com/document_sets/34397",
    "Family": "https://example.com/document_sets/4893",
    "Internal": "https://example.com/document_sets/4897",
    "Neurology": "https://example.com/document_sets/39734",
    "Obstetrics": "https://example.com/document_sets/44406",
    "Gynecology": "https://example.com/document_sets/44406",
    "Pediatrics": "https://example.com/document_sets/4891",
    "Addiction": "https://example.com/document_sets/4898"
}

# Fetch all rows
data = sheet.get_all_records()

# Environment Variables for Organization
Org_UN = os.environ.get('Org_User_ID')
Org_PW = os.environ.get('Org_Password')

# Ensure Org_UN and Org_PW are obtained
if not Org_UN or not Org_PW:
    raise ValueError("Org_User_ID or Org_Password environment variables are not set.")

# Log for debugging values
logging.info(f"Org_UN={Org_UN}")
logging.info(f"Org_PW={'*' * len(Org_PW)}")  # Masks the password for privacy in logs


# Get manual backup if 2FA method fails
def get_manual_2fa_code():
    return input("Please enter the 2FA code: ")

# Method to scroll around on a page to load content as needed
def scroll_around(page):

    try:
        page.keyboard.press("PageDown")
        page.keyboard.press("PageDown")
        time.sleep(1)
        page.keyboard.press("PageDown")
        page.keyboard.press("PageDown")
        time.sleep(1)
        page.keyboard.press("PageDown")
        page.keyboard.press("PageDown")
        time.sleep(1)
        page.keyboard.press("PageDown")
        page.keyboard.press("PageDown")
        time.sleep(1)
        page.keyboard.press("PageDown")
        page.keyboard.press("PageDown")
        time.sleep(1)
        page.keyboard.press("PageDown")
        page.keyboard.press("PageDown")
        page.keyboard.press("Home")
    except Exception as e:
        logging.error(f"scroll_around method failed: {e}")

# Method to parse Google Sheet, then use it to find the right case
def find_and_select_case(page, case):

    try:
        # Navigate to document repository based on case name
        case_repository_key = case.strip().split()[0]
        if case_repository_key in repositories:
            page.goto(repositories[case_repository_key])
            page.wait_for_load_state("networkidle")
            logging.info(f"Navigated to document repository: {page.url}")

        # Check if we were redirected to the login page due to session expiry
        if "users/sign_in" in page.url:
            logging.error("Redirected to login page. Authentication might have failed.")
            raise Exception("Session expired, redirected to login page.")
        
        # Scroll down immediately to begin loading cases, speed up find method later
        page.keyboard.press("PageDown")
        page.keyboard.press("PageDown")
        page.keyboard.press("PageDown")
        time.sleep(1)
        page.keyboard.press("Home")

        # Select and click the case title with CSS selector and xpath backup
        try:
            element = page.locator(f"a:has-text('{case}')")
            element.wait_for(state="visible", timeout=5000)
            element.scroll_into_view_if_needed()
            if not element.is_visible:
                try:
                    element.scroll_into_view_if_needed()
                    if not element.is_visible:
                        scroll_around(page)
                        element.scroll_into_view_if_needed()
                except Exception as e:
                    logging.error(f"Cannot bring case into view: {e}")

            element.click()
            logging.info(f"Clicked on {case} using CSS Selector")
        except Exception as e:
            logging.error(f"Failed to click {case} with CSS Selector: {e}")
            try:
                element = page.locator(f"xpath=//td[@class='title']//a[b[contains(text(), '{case}')]]")
                element.scroll_into_view_if_needed()
                element.wait_for(state="visible")
                element.click()
                logging.info(f"Clicked on {case} using Xpath")
            except Exception as e:
                logging.error(f"Failed to click {case} using Xpath: {e}")
            

        page.wait_for_function("() => window.location.href.includes('/document_set_document_relations')", timeout=10000)
        page.wait_for_load_state("networkidle")
        return True


    except Exception as e:
        logging.error(f"Unable to locate and click {case}: {e}")
        return False

# Method to check Content Mapping changes went through to Case Map
def verify_learning_objective_in_table(page, learning_objective, teaching_point):
    
    try:
        # Target the specific table within the reasoning tool panel using Playwright locator
        tableLocator = page.locator(".reasoning-tool-panel .fixed-height-table table.pure-table.pure-table-striped:has(tbody td)") # Udpate: removed " span.search result" from end to allow more tr matches during verification

        # Ensure the table is found
        if not tableLocator:
            raise Exception("Target table not found with CSS Selector.")

        # Extract the table body rows directly with Playwright
        rows_locators = tableLocator.locator('tbody tr')

        # Iterate through rows to find the matching learning objective and teaching point
        for row_locator in rows_locators.element_handles():
            row_text = row_locator.inner_text()
            logging.info(f"Logging a verification table row's inner text for examination: {row_text}")
            short_learning_objective = learning_objective[:20]
            short_teaching_point = teaching_point[:20]
            if short_learning_objective in row_text and short_teaching_point in row_text:
                logging.info(f"Success: Case Content Map contains match for LO '{learning_objective}' and TP '{teaching_point}': {row_text}")
                return True
        
        raise Exception(f"Learning Objective '{learning_objective}' and Teaching Point '{teaching_point}' not found in table rows.")

    except Exception as e:
        logging.error(f"No match for LO & TP found in Case Map: {e}")
        return False

def get_2fa_code(context, Org_UN, Org_PW):
    try:
        page1 = context.new_page()

        # Sign into Google Mail for 2FA
        page1.goto("https://mail.google.com")
        
        
        #Check proper redirect to Google accounts sign-in
        try:
            page1.wait_for_function("() => window.location.href.includes('accounts.google.com')", timeout=20000)
            page1.wait_for_load_state("networkidle")

            # Enter credentials
            page1.get_by_label("Email or phone").fill(Org_UN)
            page1.get_by_role("button", name="Next").click()
        except Exception as e:
            logging.error(f"Failed wait_for_url method for gmail redirect: {e}")

            # Enter credentials
            page1.get_by_label("Email or phone").fill(Org_UN)
            page1.get_by_role("button", name="Next").click()
            

        time.sleep(2)
        logging.info(f"2FA redirect to: {page1.url}")
        # Allow Redirect
        try:
            page1.wait_for_function("() => window.location.href.includes('okta')", timeout=20000)
            logging.info(f"Arrived at Okta sign-in: {page1.url}")
            page1.wait_for_load_state("networkidle")
        except Exception as e:
            logging.error(f"Okta wait failed: {e}")
            # Manually wait for all redirects
            time.sleep(2)
            page1.wait_for_load_state("networkidle")

        # Enter Credentials into Okta
        page1.get_by_label("Username").fill(Org_UN)
        page1.get_by_label("Password").fill(Org_PW)

        #consistent Okta sign-in issue, likely sign in button:
        try:
            page1.locator("#form20 > div.o-form-button-bar > input").click()
            logging.info("Clicked Okta Sign in button using CSS Selector")

        except Exception as e:
            logging.error(f"Failed to click Okta Sign in using CSS Selector: {e}")
            try:
                page1.locator("""xpath=//*[@id="form20"]/div[2]/input""").click()
                logging.info("Clicked Okta Sign in button using Xpath")
            
            except Exception as e:
                logging.error(f"Failed to click Okta Sign in using Xpath: {e}")

        time.sleep(10)
        logging.info(f"2FA redirect to: {page1.url}")

        # Wait for final navigation to inbox for gmail
        try:
            page1.wait_for_function("() => window.location.href.includes('https://mail.google.com/mail/u/0/#inbox')", timeout=20000)
            logging.info("Redirected to gmail")
            time.sleep(5)
        except Exception as e:
            logging.error(f"Failed to wait for gmail page redirect: {e}")
            time.sleep(5)
        
        try:
            page1.wait_for_load_state("domcontentloaded", timeout=20000)
            logging.info(f"Arrived at inbox: {page1.url}")
        except Exception as e:
            time.sleep(5)
            logging.error(f"Wait for Inbox load failed: {e}")
            logging.info(f"Arrived at inbox: {page1.url}")
        
        refresh_count = 0
        while refresh_count < 5:
            page1.reload()
            time.sleep(3)
            page1.wait_for_load_state("domcontentloaded")

            try:
                latest_email_subject = page1.locator("span:has-text('Organization Code')").first
                logging.info(f"Locator results for email: {latest_email_subject}")
                
                try:
                    email_content = latest_email_subject.text_content(timeout=3000)
                    logging.info(f"Email content found: {email_content}")
                except Exception as e:
                    logging.error(f"Email content could not be extracted: {e}")
                    try:
                        email_content = latest_email_subject.textContent(timeout=3000)
                        logging.info(f"2nd attempt - Email content found: {email_content}")
                    except Exception as e:
                        logging.error(f"2nd attempt to extract text failed: {e}")
                
                # Scan for 6 digit code
                code_match = re.search(r'(?<!\d)\d{6}(?!\d)', email_content)
                logging.info(f"Code match found: {code_match}")

                # Extract 6 digit code if match found
                if code_match:
                    code = code_match.group()
                    logging.info(f"Code found: {code}")
                else:
                    logging.error("No 6-digit code found in email content")

                if not code:
                    logging.error("2FA code not found in the email content")
                page1.close()
                return code
            except Exception as e:
                logging.error(f"Retrying fetching email: {e}. Attempt {refresh_count + 1}")

            refresh_count += 1

        logging.error("Failed to retrieve 2FA code after retries")
        return None
    except Exception as e:
        logging.error(f"Error retrieving 2FA code: {e}")
        return None

# Method to run the actual Playwright edit loop
def run(playwright: Playwright) -> None:
    # Updated the locate_and_click method to include more rudimentary Playwright methods prior to falling back to xpath and css attempts
    def locate_and_click(page: Page, fallback_css: str, primary_xpath: str, description: str, element_type: str, retries: int = 3, wait_time: int = 1000):
        for attempt in range(retries):
            try:
                logging.info(f"Attempt {attempt + 1}: Trying to click '{description}' using direct Playwright methods.")
                page.wait_for_load_state("networkidle", timeout=wait_time)
                page.wait_for_load_state("domcontentloaded", timeout=wait_time)
            
                # Attempt using text
                try:
                    element = page.get_by_text(description)
                    element.wait_for(state="visible", timeout=wait_time)
                    element.scroll_into_view_if_needed(timeout=wait_time)
                    element.click(timeout=wait_time)
                    logging.info(f"Click successful using text for '{description}'")
                    return
                except Exception as e:
                    logging.debug(f"Text attempt failed for '{description}': {e}")
                
                # Attempt using title
                try:
                    element = page.get_by_title(description)
                    element.wait_for(state="visible", timeout=wait_time)
                    element.scroll_into_view_if_needed(timeout=wait_time)
                    element.click(timeout=wait_time)
                    logging.info(f"Click successful using title for '{description}'")
                    return
                except Exception as e:
                    logging.debug(f"Title attempt failed for '{description}': {e}")

                # Attempt using role
                try:
                    element = page.get_by_role(element_type, name=description)
                    element.wait_for(state="visible", timeout=wait_time)
                    element.scroll_into_view_if_needed(timeout=wait_time)
                    element.click(timeout=wait_time)
                    logging.info(f"Click successful using role for '{description}'")
                    return
                except Exception as e:
                    logging.debug(f"Role attempt failed for '{description}': {e}")

                # Attempt using has-text locator
                try:
                    element = page.locator(f"{element_type}:has-text('{description}')")
                    element.wait_for(state="visible", timeout=wait_time)
                    element.scroll_into_view_if_needed(timeout=wait_time)
                    element.click(timeout=wait_time)
                    logging.info(f"Click successful using has-text for '{description}'")
                    return
                except Exception as e:
                    logging.debug(f"Has-text locator attempt failed for '{description}': {e}")

                # If above methods fail, try using xpath
                logging.info(f"Attempt {attempt + 1}: Trying to click '{description}' using XPath '{primary_xpath}'")
                element = page.locator(f"xpath={primary_xpath}").first
                element.wait_for(state="visible", timeout=wait_time)
                element.scroll_into_view_if_needed(timeout=wait_time)
                if not element.is_enabled():
                    raise Exception(f"Element '{description}' is visible but not enabled.")
                element.click(timeout=wait_time)
                page.wait_for_load_state("networkidle", timeout=wait_time)
                logging.info(f"Click successful using XPath for '{description}'")
                return
            except Exception as e:
                logging.error(f"Attempt {attempt + 1} failed for '{description}' using XPath '{primary_xpath}': {e}")
                if attempt == retries - 1:
                    # Attempt using fallback CSS selector
                    try:
                        logging.info(f"Fallback Attempt: Clicking '{description}' using CSS '{fallback_css}'")
                        # Default to first located CSS match
                        element = page.locator(fallback_css).first
                        element.wait_for(state="visible", timeout=wait_time)
                        if not element.is_enabled():
                            raise Exception(f"Fallback element '{description}' is visible but not enabled.")
                        element.scroll_into_view_if_needed()
                        element.click(timeout=wait_time)
                        page.wait_for_load_state("networkidle", timeout=wait_time)
                        logging.info(f"Click successful using fallback CSS for '{description}'")
                        return
                    except Exception as e:
                        logging.error(f"Fallback CSS attempt failed for '{description}': {e}")
                        raise
            # Added delay between retries
            time.sleep(1)

    try:
        browser = playwright.chromium.launch(headless=True)
        # Increased viewport height to stop visibility issues with buttons (Webpage Save and Add Row errors in CM Editor dropdown) - if causing load issues, resize
        context = browser.new_context(viewport={'width': 1920, 'height': 2200, 'device_scale_factor': 1}) # increase context window, maintain pixel scale

        page = context.new_page()

        # Sign into Organization
        page.goto("https://example.com/users/sign_in")
        page.wait_for_load_state("networkidle")
        logging.info(f"Accessed Organization page: {page.url}")
        page.get_by_label("Email").fill(Org_UN)
        page.get_by_label("Password").fill(Org_PW)
        page.keyboard.press("Enter")

        # Fetch the 2FA code from Google Mail
        code = get_2fa_code(context, Org_UN, Org_PW)
        logging.info(f"2FA code retrieved from email: {code}")
        if not code:
            logging.info("Fallback to manual entry for 2FA code")
            code = get_manual_2fa_code()

        logging.info(f"2FA code received from user: {code}")
        if not code:
            raise Exception("2FA code is not retrieved properly")

        # Input 2FA code
        page.wait_for_load_state("domcontentloaded", timeout=10000)
        logging.info(f"Arrived at 2FA page: {page.url}")
        page.get_by_label("Please enter the time-").fill(code)
        
        # Hit Submit to enter Organization Learning Management System
        try:
            locate_and_click(page, 
                         "#sign_in_form > div > div > div > input[type=submit]", 
                         """//*[@id="sign_in_form"]/div/div/div/input""", 
                         "Submit",
                         "input")
        except Exception:
            if page.url == "https://example.com":
                logging.info("Navigation to Organization already successful")
            else:
                SystemExit

        # Iterate through rows in Google Sheet
        for row in data:
            try:
                case = row["Case"]
                learning_objective = row["Learning Objective"]
                teaching_point = row["Teaching Point"]

                logging.info(f"Processing: Case={case}, Learning Objective={learning_objective}, Teaching Point={teaching_point}")

                # Navigate to appropriate case
                if find_and_select_case(page, case):
                    page.wait_for_load_state("domcontentloaded", timeout=5000)
                    page.wait_for_load_state("networkidle", timeout=5000)
                    logging.info(f"Arrived at {case} page: {page.url}")

                    # Hit "Editor" to enter the edit page
                    page.wait_for_load_state("networkidle", timeout=10000)
                    page.wait_for_load_state("domcontentloaded", timeout=10000)
                    locate_and_click(page, 
                                     ".panel a.button[href*='/edit']:has-text('Editor')", 
                                     """//*[@class='panel']//a[contains(@href, '/edit') and contains(text(), 'Editor')]""", 
                                     "Editor",
                                     "link")
                    # Wait for Editor to load
                    page.wait_for_function("() => window.location.href.includes('/versions')", timeout=20000)
                    page.wait_for_load_state("domcontentloaded", timeout=20000)
                    page.wait_for_load_state("networkidle", timeout=20000)
                    logging.info(f"Entered editor for {case}: {page.url}")

                    # Retry if editor not successfully entered
                    if '/versions' not in page.url:
                        logging.warning(f"Failed to enter Editor for {case}, trying fallback to CSS Selector / XPath.")
                        try:
                            element = page.locator(".panel a.button[href*='/edit']:has-text('Editor')")
                            element.wait_for(state="visible", timeout=20000)
                            element.click(timeout=20000)
                            page.wait_for_function("() => window.location.href.includes('/versions')", timeout=20000)
                            page.wait_for_load_state("networkidle", timeout=20000)
                            logging.info(f"Clicked Editor using backup CSS Selector: {page.url}")

                        except Exception as e:
                            logging.error(f"Selector backup failed: {e}")
                            try:
                                element = page.locator("""xpath=//*[@class='panel']//a[contains(@href, '/edit') and contains(text(), 'Editor')]""")
                                element.wait_for(state="visible", timeout=20000)
                                element.click(timeout=20000)
                                page.wait_for_function("() => window.location.href.includes('/versions')", timeout=20000)
                                page.wait_for_load_state("networkidle", timeout=20000)
                                logging.info(f"Clicked Editor using backup Xpath: {page.url}")

                            except Exception as e:
                                logging.error(f"Xpath backup failed: {e}")
                                SystemExit
                    

                    # Wait, then turn on Autosave feature so we don't need to click "save" at the end
                    # Due to a dev fix, autosave now correctly pops on automatically, this segment is no longer needed
                    """
                    try:
                        page.get_by_role("link", name="Auto  Off").click()
                        logging.info("Autosave initiatied")
                    except Exception as a:
                        logging.info(f"Autosave initiation failed: {a}")
                    """

                    # Locate Learning Objective in Web Editor page

                    # Refined Attempt: Using Locator before BeautifulSoup to find correct LO class, click LO
                    try:
                        page.wait_for_load_state("networkidle")

                        # attempt with full learning objective
                        learning_objective_location = page.locator('div.learning-objective-content').filter(has_text=f"{learning_objective}")
                        logging.info(f"Locator query for Learning Objective: {learning_objective_location}")
                        
                        # Resolve any duplicate matches to click on the first match
                        learning_objective_location.first.wait_for(state="visible", timeout=5000)
                        learning_objective_location.first.scroll_into_view_if_needed
                        learning_objective_location.first.click()
                        logging.info(f"Learning Objective located in full directly: {learning_objective_location}")
                    except Exception as e:
                        logging.info(f"Learning Objective not located or clicked directly, fallback start. Error message: {e}; Locator info: {learning_objective_location}")
                        try:
                            short_learning_objective = learning_objective[:40]
                            logging.info(f"Shortened Learning Objective to reduce errors, shortened version: {short_learning_objective}")
                            learning_objective_location = page.locator('div.learning-objective-content').filter(has_text=f"{short_learning_objective}")
                            logging.info(f"Locator query for Learning Objective: {learning_objective_location}")
                            
                            # Resolve any duplicate matches by clicking the first match
                            learning_objective_location.first.wait_for(state="visible", timeout=5000)
                            learning_objective_location.first.scroll_into_view_if_needed
                            learning_objective_location.first.click()
                            logging.info(f"Learning Objective located in full directly: {learning_objective_location}")
                        except Exception as e:
                            logging.info(f"Learning Objective not found as shortened version, begin BeautifulSoup fallback. Error message: {e}")
                            try:
                                html = page.content()
                                soup = BeautifulSoup(html, "html.parser")
            
                                # Identify all potential elements
                                elements = soup.find_all("span", string=re.compile(re.escape(learning_objective)))

                                if elements:
                                    # Filter elements within learning-objective-content
                                    potential_matches = [el for el in elements if 'learning-objective-content' in el.parent.get('class', [])]
                
                                    if potential_matches:
                                        logging.info(f"Filtered potential matches: {[el.get_text() for el in potential_matches]}")

                                        # Log elements found within the target section.
                                        element = potential_matches[0]  # Use the first match or implement additional logic if needed
                                        logging.info(f"Located first match: {element}")
                                        element_xpath = f"//div[contains(@class, 'learning-objective-content')]//*[normalize-space(text())='{element.get_text()}']"
                                        logging.info(f"Located LO element xpath: {element_xpath}")
                                        page.locator(f"xpath={element_xpath}").scroll_into_view_if_needed()
                                        page.locator(f"xpath={element_xpath}").click()
                                        logging.info("LO search attempt with BeautifulSoup successful.")
                                    else:
                                        raise Exception("No matches found in learning-objective-content.")
                                else:
                                    raise Exception("Element not found using BeautifulSoup.")
                            
                            except Exception as soup_fallback_e:
                                logging.error(f"Learning Objective locator with BeautifulSoup failed: {soup_fallback_e}")

                    page.wait_for_load_state("networkidle")
                    

                    # Confirm "I'm sure, let's do this" to enter Content Mapping editor
                    locate_and_click(page, 
                                     """.gen-modal .aq-button-bar.bottom-right button:has-text("I'm sure, let's do this")""", 
                                     """//*[@class='gen-modal']//div[contains(@class, 'aq-button-bar') and contains(@class, 'bottom-right')]//button[contains(text(), "I'm sure, let's do this")]""", 
                                     "I'm sure, let's do this",
                                     "button.aq-button-2")

                    # Navigate into Content Mapping Tool
                    time.sleep(1)
                    try:
                        page.get_by_role("heading", name=" Content Mapping Tool").locator("span").first.click()
                        logging.info("Successfully opened Content Mapping Tool")
                    except Exception as a:
                        logging.error(f"Could not open Content Mapping Tool: {a}")

                    # Add Row
                    time.sleep(1)
                    try:
                        page.get_by_role("button", name="Add Row").click(timeout=3000)
                        logging.info("Successfully clicked Add Row")
                    except Exception as a:
                        logging.error(f"Could not click Add Row conventionally, attempting force: {a}")
                        try:
                            page.get_by_role("button", name="Add Row").first.click(force=True, timeout=5000)
                        except Exception as a:
                            logging.error(f"Could not click CME Add Row using force: {a}")

                    # Try to make TP selection from dropdown
                    time.sleep(1)
                    try:
                        # select combobox with full LO text
                        combobox = page.get_by_role("row", name=learning_objective).get_by_role("combobox").first
                        combobox.wait_for(state="visible", timeout=5000)
                        if combobox.is_visible():
                            combobox.click()
                            tp = page.locator("option", has_text=teaching_point)
                            logging.info(f"Full LO, Full TP:  Locator results for TP matches in dropdown: {tp}")
                            # select the 'hidden' TP with select_option
                            option_value = tp.first.get_attribute("value")
                            combobox.select_option(value=option_value)
                            logging.info("Successfully selected TP in dropdown")
                    except Exception as e:
                        logging.error(f"Could not locate combobox using full LO text: {e}")
                        try: # Locate the correct combobox using the short LO
                            short_learning_objective = learning_objective[:25]
                            short_teaching_point = teaching_point[:20]
                            combobox = None
                            # Locate combobox with shortened text
                            combobox = page.get_by_role("row").filter(has_text=short_learning_objective).get_by_role("combobox").first
                            combobox.wait_for(state="visible", timeout=5000)
                            if combobox.is_visible():
                                combobox.click()
                                tp = page.locator("option", has_text=short_teaching_point)
                                logging.info(f"Short LO, Short TP: Locator results for TP matches in dropdown: {tp}")
                                # select the 'hidden' TP with select_option
                                option_value = tp.first.get_attribute("value")
                                combobox.select_option(value=option_value)
                                logging.info("Successfully selected TP in dropdown")
                        except Exception as e:
                            logging.error(f"Combobox not located with shortened LO: {e}")
                            raise SystemExit

                    
                    # Click Save in CME
                    time.sleep(1)
                    try:
                        locate_and_click(page, 
                                     "button.aq-button:has-text('Save')", 
                                     """//*[@class='aq-button' and contains(text(), 'Save')]""", 
                                     "Save",
                                     "button")
                    except Exception as e:
                        logging.info(f"Failed to click Save in Content Mapping Editor conventionally, attempting force fallback: {e}")
                        try:
                            page.locator("button.aq-button:has-text('Save')").first.click(force=True, timeout=5000)
                        except Exception as e:
                            logging.info(f"Failed to click CME Save using force: {e}")

                    # Click Continue - Try to force the click if it's hidden
                    time.sleep(1)
                    try:
                        page.click("button.aq-button[style='margin-right: 5px;']:has-text('Continue')", force=True)
                        logging.info("Clicked Continue button with force=True")
                    except Exception as d:
                        logging.info(f"Failed to click Continue button with force=True: {d}")
                        try:
                            page.evaluate("""
                                (locator) => {
                                    const element = document.querySelector(locator);
                                    element.click();
                                }
                            """, "button.aq-button[style='margin-right: 5px;']:has-text('Continue')")
                            logging.info("Clicked Continue with page.evaluate")

                        except Exception as c:
                            logging.info(f"Failed to click Continue with page.evaluate: {c}")
                            try:
                                page.locator("xpath=//button[contains(@class, 'aq-button') and contains(@style, 'margin-right') and text()='Continue']").click(force=True)
                                logging.info("Clicked Continue with Xpath Force")
                            except Exception as e:
                                logging.error(f"Failed to click Continue with direct methods: {e}")
                                locate_and_click(page, 
                                     """button.aq-button[style='margin-right: 5px;']:has-text('Continue')""", 
                                     """//button[contains(@class, 'aq-button') and contains(@style, 'margin-right') and text()='Continue']""", 
                                     "Continue",
                                     "button")

                    # Try to click "save" prior to publish, also acts as a sleep for Autosave trigger
                    try:
                        page.locator(".edit-bar-control-bar > .gen-button.highlighted.small > a.button-name > i.fa.fa-save").first.click(timeout=1000)
                        logging.info("Clicked Save using CSS Selector")
                    except Exception as a:
                        logging.info(f"Unable to hit save with CSS Selector: {a}")
                        try:
                            page.locator("xpath=//div[@class='edit-bar-control-bar']//div[contains(@class, 'gen-button') and contains(@class, 'highlighted') and contains(@class, 'small')]//a[@class='button-name'][i[contains(@class, 'fa') and contains(@class, 'fa-save')").first.click(timeout=1000)
                            logging.info("Clicked Save using xpath")
                        except Exception as b:
                            logging.info(f"Unable to click save using xpath: {b}")

                    time.sleep(1)
                    # Wait for Publish button to be enabled by save / autosave triggers
                    try:
                        publish_button = page.locator(".edit-bar-control-bar > .gen-button.highlighted.small > a:has-text('Publish')")
                        publish_button.wait_for(state="visible", timeout=30000)
                        logging.info("Enabled Publish button visible, attempting click")
                    except Exception as e:
                        logging.error(f"Issue waiting for main Publish button visibility: {e}")
                        time.sleep(10)
                        
                        locate_and_click(page, 
                                         "a[href]:has-text('Publish')", 
                                         "//*[@class='doc-inside-wrapper']//a[contains(text(), 'Publish')]", 
                                         "Publish",
                                         "link")                    

                    # Confirm Publish, sleep for dialog boxes to pop
                    time.sleep(1)
                    try:
                        page.get_by_role("button", name="Publish").click()
                    except Exception:
                        logging.info("Basic Playwright method for Publish dialog box failed.")
                        locate_and_click(page, 
                                     "#modal-wrapper-wvppmqqhjw > div.modal-dialog > div.modal-buttons > button:nth-child(1)", 
                                     """//*[@id='modal-wrapper-urqdwk5v9nd']/div[2]/div[3]/button[1]""", 
                                     "Publish",
                                     "button")

                    # Click Cancel on updating Banner, sleep for dialog boxes to pop
                    time.sleep(1)
                    try:
                        page.get_by_role("button", name="Cancel").click()
                    except Exception:
                        logging.info("Basic Playwright method for Cancel banner update failed.")
                        locate_and_click(page, 
                                     "button.modal-button.gen-button.highlighted.small", 
                                     """//*[@id='modal-wrapper-c7c78lnq75u']/div[2]/div[3]/button[2]""", 
                                     "Cancel",
                                     "button")

                    # Wait for page reload dialog, accept and reload
                    page.once("dialog", lambda dialog: dialog.accept())

                    # Wait for page to reload, then check for the change
                    time.sleep(5)

                    # If dialog box on publishing still exists, use keyboard escape to remove it
                    # page.keyboard.press("Escape")
                    page.wait_for_load_state("networkidle", timeout=20000)
                    # Attempt to verify that we have made Content Mapping changes successfully
                    try:
                        # Click into Case Map
                        page.get_by_role("link", name="CASE MAP").click()
                        page.locator("input[name=\"learning_objective\"]").click()

                        # Enter LO in the filter input
                        page.locator("input[name=\"learning_objective\"]").fill(learning_objective)
                        page.keyboard.press("Enter")
                        time.sleep(2)

                        # Verify the table contains the entered learning objective and teaching point
                        logging.info(f"Attempting verification of mapping changes")
                        if verify_learning_objective_in_table(page, learning_objective, teaching_point):
                            logging.info("Verification successful.")
                        else:
                            logging.error("Verification failed.")
                            context.close()
                            browser.close()
                            raise SystemExit
    
                    except Exception as e:
                        logging.error(f"Error during verification steps: {e}")
                        context.close()
                        browser.close()
                        raise SystemExit


                    # Close the Sidebar Nav by clicking Case Map
                    page.get_by_role("link", name="CASE MAP").click()
                    
                    # Ensure that sidebar closes
                    time.sleep(1)

                    # Click "Push to OM"
                    page.once("dialog", lambda dialog: dialog.accept())
                    page.get_by_role("button", name="Push To OM").click()
                    logging.info("Pushed to OM")
                    time.sleep(5)

                    # Wait for redirect to Projects
                    # page.wait_for_url("INSERT PROJECTS URL", timeout=30000)

                    # Log success
                    logging.info(f"Successfully updated: Case={case}, Learning Objective={learning_objective}, Teaching Point={teaching_point}")

                    # Return to main page to restart loop
                    page.goto("https://example.com")

                else:
                    logging.error(f"Failed to locate and select case: {case}")
                    context.close()
                    browser.close()
                    raise SystemExit

            except Exception as e:
                logging.error(f"Error processing row: {row}, Error: {e}")
                context.close()
                browser.close()
                raise SystemExit

        context.close()
        browser.close()

    except Exception as e:
        logging.error(f"An error occurred: {e}")

with sync_playwright() as playwright:
    run(playwright)
