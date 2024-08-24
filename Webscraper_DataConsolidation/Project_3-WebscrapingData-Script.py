import re
import pickle
import os
import gspread
import logging
from oauth2client.service_account import ServiceAccountCredentials
from playwright.async_api import async_playwright, TimeoutError, Page
from dotenv import load_dotenv, find_dotenv
from bs4 import BeautifulSoup
import time
import asyncio  # Handle asynchronous tasks
import aiohttp  # Help asyncio with HTTP requests

# Setup Pickle Save and Load States
CHECKPOINT_FILE = 'dashboard-chkpt.pkl'

def save_state(state):
    with open(CHECKPOINT_FILE, 'wb') as f:
        pickle.dump(state, f)

def load_state():
    if os.path.exists(CHECKPOINT_FILE):
        with open(CHECKPOINT_FILE, 'rb') as f:
            return pickle.load(f)
    return None

# Configure logging
logging.basicConfig(filename='dashboard_log.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# locate .env
dotenv_path = find_dotenv()
if not dotenv_path:
    raise FileNotFoundError(".env file not found.")
load_dotenv(dotenv_path)

# Environment Variables for Organization
Org_UN = os.environ.get('Org_UN')
Org_PW = os.environ.get('Org_PW')

# Ensure Org_UN and Org_PW are obtained
if not Org_UN or not Org_PW:
    raise ValueError("Org_UN or Org_PW environment variables are not set.")

# Log for debugging values
logging.info(f"Org_UN={Org_UN}")
logging.info(f"Org_PW={'*' * len(Org_PW)}")  # Masks the password for privacy

# Create a dictionary to store course names and urls for case listings by course
courses = {
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

class GoogleSheetHandler:
    def __init__(self, spreadsheet_id, credentials):
        self.spreadsheet_id = spreadsheet_id
        self.credentials = credentials
        self.client = self.authenticate()
        self.sheet = self.client.open("Curriculum_Dashboard").worksheet("All_Data")

    def authenticate(self):
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(self.credentials, scope)
        return gspread.authorize(creds)

    def read_all_records(self):
        return self.sheet.get_all_records()

    def extract_course_data(self):
        data = self.read_all_records()
        course_data = []
        for row in data:
            if "Course" in row and "Case" in row:  # Ensure both fields are present
                course_name_full = row["Course"].strip()
                course_name_first_word = course_name_full.split()[0] if course_name_full else ""
                
                course_data.append({
                    "Course": course_name_first_word,
                    "Case Name": row["Case"].strip(),
                    "Teaching Point": row["Teaching Point"].strip()
                })
        return course_data

    def write_column(self, column: str, data: list):

        try:

            # Get all records to find the matching rows
            all_records = self.sheet.get_all_records()
        
            # Header index (0-based in Python, 1-based in Sheets)
            col_index = gspread.utils.a1_range_to_grid_range(f'{column}1')['startColumnIndex']
        
            for entry in data:
                if column == "EK":  # For Case Synopsis
                    match_key = "Case Name"
                    new_value = entry.get("Case Synopsis", "")
                elif column == "EL":  # For Full Text TP
                    match_key = "Teaching Point"
                    new_value = entry.get("Full Text", "")
                else:
                    raise ValueError("Invalid column specified")
            
                match_value = entry[match_key]
            
                # Locate the matching row in all_records
                for i, row in enumerate(all_records):
                    if row.get(match_key) == match_value:
                        # Update the correct cell based on row and col_index
                        cell = self.sheet.cell(i + 2, col_index + 1)
                        cell.value = new_value
                        self.sheet.update_cell(i + 2, col_index + 1, cell.value)
            
            logging.info(f"Wrote data to Google Sheet column: {column}")

        except Exception as e:
            logging.error(f"Error in GoogleSheetHandler - write_column: {e}")


class WebScraper:
    def __init__(self, base_url):
        self.base_url = base_url
        self.browser = None
        self.context = None
        self.page = None

    async def setup_browser(self):
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(headless=True)
        self.context = await self.browser.new_context(viewport={'width': 1920, 'height': 2200, 'device_scale_factor': 1})
        self.page = await self.context.new_page()

    async def close_browser(self):
        await self.context.close()
        await self.browser.close()
        await self.playwright.stop()

    async def attempt_login_page(self):
        await self.page.goto("https://example.com/users/sign_in")
        await self.page.wait_for_load_state("networkidle")
        logging.info(f"Accessed Org page: {self.page.url}")
        await self.page.get_by_label("Email").fill(Org_UN)
        await self.page.get_by_label("Password").fill(Org_PW)
        await self.page.keyboard.press("Enter")

    async def navigate_to_repository(self, url):
        await self.page.goto(url)
        await self.page.wait_for_load_state("networkidle")

    async def clean_html(self, html_content):
        soup = BeautifulSoup(html_content, 'html.parser')
        return soup.get_text(separator="\n", strip=True)  # Extract text and separate paragraphs with new lines
    
    async def clean_teaching_point(self, text_content):
        # Add a colon and space after TP header
        text_content = re.sub(r"(TEACHING POINT)([A-Z])", r"\1: \2", text_content)
        # Add newlines between each section and each paragraph
        text_content = re.sub(r"([a-z])([A-Z])", r"\1\n\2", text_content)
        text_content = re.sub(r"(\.)([A-Z])", r"\1\n\2", text_content)
        text_content = re.sub(r"(\?)([A-Z])", r"\1\n\2", text_content)
        text_content = re.sub(r"(\:)([a-zA-Z])", r"\1\n\2", text_content)
        # Add newline between numbered or lettered list items
        text_content = re.sub(r"([^0-9])(\d+\.)", r"\1\n\2", text_content)
        text_content = re.sub(r"([^a-zA-Z\s])([a-zA-Z]\.)", r"\1\n\2", text_content)
        # Add newline between concatenated parentheticals and next sentence
        text_content = re.sub(r"(\))([a-zA-Z])", r"\1\n\2", text_content)
        return text_content


    async def locate_and_click(self, page: Page, fallback_css: str, primary_xpath: str, description: str, element_type: str, retries: int = 3, wait_time: int = 1000):
        for attempt in range(retries):
            try:
                logging.info(f"Attempt {attempt + 1}: Trying to click '{description}' using direct Playwright methods.")
                await page.wait_for_load_state("networkidle", timeout=wait_time)
                await page.wait_for_load_state("domcontentloaded", timeout=wait_time)

                # Attempt using text
                try:
                    element = page.get_by_text(description)
                    await element.wait_for(state="visible", timeout=wait_time)
                    await element.scroll_into_view_if_needed(timeout=wait_time)
                    await element.click(timeout=wait_time)
                    logging.info(f"Click successful using text for '{description}'")
                    return
                except Exception as e:
                    logging.debug(f"Text attempt failed for '{description}': {e}")

                # Attempt using title
                try:
                    element = page.get_by_title(description)
                    await element.wait_for(state="visible", timeout=wait_time)
                    await element.scroll_into_view_if_needed(timeout=wait_time)
                    await element.click(timeout=wait_time)
                    logging.info(f"Click successful using title for '{description}'")
                    return
                except Exception as e:
                    logging.debug(f"Title attempt failed for '{description}': {e}")

                # Attempt using role
                try:
                    element = page.get_by_role(element_type, name=description)
                    await element.wait_for(state="visible", timeout=wait_time)
                    await element.scroll_into_view_if_needed(timeout=wait_time)
                    await element.click(timeout=wait_time)
                    logging.info(f"Click successful using role for '{description}'")
                    return
                except Exception as e:
                    logging.debug(f"Role attempt failed for '{description}': {e}")

                # Attempt using has-text locator
                try:
                    element = page.locator(f"{element_type}:has-text('{description}')")
                    await element.wait_for(state="visible", timeout=wait_time)
                    await element.scroll_into_view_if_needed(timeout=wait_time)
                    await element.click(timeout=wait_time)
                    logging.info(f"Click successful using has-text for '{description}'")
                    return
                except Exception as e:
                    logging.debug(f"Has-text locator attempt failed for '{description}': {e}")

                # If above methods fail, try using xpath
                logging.info(f"Attempt {attempt + 1}: Trying to click '{description}' using XPath '{primary_xpath}'")
                element = page.locator(f"xpath={primary_xpath}").first
                await element.wait_for(state="visible", timeout=wait_time)
                await element.scroll_into_view_if_needed(timeout=wait_time)
                if not element.is_enabled():
                    raise Exception(f"Element '{description}' is visible but not enabled.")
                await element.click(timeout=wait_time)
                await page.wait_for_load_state("networkidle", timeout=wait_time)
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
                        await element.wait_for(state="visible", timeout=wait_time)
                        if not element.is_enabled():
                            raise Exception(f"Fallback element '{description}' is visible but not enabled.")
                        await element.scroll_into_view_if_needed()
                        await element.click(timeout=wait_time)
                        await page.wait_for_load_state("networkidle", timeout=wait_time)
                        logging.info(f"Click successful using fallback CSS for '{description}'")
                        return
                    except Exception as e:
                        logging.error(f"Fallback CSS attempt failed for '{description}': {e}")
                        raise
            # Added delay between retries
            await asyncio.sleep(1)

    async def scroll_around(self):
        await self.page.keyboard.press('PageDown')
        await self.page.keyboard.press('PageDown')
        await self.page.keyboard.press('PageDown')
        await asyncio.sleep(2)
        await self.page.keyboard.press('PageDown')
        await self.page.keyboard.press('PageDown')
        await self.page.keyboard.press('PageDown')
        await asyncio.sleep(2)
        await self.page.keyboard.press('PageDown')
        await self.page.keyboard.press('PageDown')
        await self.page.keyboard.press('PageDown')
        await asyncio.sleep(2)
        return None

    async def scrape_case_synopsis(self, case):
        try:

            try:
                # Find the modal containing the correct case name by navigating to the synopsis element
                element = self.page.locator(f"a[data-content-library-modal-title-param='{case}']").first
                # Extract all the inner text from this element
                synopsis_html = await element.get_attribute("data-content-library-modal-synopsis-param")
            except Exception as e:
                logging.error(f"Could not find synopsis with locator, attempting page scroll to load elements")
                try:
                    await self.scroll_around()
                    element = self.page.locator(f"a[data-content-library-modal-title-param='{case}']").first
                    # Extract all the inner text from this element
                    synopsis_html = await element.get_attribute("data-content-library-modal-synopsis-param")
                except Exception as e:
                    logging.error(f"Failed to scroll around and locate synopsis: {e}")

            synopsis_text = await self.clean_html(synopsis_html)

            return synopsis_text
        except Exception as e:
            logging.error(f"Error in WebScraper - scrape_case_synopsis: {e}")
    
    async def scrape_teaching_point(self, course_url, case_name, teaching_point_name):
        page2 = await self.context.new_page()
        await page2.goto(course_url)
        await page2.wait_for_load_state("domcontentloaded")

        try:
            logging.info(f"TP Scraper: Row has TP, seeking match for case: <<{case_name}>> and TP: <<{teaching_point_name}>>")
            
            try:
                # Navigate to the specific case page
                await page2.locator(f'a:has-text("{case_name}")').first.click()
            except Exception as e:
                logging.error(f"Failed to click into case, attempting scroll around to load case")
                try:
                   await self.scroll_around()
                   await page2.locator(f'a:has-text("{case_name}")').first.click()
                except Exception as e:
                    logging.error(f"Still failed to click into case after scroll around: {e}")


            await page2.wait_for_function("() => window.location.href.includes('/document_set_document_relations')", timeout=10000)
            await page2.wait_for_load_state("networkidle")
            logging.info(f"Entered case: {case_name} at url: {page2.url}")

            # Change the viewing mode to 'full'
            selector = page2.locator('select.doc-controls-select.doc-controls-view-mode').first
            await selector.select_option("full")

            # Attempt to locate TP within the ancestor 'doc-section'
            # Syntax explanation:  ">" denotes a direct child, while a space between elements denotes any descendant. Navigation up the DOM structure is ("..")
            try:
                # Find the teaching point element
                teaching_point_element = page2.locator(f"div.doc-section > div.doc-section-header > div.doc-section-header-banner > h1:has-text('{teaching_point_name}')").first
                el_1text = await teaching_point_element.text_content()
                logging.info(f"First TP locator element text should match TP: {el_1text}")
                # teaching_point_element = page2.locator(f"div.doc-section-header", has=page2.locator("div.teaching-point-topper")) 
                # nested_element = teaching_point_element.locator(f"div.doc-section-header-banner > h1:has-text('{teaching_point_name}')").first
                await teaching_point_element.scroll_into_view_if_needed(timeout=3000)
                await teaching_point_element.wait_for(state="visible", timeout=3000)
                # navigate using xpath ancestor navigation to get to the first doc-section container holding our h1 text
                # search via xpath ancestor navigated for full webpage div container: locator("xpath=ancestor::div[contains(@class, 'doc-section')]").first
                target_section = teaching_point_element.locator("..").locator("..").locator("..").locator("div.doc-section-body > div.doc-children").first
                # Get to the doc-children in the body of the doc-section element
                # target_section = parent_section.locator("div.doc-section-body > div.doc-children").first
                # el_3text = await target_section.text_content()
                # logging.info(f"Third TP locator element (TP Body) text: {el_3text}")
            except Exception as e:
                logging.error(f"Unable to locate & scrape full-text TP: {e}")
                await page2.close()
                return None  # Return None if teaching point not found            

            # Get all text
            text_content = await target_section.text_content()
            logging.info(f"Text content of TP parent element: {text_content}")

            # Clean text with regular expressions (as needed)
            full_text = await self.clean_teaching_point(text_content)
            logging.info(f"Scraped & cleaned Full TP for TP <<{teaching_point_name}>>")
            return full_text

        except Exception as e:
            logging.error(f"Error in TP Scraper: {e}")
        finally:
            await page2.close()

    async def get_2fa_code(self, org_un, org_pw):
        try:
            page1 = await self.context.new_page()

            # Sign into Google Mail for 2FA
            await page1.goto("https://mail.google.com")

            # Check proper redirect to Google accounts sign-in
            try:
                await page1.wait_for_function("() => window.location.href.includes('accounts.google.com')", timeout=20000)
                await page1.wait_for_load_state("networkidle")

                # Enter credentials
                await page1.get_by_label("Email or phone").fill(org_un)
                await page1.get_by_role("button", name="Next").click()
            except Exception as e:
                logging.error(f"Failed wait_for_url method for gmail redirect: {e}")

                # Enter credentials
                await page1.get_by_label("Email or phone").fill(org_un)
                await page1.get_by_role("button", name="Next").click()

            time.sleep(2)
            logging.info(f"2FA redirect to: {page1.url}")
            # Allow Okta Redirect
            try:
                await page1.wait_for_function("() => window.location.href.includes('okta')", timeout=20000)
                logging.info(f"Arrived at Okta sign-in: {page1.url}")
                await page1.wait_for_load_state("networkidle")
            except Exception as e:
                logging.error(f"Okta wait failed: {e}")
                # Manually wait for all redirects
                time.sleep(2)
                await page1.wait_for_load_state("networkidle")

            # Enter Credentials into Okta
            await page1.get_by_label("Username").fill(org_un)
            await page1.get_by_label("Password").fill(org_pw)

            # consistent Okta sign-in issue, likely sign in button:
            try:
                await page1.locator("#form20 > div.o-form-button-bar > input").click()
                logging.info("Clicked Okta Sign in button using CSS Selector")

            except Exception as e:
                logging.error(f"Failed to click Okta Sign in using CSS Selector: {e}")
                try:
                    await page1.locator("""xpath=//*[@id="form20"]/div[2]/input""").click()
                    logging.info("Clicked Okta Sign in button using XPath")

                except Exception as e:
                    logging.error(f"Failed to click Okta Sign in using XPath: {e}")

            time.sleep(10)
            logging.info(f"2FA redirect to: {page1.url}")

            # Wait for final navigation to inbox for gmail
            try:
                await page1.wait_for_function("() => window.location.href.includes('https://mail.google.com/mail/u/0/#inbox')", timeout=20000)
                logging.info("Redirected to gmail")
                time.sleep(5)
            except Exception as e:
                logging.error(f"Failed to wait for gmail page redirect: {e}")
                time.sleep(5)

            try:
                await page1.wait_for_load_state("domcontentloaded", timeout=20000)
                logging.info(f"Arrived at inbox: {page1.url}")
            except Exception as e:
                time.sleep(5)
                logging.error(f"Wait for Inbox load failed: {e}")
                logging.info(f"Arrived at inbox: {page1.url}")

            refresh_count = 0
            while refresh_count < 5:
                await page1.reload()
                time.sleep(3)
                await page1.wait_for_load_state("domcontentloaded")

                try:
                    latest_email_subject = page1.locator("span:has-text('Org Code')").first
                    logging.info(f"Locator results for email: {latest_email_subject}")
                    # latest_email_subject.wait_for(state="visible", timeout=10000)
                    # latest_email_subject.scroll_into_view_if_needed()

                    # Extract the text content
                    try:
                        email_content = await latest_email_subject.text_content(timeout=3000)
                        logging.info(f"Email content found: {email_content}")
                    except Exception as e:
                        logging.error(f"Email content could not be extracted: {e}")
                        try:
                            email_content = await latest_email_subject.textContent(timeout=3000)
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
                    logging.info(f"Returning code {code} and closing email page")
                    await page1.close()
                    return code
                except Exception as e:
                    logging.error(f"Retrying fetching email: {e}. Attempt {refresh_count + 1}")

                refresh_count += 1

            logging.error("Failed to retrieve 2FA code after retries")
            return None
        except Exception as e:
            logging.error(f"Error retrieving 2FA code: {e}")
            return None


class Coordinator:
    def __init__(self, sheet_handler, scraper):
        self.sheet_handler = sheet_handler
        self.scraper = scraper
        self.case_synopsis_cache = {}  # Add caching for case synopsis
        self.teaching_point_cache = {}  # Add caching for teaching points
        self.processed_cases = set()

    def get_course_url(self, course_name):
        stripped_course_name = course_name.strip()
        if stripped_course_name in courses:
            return courses[stripped_course_name]
        else:
            raise ValueError(f"No URL found for course: {stripped_course_name}")
        
    def create_state(self, course_data, case_synopsis_data, teaching_point_data):
        state = {
            'course_data': course_data,
            'case_synopsis_data': case_synopsis_data,
            'teaching_point_data': teaching_point_data,
            'case_synopsis_cache': self.case_synopsis_cache,
            'teaching_point_cache': self.teaching_point_cache,
            'processed_cases': self.processed_cases,
        }
        save_state(state)

    async def process_cases(self):
        
        try:
            # Extract structured data from Google Sheets
            logging.info("Attempting to extract Google Sheet data")
            course_data = self.sheet_handler.extract_course_data()

            # Load state if it exists
            state = load_state()
            if state:
                course_data = state['course_data']
                case_synopsis_data = state['case_synopsis_data']
                teaching_point_data = state['teaching_point_data']
                self.case_synopsis_cache = state['case_synopsis_cache']
                self.teaching_point_cache = state['teaching_point_cache']
                self.processed_cases = state['processed_cases']
                logging.info("LOADED STATE FROM CHECKPOINT")
            else:
                # Initialize data stores if they do not exist in state
                case_synopsis_data = []
                teaching_point_data = []

            tasks = []

            # Process each course asynchronously
            logging.info("Moving on to async processing of cases")
            for course in course_data:
                course_name = course["Course"]  # Strip is already handled in extract_course_data
                course_url = self.get_course_url(course_name)
                await self.scraper.navigate_to_repository(course_url)
            
                case_name = course["Case Name"]
                teaching_point_name = course["Teaching Point"]

                if not case_name or (case_name in self.case_synopsis_cache and (case_name, teaching_point_name) in self.processed_cases):
                    continue

                if case_name not in self.case_synopsis_cache:
                    tasks.append(asyncio.create_task(self.scrape_and_cache_synopsis(case_name)))
                
                if teaching_point_name:
                    teaching_key = (case_name, teaching_point_name)
                    if teaching_key not in self.teaching_point_cache:
                        tasks.append(asyncio.create_task(self.scrape_and_cache_teaching_point(course_url, case_name, teaching_point_name)))

                # Add this row's data to our save state
                self.processed_cases.add((case_name, teaching_point_name))

                # Save periodically to avoid data loss
                if len(self.processed_cases) % 25 == 0:
                    self.create_state(course_data, case_synopsis_data, teaching_point_data)
                    logging.info("SAVED STATE after 25 more rows")

            await asyncio.gather(*tasks)

            for course in course_data:
                case_name = course["Case Name"]
                teaching_point_name = course["Teaching Point"]

                case_synopsis_data.append({
                    "Case Name": case_name,
                    "Case Synopsis": self.case_synopsis_cache[case_name]
                })

                teaching_point_data.append({
                    "Teaching Point": teaching_point_name,
                    "Full Text": self.teaching_point_cache[(case_name, teaching_point_name)]
                })

            # Save the final state
            self.create_state(course_data, case_synopsis_data, teaching_point_data)
            
            # Write data back to Google Sheets
            self.sheet_handler.write_column("EK", case_synopsis_data)
            self.sheet_handler.write_column("EL", teaching_point_data)
        except Exception as e:
            logging.error(f"Error in Coordinator:process_cases: {e}")

    async def scrape_and_cache_synopsis(self, case_name):
        logging.info(f"Synopsis scrape and cache task added for case: <<{case_name}>>")
        synopsis_result = await self.scraper.scrape_case_synopsis(case_name)
        logging.info(f"Synopsis for {case_name} scraped: <<{synopsis_result}>>")
        self.case_synopsis_cache[case_name] = synopsis_result

    async def scrape_and_cache_teaching_point(self, course_url, case_name, teaching_point_name):
        logging.info(f"Teaching Point scrape and cache task added for case: <<{case_name}>> and TP: <<{teaching_point_name}>>")
        teaching_point_result = await self.scraper.scrape_teaching_point(course_url, case_name, teaching_point_name)
        logging.info(f"TP Full Text scraped for case {case_name} and TP {teaching_point_name}: <<{teaching_point_result}>>")
        self.teaching_point_cache[(case_name, teaching_point_name)] = teaching_point_result

async def main():
    # Set up your Google Sheet credentials and initialize the handler
    sheet_handler = GoogleSheetHandler(spreadsheet_id="Your_Spreadsheet_ID", credentials="GoogleCloudCredentials.json")
    
    scraper = WebScraper(base_url="https://example.com")
    await scraper.setup_browser()

    try:

        # Navigate to initial login page
        logging.info("Navigating to initial login page")
        await scraper.attempt_login_page()

        # 2FA Authentication
        logging.info("Running 2FA authentication")
        code = await scraper.get_2fa_code(Org_UN, Org_PW)
        logging.info(f"2FA code retrieved from email: {code}")
        if not code:
            logging.info("Fallback to manual entry for 2FA code")
            code = input("Please enter the 2FA code: ")

        logging.info(f"2FA code received from user: {code}")
        if not code:
            raise Exception("2FA code is not retrieved properly")

        # Input 2FA code (Assuming input field locator is known)
        await scraper.page.wait_for_load_state("domcontentloaded", timeout=10000)
        logging.info(f"Arrived at 2FA page: {scraper.page.url}")
        await scraper.page.get_by_label("Please enter the time-").fill(code)
        
        # Attempt manual submission with short timeouts
        try:
            await scraper.page.locator("input[type=submit]").first.click(timeout=1000)
            logging.info("Clicked submit button on 2FA page using CSS Selector")
        except Exception as e:
            logging.info(f"Failed to submit 2FA code for Org sign-in: {e}")
            try:
                await scraper.page.get_by_text("submit").click(timeout=1000)
                logging.info("Clicked submit on 2FA page using get_by_text method")
            except Exception as e:
                logging.error(f"Failed to click submit using fallback: {e}")

        # wait for page redirect with or without successful click, confirm proper navigation
        try:
            await scraper.page.wait_for_url("https://example.com", timeout=15000)
            await scraper.page.wait_for_load_state("domcontentloaded", timeout=10000)
        except Exception as e:
            logging.error(f"Error while waiting for target URL or load state: {e}")

        logging.info("Finished login steps, proceeding to Coordinator async methods")

        coordinator = Coordinator(sheet_handler, scraper)

        logging.info("Processing cases with the coordinator class")
    
        await coordinator.process_cases()
    
    except Exception as e:
        logging.error(f"Error in main method try block: {e}")
    
    finally:
        await scraper.close_browser()

if __name__ == "__main__":
    asyncio.run(main())