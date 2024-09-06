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
import asyncio


# Setup Pickle Save and Load States
CHECKPOINT_FILE = 'scrape-cd.pkl'

def save_state(state):
    with open(CHECKPOINT_FILE, 'wb') as f:
        pickle.dump(state, f)

def load_state():
    if os.path.exists(CHECKPOINT_FILE):
        with open(CHECKPOINT_FILE, 'rb') as f:
            return pickle.load(f)
    return None

# Configure logging
logging.basicConfig(filename='scrape_cd.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# locate .env
dotenv_path = find_dotenv()
if not dotenv_path:
    raise FileNotFoundError(".env file not found.")
load_dotenv(dotenv_path)

# Environment Variables for Org
Org_UN = os.environ.get('Org_User_ID')
Org_PW = os.environ.get('Org_Password')

# Ensure Org_UN and Org_PW are obtained
if not Org_UN or not Org_PW:
    raise ValueError("Org_User_ID or Org_Password environment variables are not set.")

# Log for debugging values
logging.info(f"Org_UN={Org_UN}")
logging.info(f"Org_PW={'*' * len(Org_PW)}")  # Masks the password for privacy

# Create a dictionary to store course names and urls for case listings by course
courses = {
    "Geriatrics": "https://placeholder.org.com/document_sets/4886",
    "Radiology": "https://placeholder.org.com/document_sets/4890",
    "Diagnostic": "https://placeholder.org.com/document_sets/19560",
    "DX": "https://placeholder.org.com/document_sets/19560",
    "Telemedicine": "https://placeholder.org.com/document_sets/39733",
    "High": "https://placeholder.org.com/document_sets/19561",
    "Palliative": "https://placeholder.org.com/document_sets/39732",
    "Social": "https://placeholder.org.com/document_sets/25437",
    "Trauma-Informed": "https://placeholder.org.com/document_sets/34397",
    "Family": "https://placeholder.org.com/document_sets/4893",
    "Internal": "https://placeholder.org.com/document_sets/4897",
    "Neurology": "https://placeholder.org.com/document_sets/39734",
    "Obstetrics": "https://placeholder.org.com/document_sets/44406",
    "Gynecology": "https://placeholder.org.com/document_sets/44406",
    "Pediatrics": "https://placeholder.org.com/document_sets/4891",
    "Addiction": "https://placeholder.org.com/document_sets/4898"
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
            # Check the sheet size to verify whether enough columns for our writing task
            sheet_properties = self.sheet.spreadsheet.fetch_sheet_metadata()
            current_columns = sheet_properties['sheets'][0]['properties']['gridProperties']['columnCount']
            required_columns = gspread.utils.a1_to_rowcol(f"{column}1")[1]

            # Expand the grid if necessary
            if current_columns < required_columns:
                self.sheet.add_cols(required_columns - current_columns)
                logging.info(f"Expanded the sheet to {required_columns} columns.")

            # Iterate through the data
            for idx, entry in enumerate(data):
                row = entry.get("Row")
                new_value = entry.get(f"{'Case Synopsis' if column=='EK' else 'Full Text'}", "")

                # Write the data to corresponding cell
                self.sheet.update_cell(row, required_columns, new_value)

                # Sleep 1 second for each write to avoid rate limits of 60 reqs / user / proj / min
                if idx < len(data) - 1:
                    time.sleep(1)

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
        await self.page.goto("https://placeholder.org.com/users/sign_in")
        await self.page.wait_for_load_state("networkidle")
        logging.info(f"Accessed Org home page: {self.page.url}")
        await self.page.get_by_label("Email").fill(Org_UN)
        await self.page.get_by_label("Password").fill(Org_PW)
        await self.page.keyboard.press("Enter")

    async def navigate_to_repository(self, url):
        await self.page.goto(url)
        await self.page.wait_for_load_state("networkidle")


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
                    logging.info("Clicked Okta Sign in button using Xpath")
                except Exception as e:
                    logging.error(f"Failed to click Okta Sign in using Xpath: {e}")

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
   
    async def scrape_case(self, case_name, course_url, case_scrapes): # need to add any other necessary attributes
        try:
            # Enter a case using existing methods
            page = await self.context.new_page()
            await page.goto(course_url)
            await page.wait_for_load_state("domcontentloaded")
            await self.scroll_around()
            await page.locator(f'a:has-text("{case_name}")').first.click()
            await page.wait_for_function("() => window.location.href.includes('/document_set_document_relations')", timeout=10000)
            await page.wait_for_load_state("networkidle")
            logging.info(f"Entered case: {case_name} at url: {page.url}")
            # Change the viewing mode to 'full'
            selector = page.locator('select.doc-controls-select.doc-controls-view-mode').first
            await selector.select_option("full")
            await page.wait_for_load_state("domcontentloaded")
            # Pull the entire text content or html content of the page
            html_content = await page.content()
            text_content = await page.text_content("div.doc-body.full-display-mode")
            await page.close()
            # return the raw scrape
            case_scrapes[case_name] = {
                "html_content": html_content,
                "text_content": text_content
            }
            logging.info(f"Scraped content for case {case_name}:\nHTML snip - {html_content[:150]} \nText snip - {text_content[:150]}")
        except Exception as e:
            logging.error(f"Error in scrape_case for case: {case_name} - {e}")

        finally:
            return case_scrapes
   
    async def get_case_names(self, course_url, course_name, case_names):
        try:
            await self.page.goto(course_url)
            await self.page.wait_for_load_state("networkidle")
            # Pull full html of repository page
            html_content = await self.page.content()
            # Parse HTML content using Soup
            soup = BeautifulSoup(html_content, 'html.parser')
            # Find all <a> elements with the target class
            case_elements = soup.find_all('a', class_='case-name-link case-name-container')
            # Extract & clean the text content for each case name
            for element in case_elements:
                case_name = element.get_text(strip=True)
                case_names.append(case_name)
            logging.info(f"Extracted case names from course repository: {course_name}: {course_url}")
        except Exception as e:
            logging.error(f"Error in get_case_names for {course_name}: {course_url} - {e}")
        finally:
            return case_names

    def parse_teaching_point(self, case_scrape, teaching_point_name):
        try:
            # Get the HTML content from the case_scrape
            html_content = case_scrape['html_content']

            # Parse the HTML content using BeautifulSoup
            soup = BeautifulSoup(html_content, 'html.parser')

            try:
                # Locate all doc-sections
                doc_sections = soup.find_all('div', class_=re.compile("doc-section"))
                logging.info(f"Total doc-sections found: {len(doc_sections)} for html-TP search: '{teaching_point_name}'")

                for section in doc_sections:
                    # Check for the presence of a teaching-point-topper within this section
                    topper = section.find('div', class_=re.compile("teaching-point-topper"))
                    logging.info(f"Located TP Topper container for '{teaching_point_name}', searching for title header next.")
                    if topper:
                        # Locate the h1 element within the section header
                        header = section.find('h1', class_=re.compile("doc-section-header-title"))
                        logging.info(f"Located TP Header Title container for '{teaching_point_name}', checking match for TP next.")
                        if header and header.get_text(strip=True) == teaching_point_name:
                            # If a match is found, extract all text from the doc-children within the doc-section-body
                            body = section.find('div', class_=re.compile("doc-section-body"))
                            logging.info(f"Located a doc-section-body match for TP Title '{teaching_point_name}', extracting text next.")
                            if body:
                                full_text = body.get_text(separator="\n", strip=True)
                                logging.info(f"Located Full Text for TP: '{teaching_point_name}' - Full Text: {full_text[:150]}")
                                return full_text

                logging.error(f'Teaching point: {teaching_point_name} not found in the given case scrape.')
                return None
            
            except Exception as e:
                logging.error(f"Error in 1st try of parse_teaching_point for TP name: {teaching_point_name} - {e}")
                try:
                    # Alternative search attempt
                    doc_section = soup.find('div', attrs={f'"class": "doc-section-header", "string":"{re.compile(teaching_point_name)}"'})
                    if doc_section:
                        element = doc_section.find_parent('div', class_='doc-section')
                        logging.info(f"Located following match for doc-section on 2nd try: {doc_section[:150]}")
                        body = element.find('div', class_='doc-section-body')
                        if body:
                            full_text = body.get_text(separator="\n", strip=True)
                            logging.info(f"Located Full Text for TP in alternative search: '{teaching_point_name}' - Unclean Text: {full_text[:150]}")
                            return full_text
                except Exception as e:
                    logging.error(f'Error in 2nd try of parse_teaching_point for TP name: {teaching_point_name} - {e}\nFAILED TP EXTRACTION FROM HTML - Examine HTML here:\n{body}')
                    return None

        except Exception as e:
            logging.error(f'Error in parse_teaching_point for teaching_point_name: {teaching_point_name} - {e}')
            return None

    def parse_synopsis(self, case_scrape):
        try:
            # Get the text content from the case_scrape
            text_content = case_scrape['text_content']
            
            # Extract text between "Case Synopsis" header and "Thank you for completing {case_name}"
            # Adjusted to remove the node tag "CASE SYNOPSIS" by using it as the start_marker
            start_marker = "Case Synopsis"
            end_marker = "Thank you for completing"
            
            try:
                start_index = text_content.index(start_marker) + len(start_marker)
                end_index = text_content.index(end_marker, start_index)
            except ValueError as e:
                logging.error(f"Markers not found in text_content for case: {e}")
                return None
            
            raw_synopsis = text_content[start_index:end_index].strip()
            
            # Clean the raw extracted synopsis text
            clean_synopsis = self.clean_text_content(raw_synopsis)
            logging.info(f"Parsed and cleaned synopsis: {clean_synopsis}")
            return clean_synopsis
            
        except Exception as e:
            logging.error(f'Error in parse_synopsis: {e}')
            return None

    def clean_text_content(self, text_content):
        # Remove "CASE SYNOPSIS" from the front of the synopses
        text_content = re.sub(r"\bCASE SYNOPSIS\b", "", text_content)
        text_content = re.sub(r"\bTEACHING POINT\b", "", text_content)
        # Add a colon and space after specific header
        text_content = re.sub(r"(Case Synopsis)([A-Z])", r"\1: \2", text_content)
        text_content = re.sub(r"(Teaching Point)([A-Z])", r"\1: \2", text_content)
        # Add newlines between each section and paragraph
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


class Coordinator:
    def __init__(self, sheet_handler, scraper):
        self.sheet_handler = sheet_handler
        self.scraper = scraper
        self.case_synopsis_cache = {} #add caching for case synopsis
        self.teaching_point_cache = {} #add caching for teaching points
        self.processed_cases = set()
        self.processed_index = 0 # New attribute to track last processed index

    def get_course_url(self, course_name):
        stripped_course_name = course_name.strip()
        if stripped_course_name in courses:
            return courses[stripped_course_name]
        else:
            raise ValueError(f"No URL found for course: {stripped_course_name}")
       
    def create_state(self, case_scrapes, processed_cases):
        state = {
            'case_scrapes': case_scrapes,
            'processed_cases': processed_cases,
        }
        save_state(state)


    async def process_cases(self):
        try:
            # Set Semaphore to limit concurrency
            max_concurrent_tasks = 15
            semaphore = asyncio.Semaphore(max_concurrent_tasks)
            
            # Load state if it exists
            state = load_state()
            if state:
                case_scrapes = state['case_scrapes']
                processed_cases = state['processed_cases']
                logging.info("LOADED STATE FROM CHECKPOINT")
            else:
                # Initialize data stores if they do not exist in state
                logging.info("NO CHECKPOINT FOUND - INITIALIZING FROM START")
                case_scrapes = {}
                processed_cases = set()

            tasks = []

            # Process each course asynchronously
            logging.info("Starting async scraping of courses")

            async def sem_scrape_case(case_name, course_url, case_scrapes):
                async with semaphore:
                    await self.scraper.scrape_case(case_name, course_url, case_scrapes)

            for course_name, course_url in courses.items():
                await self.scraper.navigate_to_repository(course_url)
                await self.scraper.scroll_around()

                # Get case names from the repository
                case_names = []
                await self.scraper.get_case_names(course_url, course_name, case_names)

                counter = len(processed_cases)
                for case_name in case_names:
                    if case_name not in case_scrapes:
                        
                        tasks.append(asyncio.create_task(sem_scrape_case(case_name, course_url, case_scrapes)))
                        processed_cases.add(case_name)

                        # Save periodically to avoid data loss
                        if len(processed_cases) % 10 == 0:
                            self.create_state(case_scrapes, processed_cases)
                            logging.info(f"SAVED STATE after 10 more cases.")
                logging.info(f"Added {len(processed_cases) - counter} case scraping tasks in {course_name}")

            await asyncio.gather(*tasks)

            # Now that we have all the scrapes, process and update the Google Sheet
            logging.info("Processing scraped data and updating Google Sheets")

            await self.scraper.close_browser()

            # Processing data from Google Sheets
            course_data = self.sheet_handler.extract_course_data()
            case_synopsis_data = []
            teaching_point_data = []

            for idx, row in enumerate(course_data):
                case_name = row["Case Name"]
                teaching_point_name = row["Teaching Point"]

                synopsis = None
                full_teaching_point = None

                # Check if the case name has been scraped, if so parse the necessary data
                if case_name in case_scrapes:
                    clean_synopsis = self.scraper.parse_synopsis(case_scrapes[case_name])
                    if clean_synopsis:
                        synopsis = {"Row": idx + 2, "Case Name": case_name, "Case Synopsis": clean_synopsis}  # +2 for Google Sheets index

                # Check if the teaching point needs to be parsed, if so parse it
                if case_name in case_scrapes and teaching_point_name:
                    clean_teaching_point = self.scraper.parse_teaching_point(case_scrapes[case_name], teaching_point_name)
                    if clean_teaching_point:
                        full_teaching_point = {"Row": idx + 2, "Teaching Point": teaching_point_name, "Full Text": clean_teaching_point}

                if synopsis:
                    case_synopsis_data.append(synopsis)

                if full_teaching_point:
                    teaching_point_data.append(full_teaching_point)

            # Write data back to Google Sheets
            if case_synopsis_data:
                self.sheet_handler.write_column("EK", case_synopsis_data)
            if teaching_point_data:
                self.sheet_handler.write_column("EL", teaching_point_data)

            logging.info("Data successfully written back to Google Sheets")

        except Exception as e:
            logging.error(f"Error in Coordinator:process_cases: {e}")


async def main():
    # Set up your Google Sheet credentials and initialize the handler
    sheet_handler = GoogleSheetHandler(spreadsheet_id="1dpK7QX-MtHgVV1lpH1ZFR4FpOxFtOz2QHxB1tASgFNY", credentials="GoogleCloudCredentials.json")
   
    scraper = WebScraper(base_url="https://placeholder.org.com")
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
            # await scraper.page.locator(scraper.page, "input[type=submit]", """//*[@id="sign_in_form"]/div/div/div/input""", "commit", "input")

        # wait for page redirect with or without successful click, confirm proper navigation
        try:
            await scraper.page.wait_for_url("https://placeholder.org.com", timeout=10000)
            await scraper.page.wait_for_load_state("domcontentloaded", timeout=10000)
        except Exception as e:
            logging.error(f"Error while waiting for target URL or load state: {e}")


        logging.info("Finished login steps, proceeding to Coordinator async methods")


        coordinator = Coordinator(sheet_handler, scraper)

        logging.info("Processing cases with the coordinator class")
   
        await coordinator.process_cases()
   
    except Exception as e:
        logging.error(f"Error in main method try block: {e}")


if __name__ == "__main__":
    asyncio.run(main())