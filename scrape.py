from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import time
import json
import xml.etree.ElementTree as ET
from xml.dom import minidom
import re
from datetime import datetime


def setup_driver(chromedriver_path):
    """Setup Chrome driver with network logging enabled"""
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    # Enable logging for network requests
    chrome_options.add_argument("--enable-logging")
    chrome_options.add_argument("--log-level=0")
    chrome_options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})

    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    driver.maximize_window()
    return driver


def handle_cookie_popup(driver, wait):
    """Handle cookie popup if it exists"""
    cookie_selectors = [
        "//button[contains(text(), 'Accept')]",
        "//button[contains(text(), 'OK')]",
        "//button[contains(text(), 'Agree')]",
        "//button[@id='cookie-accept']",
        "//div[@class='cookie-banner']//button"
    ]

    for selector in cookie_selectors:
        try:
            cookie_button = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, selector))
            )
            cookie_button.click()
            print("✅ Cookie popup accepted.")
            time.sleep(1)
            return True
        except TimeoutException:
            continue

    print("ℹ️ No cookie popup found.")
    return False


def select_c2_level(driver, wait):
    """Select C2 level based on the actual HTML structure"""
    c2_selectors = [
        "//div[contains(@class, 'clickable-element') and div[text()='C2']]",
        "//div[contains(@class, 'Text') and contains(@class, 'clickable-element')]//div[text()='C2']/parent::div",
        "//div[text()='C2' and ancestor::div[contains(@class, 'clickable-element')]]",
        "//div[contains(@class, 'entry-6')]//div[contains(@class, 'clickable-element')]",
        "//div[contains(@style, 'background-color: rgb(160, 48, 160)') and contains(@class, 'clickable-element')]"
    ]

    print("🔍 Searching for C2 level selector...")

    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'RepeatingGroup')]")))
        print("✅ Level selector container loaded.")
        time.sleep(2)
    except TimeoutException:
        print("❌ Level selector container not found.")
        return False

    for i, selector in enumerate(c2_selectors):
        try:
            print(f"🔍 Trying selector {i + 1}: {selector}")
            c2_element = wait.until(EC.presence_of_element_located((By.XPATH, selector)))

            if c2_element.is_displayed():
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", c2_element)
                time.sleep(0.5)

                try:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, selector)))
                    c2_element.click()
                    print("✅ C2 selected with normal click!")
                    time.sleep(1)
                    return True
                except:
                    driver.execute_script("arguments[0].click();", c2_element)
                    print("✅ C2 selected with JavaScript click!")
                    time.sleep(1)
                    return True

        except TimeoutException:
            continue
        except Exception as e:
            continue

    # Debug search
    try:
        print("\n🔍 DEBUG: Searching for all level elements...")
        level_elements = driver.find_elements(By.XPATH, "//div[contains(@class, 'clickable-element')]//div")

        for i, element in enumerate(level_elements):
            try:
                text = element.text.strip()
                if text and len(text) <= 3:
                    if text == "C2":
                        print(f"🎯 Found C2 element, attempting click...")
                        parent_element = element.find_element(By.XPATH, "./parent::div")
                        driver.execute_script("arguments[0].click();", parent_element)
                        print("✅ C2 clicked via debug search!")
                        return True
            except Exception:
                continue
    except Exception as e:
        print(f"Debug search failed: {e}")

    return False


def click_search_button(driver, wait):
    """Click the search button"""
    search_selectors = [
        "//span[text()='Search']/ancestor::button",
        "//button[contains(text(), 'Search')]",
        "//input[@type='submit' and @value='Search']",
        "//button[@type='submit']",
        "//div[contains(@class, 'clickable-element') and contains(text(), 'Search')]",
        "//*[text()='Search' and (@role='button' or contains(@class, 'button') or contains(@class, 'clickable'))]"
    ]

    print("🔍 Searching for Search button...")

    for i, selector in enumerate(search_selectors):
        try:
            print(f"Trying search selector {i + 1}: {selector}")
            search_button = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", search_button)
            time.sleep(0.5)
            search_button.click()
            print("✅ Search button clicked!")
            time.sleep(2)
            return True
        except TimeoutException:
            continue
        except Exception as e:
            continue

    return False


def select_all_from_dropdown(driver, wait):
    """Select 'All' from the display dropdown"""
    print("🔍 Looking for display dropdown...")

    dropdown_selectors = [
        "//select[contains(@class, 'Dropdown')]",
        "//select[contains(@class, 'dropdown')]",
        "//select[option[text()='All']]",
        "//div[text()='Display: ']/following-sibling::select",
        "//div[contains(text(), 'Display')]/following-sibling::select"
    ]

    for i, selector in enumerate(dropdown_selectors):
        try:
            print(f"Trying dropdown selector {i + 1}: {selector}")
            dropdown_element = wait.until(EC.presence_of_element_located((By.XPATH, selector)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_element)
            time.sleep(0.5)

            select = Select(dropdown_element)
            select.select_by_visible_text('All')
            print("✅ Selected 'All' from dropdown!")
            time.sleep(2)
            return True

        except TimeoutException:
            continue
        except Exception as e:
            continue

    return False


def extract_network_data(driver):
    """Extract msearch responses from network logs"""
    print("🔍 Extracting network data...")

    # Get network logs
    logs = driver.get_log('performance')
    msearch_responses = []

    for log in logs:
        try:
            message = json.loads(log['message'])

            # Look for network responses
            if message['message']['method'] == 'Network.responseReceived':
                response = message['message']['params']['response']

                # Check if this is an msearch request
                if 'msearch' in response.get('url', ''):
                    request_id = message['message']['params']['requestId']
                    print(f"📡 Found msearch response: {response['url']}")

                    # Get response body
                    try:
                        response_body = driver.execute_cdp_cmd('Network.getResponseBody', {'requestId': request_id})
                        if response_body.get('body'):
                            try:
                                # Parse JSON response
                                json_data = json.loads(response_body['body'])
                                msearch_responses.append(json_data)
                                print(f"✅ Extracted msearch data with {len(json_data.get('responses', []))} responses")
                            except json.JSONDecodeError:
                                print("❌ Failed to parse JSON response")
                    except Exception as e:
                        print(f"❌ Failed to get response body: {e}")

        except Exception as e:
            continue

    print(f"📊 Total msearch responses extracted: {len(msearch_responses)}")
    return msearch_responses


def scroll_and_collect_data(driver, wait):
    """Scroll down and collect all msearch data"""
    print("🔄 Starting to scroll and collect data...")

    all_vocabulary_data = []
    time.sleep(3)  # Wait for initial results

    last_height = driver.execute_script("return document.body.scrollHeight")
    scroll_attempts = 0
    max_scroll_attempts = 100
    no_change_count = 0
    max_no_change = 5

    while scroll_attempts < max_scroll_attempts:
        # Extract current network data before scrolling
        current_data = extract_network_data(driver)

        # Process and add new vocabulary entries
        for response_data in current_data:
            if 'responses' in response_data:
                for response in response_data['responses']:
                    if 'hits' in response and 'hits' in response['hits']:
                        for hit in response['hits']['hits']:
                            if '_source' in hit:
                                vocab_entry = hit['_source']
                                # Check if we already have this entry (by ID)
                                entry_id = vocab_entry.get('_id', vocab_entry.get('id_text', ''))
                                if not any(existing.get('_id') == entry_id or existing.get('id_text') == entry_id
                                           for existing in all_vocabulary_data):
                                    all_vocabulary_data.append(vocab_entry)

        # Scroll down
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)  # Wait for content to load

        # Check new height
        new_height = driver.execute_script("return document.body.scrollHeight")
        scroll_attempts += 1

        if new_height == last_height:
            no_change_count += 1
            print(
                f"🔄 Scroll {scroll_attempts}: No height change ({no_change_count}/{max_no_change}) - Vocab entries: {len(all_vocabulary_data)}")

            if no_change_count >= max_no_change:
                print("✅ No more content to load")
                break
        else:
            no_change_count = 0
            print(
                f"🔄 Scroll {scroll_attempts}: New content! Height: {last_height} → {new_height}, Vocab entries: {len(all_vocabulary_data)}")
            last_height = new_height

    # Final data extraction
    final_data = extract_network_data(driver)
    for response_data in final_data:
        if 'responses' in response_data:
            for response in response_data['responses']:
                if 'hits' in response and 'hits' in response['hits']:
                    for hit in response['hits']['hits']:
                        if '_source' in hit:
                            vocab_entry = hit['_source']
                            entry_id = vocab_entry.get('_id', vocab_entry.get('id_text', ''))
                            if not any(existing.get('_id') == entry_id or existing.get('id_text') == entry_id
                                       for existing in all_vocabulary_data):
                                all_vocabulary_data.append(vocab_entry)

    print(f"📊 Total unique vocabulary entries collected: {len(all_vocabulary_data)}")
    return all_vocabulary_data


def clean_text(text):
    """Clean text for XML compatibility"""
    if not text:
        return ""

    # Convert to string if not already
    text = str(text)

    # Remove or replace problematic characters
    text = re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\U00010000-\U0010FFFF]', '', text)

    # Replace XML special characters
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('"', '&quot;')
    text = text.replace("'", '&apos;')

    return text.strip()


def create_xml_from_vocabulary(vocabulary_data, filename="english_profile_c2_vocabulary.xml"):
    """Create XML file from vocabulary data"""
    print(f"📝 Creating XML file with {len(vocabulary_data)} entries...")

    # Create root element
    root = ET.Element("english_profile_vocabulary")
    root.set("level", "C2")
    root.set("total_entries", str(len(vocabulary_data)))
    root.set("extracted_date", datetime.now().isoformat())

    for entry in vocabulary_data:
        # Create vocabulary entry element
        vocab_elem = ET.SubElement(root, "vocabulary_entry")

        # Add all available fields
        fields_mapping = {
            '_id': 'entry_id',
            'id_text': 'id_text',
            'base_text': 'base_word',
            'hw_text': 'headword',
            'definition_text': 'definition',
            'learnerexamples_text': 'learner_examples',
            'searchterms_text': 'search_terms',
            'pos_text': 'part_of_speech',
            'cefr_text_text': 'cefr_level',
            'ukpron_text': 'uk_pronunciation',
            'audiofilename_text': 'audio_filename',
            'l_topic_text_text': 'topic',
            'culture_number': 'culture_number',
            'refid_text': 'reference_id',
            'Created Date': 'created_date',
            'Modified Date': 'modified_date'
        }

        # Add basic fields
        for json_key, xml_key in fields_mapping.items():
            if json_key in entry and entry[json_key]:
                field_elem = ET.SubElement(vocab_elem, xml_key)
                field_elem.text = clean_text(entry[json_key])

        # Handle complex fields (lists and lookups)
        if 'l_grammars_list_custom_evp_l_grammar' in entry:
            grammars_elem = ET.SubElement(vocab_elem, "grammars")
            for grammar in entry['l_grammars_list_custom_evp_l_grammar']:
                grammar_elem = ET.SubElement(grammars_elem, "grammar")
                grammar_elem.text = clean_text(grammar)

        if 'l_topics_list_custom_evp_l_topic' in entry:
            topics_elem = ET.SubElement(vocab_elem, "topics")
            for topic in entry['l_topics_list_custom_evp_l_topic']:
                topic_elem = ET.SubElement(topics_elem, "topic")
                topic_elem.text = clean_text(topic)

        # Add any other fields that might exist
        for key, value in entry.items():
            if key not in fields_mapping and not key.startswith('l_') and value:
                # Create element for unmapped fields
                clean_key = re.sub(r'[^a-zA-Z0-9_]', '_', key).lower()
                if clean_key and not any(child.tag == clean_key for child in vocab_elem):
                    other_elem = ET.SubElement(vocab_elem, clean_key)
                    other_elem.text = clean_text(value)

    # Write to file with pretty formatting
    xml_str = ET.tostring(root, encoding='unicode')
    dom = minidom.parseString(xml_str)
    pretty_xml = dom.toprettyxml(indent="  ")

    # Remove empty lines
    pretty_xml = '\n'.join([line for line in pretty_xml.split('\n') if line.strip()])

    with open(filename, 'w', encoding='utf-8') as f:
        f.write(pretty_xml)

    print(f"✅ XML file created: {filename}")
    return filename


def main():
    """Main function to run the scraper"""
    chromedriver_path = r"D:\EnglishScraper\chromedriver.exe"

    print("🚀 Starting Enhanced English Profile XML scraper...")

    # Setup driver
    driver = setup_driver(chromedriver_path)
    wait = WebDriverWait(driver, 20)

    try:
        # Navigate to page
        print("📖 Loading page...")
        driver.get("https://englishprofile.org/?menu=evp-online")
        time.sleep(3)

        # Handle cookie popup
        handle_cookie_popup(driver, wait)

        # Select C2 level
        if select_c2_level(driver, wait):
            print("✅ C2 level selected successfully!")

            # Click search
            if click_search_button(driver, wait):
                print("✅ Search initiated successfully!")

                # Wait for initial results
                print("⏳ Waiting for initial results...")
                time.sleep(3)

                # Select 'All' from dropdown
                if select_all_from_dropdown(driver, wait):
                    print("✅ 'All' selected from dropdown!")

                    # Scroll and collect all vocabulary data
                    vocabulary_data = scroll_and_collect_data(driver, wait)

                    if vocabulary_data:
                        # Create XML file
                        xml_filename = create_xml_from_vocabulary(vocabulary_data)

                        print(f"\n🎯 Successfully scraped {len(vocabulary_data)} vocabulary entries!")
                        print(f"📄 XML file saved as: {xml_filename}")

                        # Show sample of extracted data
                        print(f"\n📝 Sample vocabulary entries:")
                        for i, entry in enumerate(vocabulary_data[:3], 1):
                            word = entry.get('base_text', 'N/A')
                            definition = entry.get('definition_text', 'N/A')[:100] + "..." if len(
                                entry.get('definition_text', '')) > 100 else entry.get('definition_text', 'N/A')
                            cefr = entry.get('cefr_text_text', 'N/A')
                            pos = entry.get('pos_text', 'N/A')
                            print(f"   {i}. {word} ({pos}, {cefr}): {definition}")

                    else:
                        print("❌ No vocabulary data collected!")

                else:
                    print("❌ Failed to select 'All' from dropdown.")
            else:
                print("❌ Failed to click search button.")
        else:
            print("❌ Failed to select C2 level.")

    except Exception as e:
        print(f"❌ Unexpected error: {str(e)}")
        import traceback
        traceback.print_exc()

    finally:
        print("🔄 Keeping browser open for 10 seconds...")
        time.sleep(10)
        driver.quit()
        print("✅ Browser closed.")


if __name__ == "__main__":
    main()