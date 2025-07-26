
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
import pandas as pd


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
    max_scroll_attempts = 200  # Increased attempts for more data
    no_change_count = 0
    max_no_change = 10  # Increased tolerance for no change

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
                                # Use a unique identifier, combining _id and id_text if available
                                entry_id = vocab_entry.get('_id', '') + vocab_entry.get('id_text', '')
                                if not any((existing.get('_id', '') + existing.get('id_text', '') == entry_id)
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

    # Final data extraction to ensure all last-minute entries are caught
    final_data = extract_network_data(driver)
    for response_data in final_data:
        if 'responses' in response_data:
            for response in response_data['responses']:
                if 'hits' in response and 'hits' in response['hits']:
                    for hit in response['hits']['hits']:
                        if '_source' in hit:
                            vocab_entry = hit['_source']
                            entry_id = vocab_entry.get('_id', '') + vocab_entry.get('id_text', '')
                            if not any((existing.get('_id', '') + existing.get('id_text', '') == entry_id)
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


def create_xml_from_vocabulary(vocabulary_data, filename="english_profile_all_vocabulary.xml"):
    """Create XML file from vocabulary data"""
    print(f"📝 Creating XML file with {len(vocabulary_data)} entries...")

    # Create root element
    root = ET.Element("english_profile_vocabulary")
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
            'Modified Date': 'modified_date',
            'guideword_text': 'guideword_text'
        }

        # Add basic fields
        for json_key, xml_key in fields_mapping.items():
            if json_key in entry and entry[json_key] is not None:  # Check for None as well
                field_elem = ET.SubElement(vocab_elem, xml_key)
                field_elem.text = clean_text(entry[json_key])

        # Handle complex fields (lists and lookups)
        if 'l_grammars_list_custom_evp_l_grammar' in entry and entry['l_grammars_list_custom_evp_l_grammar']:
            grammars_elem = ET.SubElement(vocab_elem, "grammars")
            for grammar in entry['l_grammars_list_custom_evp_l_grammar']:
                grammar_elem = ET.SubElement(grammars_elem, "grammar")
                grammar_elem.text = clean_text(grammar)

        if 'l_topics_list_custom_evp_l_topic' in entry and entry['l_topics_list_custom_evp_l_topic']:
            topics_elem = ET.SubElement(vocab_elem, "topics")
            for topic in entry['l_topics_list_custom_evp_l_topic']:
                topic_elem = ET.SubElement(topics_elem, "topic")
                topic_elem.text = clean_text(topic)

        # Add any other fields that might exist and are not explicitly mapped
        for key, value in entry.items():
            if key not in fields_mapping and not key.startswith('l_') and value is not None:
                # Create element for unmapped fields
                clean_key = re.sub(r'[^a-zA-Z0-9_]', '_', key).lower()
                # Ensure the key is not empty and not already added as a direct mapping
                if clean_key and clean_key not in fields_mapping.values() and not any(
                        child.tag == clean_key for child in vocab_elem):
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


def create_excel_from_vocabulary(vocabulary_data, filename="english_profile_all_vocabulary.xlsx"):
    """Create Excel file with teaching-relevant fields only"""
    print(f"📊 Creating Excel file with {len(vocabulary_data)} entries...")

    # Define teaching-relevant fields
    teaching_fields = {
        'base_text': 'Word',
        'hw_text': 'Headword',
        'definition_text': 'Definition',
        'learnerexamples_text': 'Examples',
        'pos_text': 'Part of Speech',
        'cefr_text_text': 'CEFR Level',
        'ukpron_text': 'UK Pronunciation',
        'guideword_text': 'Guide Word',
        'l_topic_text_text': 'Topic',
        'searchterms_text': 'Search Terms',
        'audiofilename_text': 'Audio File'
    }

    # Prepare data for DataFrame
    excel_data = []

    for entry in vocabulary_data:
        row = {}

        # Extract basic fields
        for json_key, excel_col in teaching_fields.items():
            value = entry.get(json_key, '')

            # Clean and format the value
            if value is not None:
                # Handle different data types
                if isinstance(value, (list, tuple)):
                    value = '; '.join(str(v) for v in value)
                else:
                    value = str(value).strip()

                # Clean text for Excel
                value = clean_text_for_excel(value)
            else:
                value = ""  # Ensure None values are empty strings

            row[excel_col] = value

        # Add derived/calculated fields that might be useful for teaching

        # Word length (useful for difficulty assessment)
        base_word = row.get('Word', '')
        row['Word Length'] = len(base_word.split()[0]) if base_word else 0

        # Number of examples
        examples = row.get('Examples', '')
        row['Example Count'] = len([ex.strip() for ex in examples.split(';') if ex.strip()]) if examples else 0

        # Has pronunciation guide
        row['Has Pronunciation'] = 'Yes' if row.get('UK Pronunciation', '') else 'No'

        # Has audio
        row['Has Audio'] = 'Yes' if row.get('Audio File', '') else 'No'

        # Simple/Complex definition (word count)
        definition = row.get('Definition', '')
        row['Definition Length'] = len(definition.split()) if definition else 0

        excel_data.append(row)

    # Create DataFrame
    df = pd.DataFrame(excel_data)

    # Reorder columns for better teaching use
    preferred_order = [
        'Word', 'Headword', 'Part of Speech', 'CEFR Level',
        'Definition', 'Guide Word', 'Examples',
        'UK Pronunciation', 'Topic', 'Has Pronunciation', 'Has Audio',
        'Word Length', 'Example Count', 'Definition Length',
        'Search Terms', 'Audio File'
    ]

    # Reorder columns (keep only existing ones)
    existing_cols = [col for col in preferred_order if col in df.columns]
    remaining_cols = [col for col in df.columns if col not in existing_cols]
    final_order = existing_cols + remaining_cols

    df = df[final_order]

    # Sort by word alphabetically
    df = df.sort_values('Word', na_position='last')

    # Create Excel file with formatting
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Main vocabulary sheet
        df.to_excel(writer, sheet_name='All Vocabulary', index=False)

        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['All Vocabulary']

        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            # Set reasonable width limits
            adjusted_width = min(max_length + 2, 75)  # Increased max width for longer texts
            worksheet.column_dimensions[column_letter].width = adjusted_width

        # Create summary statistics sheet
        create_summary_sheet(df, writer, workbook)

    print(f"✅ Excel file created: {filename}")
    return filename


def clean_text_for_excel(text):
    """Clean text specifically for Excel compatibility"""
    if not text:
        return ""

    text = str(text)

    # Remove problematic characters for Excel
    text = re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]', '', text)

    # Replace HTML entities that might have been decoded
    text = text.replace('&amp;', '&')
    text = text.replace('&lt;', '<')
    text = text.replace('&gt;', '>')
    text = text.replace('&quot;', '"')
    text = text.replace('&apos;', "'")

    # Clean up extra whitespace
    text = ' '.join(text.split())

    return text.strip()


def create_summary_sheet(df, writer, workbook):
    """Create a summary statistics sheet"""
    summary_data = []

    # Basic statistics
    total_words = len(df)
    summary_data.append(['Total Words', total_words])

    # Part of speech distribution
    pos_counts = df['Part of Speech'].value_counts().sort_index()  # Sort for consistent order
    summary_data.append(['', ''])  # Empty row
    summary_data.append(['Part of Speech Distribution', ''])
    for pos, count in pos_counts.items():
        summary_data.append([f'  {pos}', count])

    # CEFR level distribution
    cefr_counts = df['CEFR Level'].value_counts().sort_index()  # Sort for consistent order
    summary_data.append(['', ''])  # Empty row
    summary_data.append(['CEFR Level Distribution', ''])
    for level, count in cefr_counts.items():
        summary_data.append([f'  {level}', count])

    # Topic distribution (top 15, as there might be many)
    topic_counts = df['Topic'].value_counts().head(15)
    summary_data.append(['', ''])  # Empty row
    summary_data.append(['Top 15 Topics', ''])
    for topic, count in topic_counts.items():
        topic_name = topic if len(str(topic)) < 30 else str(topic)[:30] + '...'
        summary_data.append([f'  {topic_name}', count])

    # Audio availability
    audio_count = (df['Has Audio'] == 'Yes').sum()
    summary_data.append(['', ''])  # Empty row
    summary_data.append(['Audio Availability', ''])
    summary_data.append(['  Words with Audio', audio_count])
    summary_data.append(['  Words without Audio', total_words - audio_count])

    # Pronunciation availability
    pron_count = (df['Has Pronunciation'] == 'Yes').sum()
    summary_data.append(['', ''])  # Empty row
    summary_data.append(['Pronunciation Availability', ''])
    summary_data.append(['  Words with Pronunciation', pron_count])
    summary_data.append(['  Words without Pronunciation', total_words - pron_count])

    # Average Word Length
    avg_word_length = df['Word Length'].mean()
    summary_data.append(['', ''])
    summary_data.append(['Average Word Length (of base word)', f'{avg_word_length:.2f}'])

    # Average Definition Length
    avg_def_length = df['Definition Length'].mean()
    summary_data.append(['Average Definition Length (words)', f'{avg_def_length:.2f}'])

    # Average Example Count
    avg_example_count = df['Example Count'].mean()
    summary_data.append(['Average Example Count', f'{avg_example_count:.2f}'])

    # Create summary DataFrame
    summary_df = pd.DataFrame(summary_data, columns=['Metric', 'Value'])

    # Write to Excel
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

    # Format the summary sheet
    summary_sheet = writer.sheets['Summary']
    summary_sheet.column_dimensions['A'].width = 40
    summary_sheet.column_dimensions['B'].width = 15


def main():
    """Main function to run the scraper"""
    chromedriver_path = r"D:\EnglishScraper\chromedriver.exe"

    print("🚀 Starting Enhanced English Profile Scraper (All Difficulties)...")

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

        # Skip C2 level selection as we want all difficulties (default)
        print("ℹ️ Skipping C2 level selection. All difficulty levels will be collected by default.")

        # Click search
        if click_search_button(driver, wait):
            print("✅ Search initiated successfully!")

            # Wait for initial results
            print("⏳ Waiting for initial results...")
            time.sleep(5)  # Increased wait time for initial load

            # Select 'All' from dropdown
            if select_all_from_dropdown(driver, wait):
                print("✅ 'All' selected from dropdown!")

                # Scroll and collect all vocabulary data
                vocabulary_data = scroll_and_collect_data(driver, wait)

                if vocabulary_data:
                    # Create both XML and Excel files
                    xml_filename = create_xml_from_vocabulary(vocabulary_data)
                    excel_filename = create_excel_from_vocabulary(vocabulary_data)

                    print(f"\n🎯 Successfully scraped {len(vocabulary_data)} vocabulary entries!")
                    print(f"📄 XML file saved as: {xml_filename}")
                    print(f"📊 Excel file saved as: {excel_filename}")

                    # Show sample of extracted data
                    print(f"\n📝 Sample vocabulary entries:")
                    for i, entry in enumerate(vocabulary_data[:5], 1):  # Display more samples
                        word = entry.get('base_text', 'N/A')
                        definition = entry.get('definition_text', 'N/A')
                        # Truncate definition for display
                        definition_display = definition[:100] + "..." if len(definition) > 100 else definition
                        cefr = entry.get('cefr_text_text', 'N/A')
                        pos = entry.get('pos_text', 'N/A')
                        print(f"  {i}. {word} ({pos}, {cefr}): {definition_display}")

                else:
                    print("❌ No vocabulary data collected!")

            else:
                print("❌ Failed to select 'All' from dropdown.")
        else:
            print("❌ Failed to click search button.")

    except Exception as e:
        print(f"❌ Unexpected error: {str(e)}")
        import traceback
        traceback.print_exc()

    finally:
        print("🔄 Keeping browser open for 10 seconds for review...")
        time.sleep(10)
        driver.quit()
        print("✅ Browser closed.")


if __name__ == "__main__":
    main()
