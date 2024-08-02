"""
Description:
This script uses Selenium WebDriver to scrape product information from a specific category and subcategory on the Provi website. It logs in to the site, navigates through the product listings, and extracts various details about each product, saving the data to an Excel file. It also downloads product images and stores them in a local folder.

Prerequisites:
1. Download and install the Edge WebDriver: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
2. Install the required Python libraries:
   - selenium
   - requests
   - openpyxl

How to run the script:
1. Update the `executable_path` in the WebDriver initialization with the path to your Edge WebDriver executable.
2. Set your login credentials (email and password) in the script where indicated.
3. Modify the `categoryType`, `category_id`, `subcategory_id`, `startPage`, and `endPage` variables as needed to target the desired product category and page range.
4. Run the script. The product information and images will be saved in the current working directory.

Note: Ensure that the website's terms of service allow web scraping before running this script.
"""
from selenium import webdriver
import requests
import os
import re
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from urllib.parse import urlparse

# Optional: Set Edge options
edge_options = Options()
edge_options.add_argument("--headless")  # Run in headless mode
edge_options.add_argument("--disable-gpu")  # Disable GPU acceleration
edge_options.add_argument("--window-size=1920x1080")  # Set window size
edge_options.add_argument("--log-level=3")
edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])

# Initialize Edge WebDriver
driver = webdriver.Edge(service=Service(executable_path=r'c:\Users\GurumurthyThonukunoo\Downloads\edgedriver_win64\msedgedriver.exe'), options=edge_options)

categoryType = 2
category_id = 156
subcategory_id = 1749
startPage = 2
endPage = 3

# Use the WebDriver instance
driver.get("https://app.provi.com")

try:
    wb = Workbook()
    ws = wb.active
    headers = ["productName", "region", "productType", "category", "subcategory", "containerType", "productDescription", "probableBrandFamily", "country", "abv", "siteProductId", "seqNum", "vintage", "rawMaterials", "appellation", "producer", "feature", "aboutProducer"]
    ws.append(headers)

    # Workbook and worksheet for failed products
    failed_wb = Workbook()
    failed_ws = failed_wb.active
    failed_ws.append(["productName", "seqNum"])

    folder_name = 'product_images'
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    # Wait until the input field is present and interactable
    input_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@id='user_email']"))
    )
    # Enter text into the input field
    input_element.send_keys("siva@1800spirits.com")
    input_element1 = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@id='user_password']"))
    )
    input_element1.send_keys("Password@01")

    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[text()='Log in']"))
    )
    # Click the button
    login_button.click()

    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//input[@id='retailer_header_product_search_input']"))
    )

    # Navigate to the specified URL with categoryType and product_sort
    url = f"https://app.provi.com/product_listing?type={categoryType}&product_sort=best-seller&c={category_id}&fired_filter=subcategory_id&s={subcategory_id}&page={startPage}"
    driver.get(url)

    def extract_product_number(url):
        """Extract the product number from the URL.
        :param url: str (URL containing the product number)
        :return: str (extracted product number)"""
        parsed_url = urlparse(url)
        path = parsed_url.path
        match = re.search(r'/products/(\d+)', path)
        if match:
            return match.group(1)
        else:
            raise ValueError("Product number not found in the URL.")

    def process_products_on_page(categoryType, category_id, subcategory_id, page_number):
        # Initialize a set to track processed product indices on the current page
        processed_indices = set()

        while True:
            try:
                # Wait until the product list is present
                product_list = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='product-line-list' and @data-testid='product-line-list']"))
                )

                # Find all child divs within the product list
                child_divs = product_list.find_elements(By.XPATH, "./div")

                if not child_divs:
                    print("No more child divs found.")
                    break  # Exit the loop if no more child divs are found

                all_processed = True

                for index in range(len(child_divs)):
                    if index in processed_indices:
                        continue  # Skip already processed indices

                    # Initialize seqNum for each product to ensure it's unique
                    pageNumber_formatted = f"{page_number:04}"
                    pageNumber_formatted1 = f"{(index + 1):02}"
                    seqNum = f"{category_id}-{subcategory_id}-{categoryType}{pageNumber_formatted}{pageNumber_formatted1}"

                    try:
                        # Re-fetch the list of child divs to avoid stale element references
                        product_list = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, "//div[@class='product-line-list' and @data-testid='product-line-list']"))
                        )
                        child_divs = product_list.find_elements(By.XPATH, "./div")

                        # Click on the index-th child div
                        child_div = child_divs[index]
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, ".//a[@id='product-card-name']"))
                        )
                        product_link = child_div.find_element(By.XPATH, ".//a[@id='product-card-name']")
                        product_name = product_link.text  # You can get the product name if needed

                        # Open the product link in a new tab
                        driver.execute_script("window.open(arguments[0]);", product_link.get_attribute('href'))
                        driver.switch_to.window(driver.window_handles[1])

                        print(f"Processing product {index + 1}: {product_name}")

                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, "//div[@data-testid='product-line-box']"))
                        )

                        try:
                            productType = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//div[@data-testid='breadcrumb-navigation']//ol//li[@data-testid='product-type-breadcrumb']"))
                            ).text
                        except:
                            productType = None
                        try:
                            category_text = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//div[@data-testid='breadcrumb-navigation']//ol//li[@data-testid='category-breadcrumb']"))
                            ).text
                        except:
                            category_text = None
                        try:
                            subcategory_text = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//div[@data-testid='breadcrumb-navigation']//ol//li[@data-testid='subcategory-breadcrumb']"))
                            ).text
                        except:
                            subcategory_text = None

                        try:
                            productName = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//h1[@class='fizz-heading-1 notranslate fizz-stack-8']"))
                            ).text
                            productName = productName.replace(',', '|')
                        except:
                            productName = None
                        try:
                            region = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Region:')]/following-sibling::div/a"))
                            ).text
                        except:
                            region = None
                        try:
                            text_large_element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Product information')]/following-sibling::div"))
                            ).text
                            productDescription = text_large_element.replace(',', '|')
                        except:
                            productDescription = None
                        try:
                            country = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Country:')]/following-sibling::div/a"))
                            ).text
                            country = country.replace(',', '|')
                        except:
                            country = None
                        try:
                            abv = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'ABV:')]/following-sibling::div"))
                            ).text
                        except:
                            abv = None
                        try:
                            feature = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Feature')]/following-sibling::a"))
                            ).text
                            feature = feature.replace(',', '|')
                        except:
                            feature = None
                        try:
                            aboutProducer = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'About the producer')]/following-sibling::p"))
                            ).text
                            aboutProducer = aboutProducer.replace(',', '|')
                        except:
                            aboutProducer = None
                        try:
                            rawMaterials1 = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Raw Materials')]"))
                            )

                            # Locate the "show more" link if it exists and click it
                            try:
                                show_more = driver.find_element(By.XPATH, "//a[contains(text(), 'show more...')]")
                                show_more.click()
                            except:
                                pass  # If "show more" is not found, proceed without clicking

                            button_elements1 = rawMaterials1.find_elements(By.XPATH, "./following-sibling::a")

                            texts = [a.text.strip() for a in button_elements1 if 'show less' not in a.text.strip()]

                            formatted_text = '|'.join(texts)

                            print("rawMaterials texts", formatted_text)
                            rawMaterials = formatted_text  # For example, storing in a variable

                        except:
                            rawMaterials = None
                        try:
                            appellation = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Appellation')]/following-sibling::div/a"))
                            ).text
                        except:
                            appellation = None

                        try:
                            text_content = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Producer:')]/ancestor::p"))
                            ).text
                            producer = text_content.split('Producer:')[1].strip()

                        except:
                            producer = None

                        try:
                            probableBrandFamily = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//section[@data-testid='brand_family_best_sellers-carousel']//header//div//h3"))
                            ).text
                            print("probableBrandFamily texts", probableBrandFamily)
                            probableBrandFamily = probableBrandFamily.split('More from ')[1].strip().replace(',', '|')

                        except:
                            probableBrandFamily = None

                        try:
                            vintage_exists = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//p[text()='Vintage']"))
                            )
                            button_section = vintage_exists.find_element(By.XPATH, "./ancestor::div[contains(@class, 'fizz-card')]//div[contains(@class, 'button-group')]")

                            button_elements = button_section.find_elements(By.XPATH, ".//button")

                            button_texts = [button.text.strip() for button in button_elements]
                            formatted_text = '|'.join(button_texts)

                            print("Vintage button texts:", button_texts)

                            vintage = formatted_text
                        except:
                            vintage = None

                        try:
                            containerType_section = driver.find_element(By.XPATH, "//div//ul[@data-testid='container-choices']")

                            containerType_elements = containerType_section.find_elements(By.XPATH, ".//li")

                            container_texts = [li.text.strip() for li in containerType_elements]
                            formatted_text1 = '|'.join(container_texts)

                            print("Container Type texts:", container_texts)

                            containerType = formatted_text1
                        except:
                            containerType = None
                        currentUrl = driver.current_url
                        product_number = extract_product_number(currentUrl)
                        image_element = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//div[@id='product-image']//img[@class='product-image-class fizz-full-width s-aj8RbU3kIi4D']"))
                        )
                        image_url = image_element.get_attribute('src')
                        product_name_filter = re.sub(r'[\/:*?"<>|]', '_', product_name)
                        productPath = product_name_filter + '_' + product_number + '.jpg'
                        # productPath = product_name.replace('/', '_') + '_' + product_number + '.jpg'
                        if image_url.startswith('//'):
                            image_url = 'http:' + image_url

                        print(f"image Url: {image_url}")
                        response = requests.get(image_url)
                        response.raise_for_status()

                        image_filename = os.path.join(folder_name, productPath)
                        with open(image_filename, 'wb') as file:
                            file.write(response.content)
                        siteProductId = product_number

                        print(f'All Fields, productName: {productName}, region: {region}, productDescription: {productDescription}, probableBrandFamily: {probableBrandFamily}, country: {country}, abv: {abv}, siteProductId: {siteProductId}, seqNumber: {seqNum}, appellation: {appellation}, containerType: {containerType}, vintage: {vintage}, producer: {producer}, feature: {feature}, rawMaterials: {rawMaterials}, aboutProducer: {aboutProducer}')

                        ws.append([productName, region, productType, category_text, subcategory_text, containerType, productDescription, probableBrandFamily, country, abv, siteProductId, seqNum, vintage, rawMaterials, appellation, producer, feature, aboutProducer])

                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])

                        processed_indices.add(index)

                        all_processed = False

                    except Exception as e:
                        print(f"An error occurred while processing product {index + 1}: {product_name} : {e}")
                        failed_ws.append([product_name, seqNum])

                if all_processed:
                    print("All products have been processed on this page.")
                    break

            except Exception as e:
                print(f"An error occurred while processing the product list: {e}")
                break

    current_page_number = startPage

    while current_page_number <= endPage:
        process_products_on_page(categoryType, category_id, subcategory_id, current_page_number)

        # Save after processing each page
        wb.save(f"Site1ScreenScrape-{categoryType}-{category_id}-{subcategory_id}-{startPage:04d}-{endPage:04d}.xlsx")
        failed_wb.save(f"Err-Site1ScreenScrape-{categoryType}-{category_id}-{subcategory_id}-{startPage:04d}-{endPage:04d}.xlsx")

        # Check if we have processed the last page
        if current_page_number >= endPage:
            print("Reached the end page.")
            break

        # Navigate to the next page using the "Next" button
        try:
            next_page_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='next-page']"))
            )
            next_page_button.click()

            # Update the current page number
            current_page_number += 1

        except Exception as e:
            print(f"An error occurred while navigating to the next page: {e}")
            break

except Exception as e:
    print(f'An error occurred: {e}')

driver.quit()
