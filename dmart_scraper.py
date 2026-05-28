import time
import openpyxl
from playwright.sync_api import sync_playwright

# Initialize Excel workbook and sheet
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Products"
ws.append(["Category", "Subcategory", "Product Name", "MRP", "DMart Price", "Discount", "Available Sizes", "Price Per Unit"])

# Define a function to extract product details for each product
def extract_product_details(product, category, subcategory):
    try:
        # Product Name Extraction
        product_name = product.query_selector("div.vertical-card_title__pMGg9").text_content().strip()
        print(f"Found product: {product_name}")

        # MRP Extraction
        mrp_section = product.query_selector("section.vertical-card_price-container__tPCU9.vertical-card_strike-through__rRL1B")
        if mrp_section:
            mrp_value = mrp_section.query_selector("p.vertical-card_value__2EBnX span.vertical-card_amount__80Zwk")
            if mrp_value:
                mrp = mrp_value.text_content().strip()
                print(f"MRP found: {mrp}")
            else:
                mrp = "Not available"
        else:
            mrp = "Not available"

        # DMart Price Extraction
        dmart_price_section = product.query_selector("div.vertical-card_price-left__1ecs8 section.vertical-card_price-container__tPCU9")
        if dmart_price_section:
            dmart_price_value = dmart_price_section.query_selector("p.vertical-card_value__2EBnX span.vertical-card_amount__80Zwk")
            if dmart_price_value:
                dmart_price = dmart_price_value.text_content().strip()
                print(f"DMart Price found: {dmart_price}")
            else:
                dmart_price = "Not available"
        else:
            dmart_price = "Not available"

        # Discount Extraction
        discount_section = product.query_selector("section.vertical-card_section-right__4rjsN section.vertical-card_price-container__tPCU9")
        if discount_section:
            discount_value = discount_section.query_selector("p.vertical-card_value__2EBnX span.vertical-card_amount__80Zwk")
            if discount_value:
                discount = discount_value.text_content().strip()
                print(f"Discount found: {discount}")
            else:
                discount = "No discount"
        else:
            discount = "No discount"

        # Size and Price per Unit Extraction
        size_section = product.query_selector("div.bootstrap-select_option__SB_Xy")
        if size_section:
            size_text = size_section.query_selector("span").text_content().strip()
            price_per_unit = size_section.query_selector("span.bootstrap-select_infoTxt-value__kT4zZ")
            price_per_unit_text = price_per_unit.text_content().strip() if price_per_unit else "Not available"
            sizes_str = f"{size_text} - {price_per_unit_text}"
        else:
            sizes_str = "Not available"

        # Append product details to the Excel sheet
        ws.append([category, subcategory, product_name, mrp, dmart_price, discount, sizes_str])

        # Log for debugging
        print(f"Category: {category}, Subcategory: {subcategory}, MRP: {mrp}, DMart Price: {dmart_price}, Discount: {discount}, Sizes: {sizes_str}")

    except Exception as e:
        print(f"Error processing product '{product_name}': {e}")

# Start Playwright and perform the web scraping
with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)  # Set to True to run in headless mode
    page = browser.new_page()

    try:
        # Go to the dmart.in website
        page.goto("https://www.dmart.in")
        time.sleep(5)

        # Handle the pincode input (Replace with your own pincode if needed)
        pincode_input = page.query_selector("#pincodeInput")
        if pincode_input:
            pincode_input.fill("560037")
            pincode_input.press("Enter")
            print("Pincode entered.")
            time.sleep(5)

        # Select store and confirm location
        store_buttons = page.query_selector_all("div.pincode-widget_pincode-body__g684i button")
        if store_buttons:
            store_buttons[0].click()
            time.sleep(5)

            confirm_button = page.query_selector("button:has-text('CONFIRM LOCATION')")
            if confirm_button:
                confirm_button.click()
                time.sleep(5)
        else:
            print("No store buttons found.")

        # Open categories and navigate to 'Grocery'
        all_categories_button = page.query_selector("span.categories-header_listStaticItemLink__nv212:has-text('All Categories')")
        if all_categories_button:
            all_categories_button.click()
            time.sleep(3)

        # List of categories and subcategories to process
        categories_and_subcategories = [
            # Grocery categories
            ("Grocery", [
                ("Dals", 47),
                ("Pulses", 49),
                ("Dry Fruits", 131),
                ("DMart Grocery", 83),
                ("Cooking Oil", 68),
                ("Ghee & Vanaspati", 28),
                ("Flours & Grains", 74),
                ("Rice & Rice Products", 68),
                ("Masala & Spices", 207),
                ("Salt / Sugar / Jaggery", 53)
            ]),
            # Dairy & Beverages categories
            ("Dairy & Beverages", [
                ("Beverages", 256),
                ("Dairy", 137)
            ])
        ]

        # Iterate through each category and subcategory
        for category, subcategories in categories_and_subcategories:
            print(f"Processing {category}...")

            # Open the category in the menu
            category_button = page.query_selector(f"p:has-text('{category}')")
            if category_button:
                category_button.click()
                time.sleep(5)

            # Iterate through subcategories
            for subcategory, _ in subcategories:
                print(f"Processing {subcategory}...")

                # Click on the subcategory
                subcategory_button = page.query_selector(f"p:has-text('{subcategory}')")
                if subcategory_button:
                    subcategory_button.click()
                    time.sleep(5)

                    # Scroll down to load all products
                    scroll_count = 0
                    while scroll_count < 10:  # Adjust the count based on the number of products
                        page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                        time.sleep(2)  # Wait for new products to load
                        scroll_count += 1

                    # Extract products for the current subcategory
                    category_products = page.query_selector_all("div.vertical-card_card-vertical__Q8seS")
                    print(f"Found {len(category_products)} products in {subcategory}.")
                    for product in category_products:
                        extract_product_details(product, category, subcategory)

        # Save the data to a single Excel file after all categories have been processed
        wb.save("Products.xlsx")
        print("All product data saved to Products.xlsx.")

    except Exception as e:
        print(f"Error occurred: {e}")

    finally:
        browser.close()
        print("Browser closed.")
