# DMart Product Scraper

A Python automation project that scrapes product details from the DMart website using Playwright and exports structured product data into Excel.

---

## Features

* Automated DMart product scraping
* Dynamic category navigation
* Product price extraction
* MRP and discount tracking
* Product size detection
* Excel export automation
* Dynamic scrolling for loading products

---

## Technologies Used

* Python
* Playwright
* OpenPyXL

---

## Data Extracted

The scraper collects:

* Category
* Subcategory
* Product Name
* MRP
* DMart Price
* Discount
* Available Sizes
* Price Per Unit

---

## Project Structure

| File          | Description            |
| ------------- | ---------------------- |
| app.py        | Main scraping script   |
| Products.xlsx | Generated Excel output |
| README.md     | Project documentation  |

---

## Installation

### Clone Repository

```bash
git clone https://github.com/anandpraveen71/dmart-product-scraper.git
```

### Install Dependencies

```bash
pip install playwright openpyxl
```

### Install Playwright Browser

```bash
playwright install
```

---

## Run Project

```bash
python app.py
```

---

## Output

The scraper automatically generates:

```txt
Products.xlsx
```

containing all scraped product data.

---

## Skills Demonstrated

* Web Automation
* Browser Automation
* Dynamic Web Scraping
* Data Extraction
* Excel Automation
* Playwright Automation

---

## Author

Praveen Kumar
