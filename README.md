# PayFast Email Parser

This script automates the process of extracting order details from PayFast confirmation emails. It connects to your email account via IMAP, reads unread order notifications, and compiles the data (Date, Product, Customer, Discount, Payment Method) into a clean Excel spreadsheet.

It also includes logic to handle "Bundle" products by filtering out individual sub-components when a bundle is detected.

## Features

* **Automated Fetching:** Connects securely to your email provider using IMAP.
* **Smart Parsing:** Uses BeautifulSoup to scrape HTML email content for specific order tables.
* **Bundle Logic:** Automatically detects bundle products and removes redundant sub-products from the final report.
* **Excel Export:** Generates a timestamped `.xlsx` file with formatted dates and currency values.
* **Status Management:** Marks processed emails as "Read" so they aren't duplicated in future runs.

## Prerequisites

* Python 3.x
* An email account that supports IMAP

## Installation

1.  **Download:**
    * Click the green **<> Code** button at the top of this page.
    * Select **Download ZIP**.
    * Extract the ZIP file to a folder on your computer.

2.  **Install Dependencies:**
    Open your terminal or command prompt in the extracted folder and run:
    ```bash
    pip install -r requirements.txt
    ```

## Configuration

**Crucial Step:** You must add your email credentials before running the script.

1.  Open `imap_script.py` in a text editor (Notepad, VS Code, etc.).
2.  Locate the **CONFIGURATION** section at the top:

    ```python
    # ==============================
    # CONFIGURATION
    # ==============================
    IMAP_SERVER = "imap.yourprovider.com"  # e.g., imap.gmail.com or outlook.office365.com
    EMAIL_ACCOUNT = "orders@yourdomain.com"
    PASSWORD = "your-password-here"
    FOLDER = '"PayFast/Orders"'  # Ensure this matches your actual folder name
    ```
3.  Save the file.

## Usage

1.  **Run the Script:**
    ```bash
    python imap_script.py
    ```

2.  **Output:**
    * The script will print the number of unread emails found.
    * A new folder named `exports` will be created.
    * An Excel file (e.g., `PayfastOrders_2023-10-25_14-30-00.xlsx`) will be saved inside the `exports` folder.
