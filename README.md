## Overview
This project is a web scraper built using Selenium and OpenPyXL in Python. It extracts text, images, and links from a list of websites provided in an Excel file and writes the scraped data into a new Excel file.

## Prerequisites
- Python 3.x
- Selenium
- WebDriver Manager for Chrome
- OpenPyXL

## Installation
1. **Clone the repository** or download the script files.
2. **Install the required Python packages**:
   ```
   pip install selenium webdriver-manager openpyxl
   ```

## Usage
1. **Prepare the Input Excel File**:
   - Create an Excel file named `LinksSeleniumWEBSC.xlsx` and place it in the directory `C:\Users\ACER\Desktop\Web Scraper selenium\`.
   - The Excel file should have a worksheet named "Sheet1".
   - List the website URLs starting from the second row (A2, A3, etc.).

2. **Run the Script**:
   - Open a command prompt or terminal.
   - Navigate to the directory where the script is located.
   - Run the script using Python:
     ```
     python main.py
     ```

## Output
- The script will read the URLs from the input Excel file and scrape each website for text, images, and links.
- The scraped data will be written to a new worksheet named "Scraped_Data" in the `link2.xlsx` file located in `C:\Users\ACER\Desktop\Web Scraper selenium\`.

## Script Breakdown

### `scrape_data(url, driver)`
This function:
1. Loads the webpage using the provided URL.
2. Waits until the page is fully loaded.
3. Extracts text, images, and links from the webpage.
4. Returns a dictionary containing the scraped data.

### `main()`
This function:
1. Configures the Selenium WebDriver to run in headless mode.
2. Reads website URLs from the input Excel file.
3. Iterates over the URLs, scraping each one and storing the results.
4. Writes the scraped data to a new worksheet in the output Excel file.

## Error Handling
- The script handles missing or unreadable Excel files gracefully, printing appropriate error messages.
- If scraping a particular URL fails, the error is caught, and the script continues with the next URL.

## Example
Ensure your `LinksSeleniumWEBSC.xlsx` contains:
```
| URL                   |
|-----------------------|
| https://example.com   |
| https://another.com   |
```

After running the script, the output in `link2.xlsx` will look like:
```
| Text                       | Images                       | Links                          |
|----------------------------|------------------------------|--------------------------------|
| Example Domain             | https://example.com/image1   | https://example.com/link1      |
| Another Example            | https://another.com/image1   | https://another.com/link1      |
```
