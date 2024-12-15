# 🚀 **Web Scraper with Selenium & OpenPyXL** 🌐

**A Python-based web scraper** that uses **Selenium** and **OpenPyXL** to extract data (text, images, and links) from websites listed in an Excel file. The scraper automatically reads the URLs, scrapes the content, and stores it in a new Excel file.

---

## 🛠 **Prerequisites**

Make sure you have the following installed:

- Python 3.x
- **Selenium** for browser automation
- **WebDriver Manager** for managing the Chrome WebDriver
- **OpenPyXL** for handling Excel files

### Install the necessary Python packages:

```bash
pip install selenium webdriver-manager openpyxl
```

---

## 🔧 **Setup & Installation**

### 1. **Clone or Download the Repository**:
Clone the repository or download the script files to your local machine.

### 2. **Prepare the Input Excel File**:
Create an Excel file named `LinksSeleniumWEBSC.xlsx` and place it in your working directory, e.g., `C:\Users\ACER\Desktop\Web Scraper selenium\`. 

- The file should contain a sheet named **"Sheet1"**.
- List the website URLs starting from row **A2** (A3, A4, etc.).

---

## 🚀 **How to Use the Scraper**

1. **Run the Script**:
   - Open a terminal or command prompt.
   - Navigate to the folder containing the script.
   - Execute the script using Python:

   ```bash
   python main.py
   ```

2. **What Happens Next**:
   - The script will read all the URLs from the input file (`LinksSeleniumWEBSC.xlsx`).
   - It will scrape the **text**, **images**, and **links** from each website.
   - A new Excel file called `link2.xlsx` will be generated with the scraped data.

---

## 📊 **Output Example**

### Input Excel (`LinksSeleniumWEBSC.xlsx`):

| **URL**                |
|------------------------|
| https://example.com    |
| https://another.com    |

### Output Excel (`link2.xlsx`):

| **Text**               | **Images**                  | **Links**                     |
|------------------------|-----------------------------|-------------------------------|
| Example Domain         | ![Image](https://example.com/image1) | [example link](https://example.com/link1) |
| Another Example        | ![Image](https://another.com/image1) | [another link](https://another.com/link1) |

The data is organized into three columns:
- **Text**: The extracted text from the webpage.
- **Images**: Links to all the images found on the page.
- **Links**: All internal and external links found on the page.

---

## 🧑‍💻 **How the Script Works**

### `scrape_data(url, driver)`
This function:
1. Loads the provided URL using Selenium.
2. Waits for the page to load completely.
3. Scrapes the text, images, and links from the page.
4. Returns a dictionary with the scraped data.

### `main()`
This function:
1. Configures the Selenium WebDriver to run in **headless mode** (no UI).
2. Reads URLs from the input Excel file.
3. Scrapes data from each URL and stores the results.
4. Saves the data into a new Excel file (`link2.xlsx`).

---

## ⚠️ **Error Handling**

- The script handles errors related to missing or unreadable Excel files, and outputs error messages if necessary.
- If scraping a particular URL fails, the script will continue with the next URL, ensuring the process doesn't halt.

---

## 💡 **Example Input File**

Ensure that the **LinksSeleniumWEBSC.xlsx** file has a structure like this:

| **URL**                |
|------------------------|
| https://example.com    |
| https://another.com    |

---

## 📝 **Additional Notes**

- **Dynamic Websites**: The scraper is capable of handling basic websites. If a website loads content dynamically with JavaScript, the scraper still captures the content rendered by Selenium.
- **Customization**: Feel free to extend the functionality to scrape additional data such as headers, tables, or specific sections from a webpage.

---

## 👐 **Contribute or Ask Questions**

- **Contributing**: Feel free to fork this project, submit pull requests, or open issues if you have any suggestions for improvement or bug reports.
- **Questions?**: If you encounter any problems or have any questions, don't hesitate to reach out!

---
