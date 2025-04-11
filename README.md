# 🧪 Survey Link Automation Tool
A Python-based GUI application that automates testing of online surveys by simulating multiple respondent interactions.

## 🔍 Description
This tool uses Selenium WebDriver to open and interact with survey links, automatically detecting and answering question types such as:

- ✅ Single-select (radio buttons)

- ✅ Multi-select (checkboxes)

- ℹ️ Info-only pages (skips when detected)

A simple Tkinter GUI allows users to input a survey base URL and the number of test respondents. For each test respondent, the tool:

1. Loads the survey link (with a unique respondent ID).

2. Detects question types using CSS selectors.

3. Randomly selects and submits valid responses.

4. Clicks through each page until the survey is completed.

5. Logs all respondent activity (question ID, type, value) to an Excel file (survey_log.xlsx).

## 📁 Key Features
- 🖱️ *Fully automated* survey navigation and interaction

- 📊 *Data logging* into Excel via openpyxl

- 🧠 *Smart detection* of question types

- ⚡ *Multithreaded* execution to keep the UI responsive

- 🔧 *Error handling* for timeouts, missing elements, and iframe switching

## 🛠 Tech Stack

- [Python](https://www.python.org/)
- [Tkinter](https://docs.python.org/3/library/tkinter.html) – GUI
- [Selenium](https://www.selenium.dev/) – Browser automation
- [OpenPyXL](https://openpyxl.readthedocs.io/) – Excel export
- [Pandas](https://pandas.pydata.org/) *(optional, for future expansion)*

## 🚀 Getting Started

### Prerequisites

- Python 3.x
- Google Chrome
- ChromeDriver (make sure it's added to your PATH)

### Installation

```bash
pip install selenium openpyxl pandas


