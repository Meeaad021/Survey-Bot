import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
import time
import threading
import random
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import pandas as pd


# Create a global log
log_data = []

def identify_question_type(driver):
    global log_data
    try:
        action = ActionChains(driver)

        # Single-select question (radio buttons)
        if driver.find_elements(By.CSS_SELECTOR, '.cTable.rsSingle'):
            print("Single-select question detected.")
            radio_buttons = driver.find_elements(By.CLASS_NAME, "cRadio")

            already_selected = any(rb.is_selected() for rb in radio_buttons)
            if already_selected:
                print("Radio button already selected.")
            elif radio_buttons:
                select_random = random.choice(radio_buttons)
                question_id = select_random.get_attribute("name") or "unknown"
                answer_value = select_random.get_attribute("value")

                action.move_to_element(select_random).click().perform()
                print(f"Selected radio: QID={question_id}, Value={answer_value}")

                log_data.append([respondent_id, question_id, "Single-select", answer_value])
            else:
                print("No radio buttons found.")

        # Multi-select question (checkboxes)
        elif driver.find_elements(By.CSS_SELECTOR, '.cTable.rsMulti'):
            print("Multi-select question detected.")
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CLASS_NAME, "cCheck")))
            checkboxes = driver.find_elements(By.CLASS_NAME, "cCheck")

            if not checkboxes:
                print("No checkboxes found.")
                return

            int_num_checks = random.randint(1, len(checkboxes))
            selected_checks = random.sample(checkboxes, int_num_checks)
            for cb in selected_checks:
                question_id = cb.get_attribute("name") or "unknown"
                answer_value = cb.get_attribute("value")
                action.click(cb).perform()
                print(f"Selected checkbox: QID={question_id}, Value={answer_value}")
                log_data.append([Respondent_id, question_id, "Multi-select", answer_value])
        
        # If no question types are available but .cSay exists, click 'Next'
        elif driver.find_elements(By.CSS_SELECTOR, '.cSay'):
            print("Info box found.")

            btn_next = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "btnNext"))
                )
            action.click(btn_next).perform()
            print("Moved to the next page.")
              
    except Exception as e:
        print(f"Error during question identification: {e}")

def start_testing():
    try:
        str_base_url = entry_url.get()
        int_num_tests = int(entry_tests.get())

        if not str_base_url or int_num_tests <= 0:
            messagebox.showerror("Error", "Please enter both a valid URL and a positive number of respondents.")
            return

        start_button.config(state="disabled")
        for i in range(1, int_num_tests + 1):
            try:
                entry_url.delete(0, tk.END)
                entry_url.insert(0, str_base_url)  # Reinsert the base URL for clarity
                driver = webdriver.Chrome()

                # Build unique URL
                str_test_url = f"{str_base_url}&altid=myj_trial_27_{i}"
                driver.get(str_test_url)

                while True:

                    try:
                        # Always switch to iframe before interacting
                        WebDriverWait(driver, 10).until(
                            EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe"))
                        )
                        print("Switched to iframe.")
                    except TimeoutException:
                        print("Could not switch to iframe") 

                        if driver.find_elements(By.CLASS_NAME, "home-container.text-center"):
                            print("Test completed.")
                            break

                    # Detect question type and interact
                    identify_question_type(driver)

                    try:
                        # Look for the Next button and click it
                        next_button = WebDriverWait(driver, 7).until(
                            EC.element_to_be_clickable((By.ID, "btnNext"))
                        )
                        driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                        time.sleep(0.5)
                        next_button.click()
                        print(f"Clicked 'Next' button for test {i}.")
                        time.sleep(2)

                        # After clicking Next, switch back to main and then to the new iframe
                        driver.switch_to.default_content()
                    except TimeoutException:
                        print("Next button not found, trying again...")

                        try:
                            if not driver.find_elements(By.TAG_NAME, "iframe"):
                                print("Test completed.")
                            
                                break
                        except TimeoutException:
                                # Not the last page, continue processing
                                pass
                        
            except WebDriverException as e:
                print(f"WebDriver error on test {i}: {e}")
            finally:
                if 'driver' in locals():
                    driver.quit()

        messagebox.showinfo("Completed", f"Tested {int_num_tests} survey links.")
        root.destroy()

    except Exception as e:
        print(f"An error occurred: {e}")
        messagebox.showerror("Unexpected Error", str(e))
    finally:
        start_button.config(state="normal")
        save_logs_to_excel()


def save_logs_to_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Survey Logs"

    headers = ["Respondent ID", "Question ID", "Type", "Selected Value"]
    ws.append(headers)

    for row in log_data:
        ws.append(row)

    wb.save("survey_log.xlsx")
    print("Logs saved to 'survey_log.xlsx'")


# Create the Tkinter GUI
root = tk.Tk()
root.title("Link Testing")

root.update_idletasks()
width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f'+{x}+{y}')

# URL entry
tk.Label(root, text="Survey URL:").grid(row=0, column=0, padx=5, pady=5)
entry_url = tk.Entry(root, width=40)
entry_url.grid(row=0, column=1, padx=5, pady=5)

# Respondents entry
tk.Label(root, text="Number of Respondents:").grid(row=1, column=0, padx=5, pady=5)
entry_tests = tk.Entry(root, width=10)
entry_tests.grid(row=1, column=1, padx=5, pady=5)

# Start testing button
start_button = tk.Button(root, text="Start Testing", command=lambda: threading.Thread(target=start_testing).start())
start_button.grid(row=2, column=0, columnspan=2, pady=10)

# Run the application
root.mainloop()