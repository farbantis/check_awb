import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from tkinter import filedialog, Tk, Text, END, messagebox,Toplevel, Label
import time


def get_header():
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    firefox_options = webdriver.FirefoxOptions()
    firefox_options.add_argument(f'user-agent={user_agent}')
    firefox_options.add_argument('--headless')  # Run in headless mode (no browser window)
    firefox_options.add_argument('--disable-gpu')  # Disable GPU acceleration
    firefox_options.add_argument('--disable-dev-shm-usage')
    firefox_options.add_argument('--no-sandbox')
    return firefox_options


def write_results_to_excel(result_df, file_name):
    file_name = f"{file_name[:-5]}_checked{file_name[-5:]}"
    print(f'saving info to {file_name}')
    try:
        with pd.ExcelWriter(file_name,
                            engine='openpyxl',
                            mode='a',
                            if_sheet_exists='overlay') as writer:

            result_df.to_excel(writer,
                               sheet_name='DPD',
                               index=False,
                               header=False,
                               startrow=writer.sheets['DPD'].max_row)
    except FileNotFoundError:
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            result_df.to_excel(writer, sheet_name='DPD', index=False, header=False)


def get_tracking_info(awb_numbers):
    driver = webdriver.Firefox(options=get_header())
    base_url = f"https://tracktrace.dpd.com.pl/EN/parcelDetails?typ=1"
    url = base_url
    for i, awb_num in enumerate(awb_numbers, start=1):
        url += f"&p{i}={awb_num}"
    try_limit = 2
    try_count = 0
    result = []
    while try_count <= try_limit:
        try:
            driver.get(url)
            time.sleep(1)
            elements = driver.find_elements(By.CLASS_NAME, 'table-track')
            for index, element in enumerate(elements, start=awb_numbers.index[0]):
                if "parcel delivered" in element.text.lower():
                    delivery_status = 'delivered'
                elif 'wrong addresss' in element.text.lower():
                    delivery_status = 'wrong address'
                else:
                    delivery_status = 'check status'
                # print(awb_numbers[index], delivery_status)
                result.append((awb_numbers[index], delivery_status, element.text))
            driver.quit()
            return result
        except:
            if try_count < try_limit:
                try_count += 1
                continue
            else:
                driver.quit()
                return 'error', 'error', 'error'


def get_awb_list(excel_file):
    df = pd.read_excel(excel_file, dtype=str)
    df = df['Primary_Consigment_No'].dropna().reset_index(drop=True)
    count = 1
    batch_size = 10
    for i in range(0, len(df), batch_size):
        if i + batch_size > len(df):
            batch = df[i:]
        else:
            batch = df[i:i+batch_size]
        batch_result = get_tracking_info(batch)
        results = []
        for element in batch_result:
            results.append(element)
            awb, status, details = element
            print(f'{count}) {awb} {status}')
            count += 1
        result_df = pd.DataFrame(results, columns=['AWB_Num', 'Delivery_Status', 'Details'])
        write_results_to_excel(result_df, excel_file)


excel_file_path = filedialog.askopenfilename(filetypes=[("XLSX files", "*.xlsx")])
get_awb_list(excel_file_path)


