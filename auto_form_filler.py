import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import multiprocessing
import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException, ElementNotInteractableException
import pandas as pd
from datetime import datetime, timedelta
import os
import uuid # Để tạo ID duy nhất cho mỗi trình duyệt


# --- Cấu hình WebDriver ---
driver_path = 'chromedriver.exe' 

def get_chrome_options():
    options = webdriver.ChromeOptions()
    # Bỏ comment dòng này nếu bạn muốn trình duyệt CHẠY ẨN trong nền (headless mode).
    # options.add_argument("--headless") 
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--log-level=3") 
    # Thêm các tùy chọn khác để tối ưu hiệu suất và giảm tải trình duyệt
    options.add_argument("--disable-extensions") # Tắt tiện ích mở rộng
    options.add_argument("--no-sandbox") # Tắt chế độ sandbox (đôi khi cần thiết)
    # Loại bỏ thanh thông báo "Chrome is being controlled by automated test software"
    options.add_experimental_option("excludeSwitches", ["enable-automation"]) 
    options.add_experimental_option('useAutomationExtension', False)
    return options

# Định nghĩa các heuristics (phán đoán) cho từng trường
FIELD_LOCATORS = {
    "full_name": [
        (By.ID, "fullName"), (By.NAME, "fullName"),
        (By.XPATH, "//input[contains(@placeholder, 'tên') or contains(@placeholder, 'name') or contains(@placeholder, 'full name') or contains(@aria-label, 'name')]"),
        (By.XPATH, "//input[contains(@id, 'name') or contains(@id, 'fullname')]"),
        (By.XPATH, "//input[contains(@name, 'name') or contains(@name, 'fullname')]"),
        (By.XPATH, "//label[contains(text(), 'Họ tên') or contains(text(), 'Full Name')]/following-sibling::input"),
        (By.CSS_SELECTOR, "input[placeholder*='tên'], input[placeholder*='name'], input[placeholder*='full name']"),
        (By.CSS_SELECTOR, "input[id*='name'], input[id*='fullname']"),
        (By.CSS_SELECTOR, "input[name*='name'], input[name*='fullname']"),
    ],
    "email": [
        (By.XPATH, "//input[contains(@type, 'email')]"),
        (By.XPATH, "//input[contains(@placeholder, 'email') or contains(@aria-label, 'email')]"),
        (By.XPATH, "//input[contains(@id, 'email')]"),
        (By.XPATH, "//input[contains(@name, 'email')]"),
        (By.XPATH, "//label[contains(text(), 'Email')]/following-sibling::input"),
        (By.CSS_SELECTOR, "input[type='email']"),
        (By.CSS_SELECTOR, "input[placeholder*='email']"),
        (By.CSS_SELECTOR, "input[id*='email']"),
        (By.CSS_SELECTOR, "input[name*='email']"),
    ],
    "phone_number": [
        (By.XPATH, "//input[contains(@type, 'tel')]"),
        (By.XPATH, "//input[contains(@placeholder, 'thoại') or contains(@placeholder, 'phone') or contains(@placeholder, 'sdt') or contains(@aria-label, 'phone')]"),
        (By.XPATH, "//input[contains(@id, 'phone') or contains(@id, 'tel')]"),
        (By.XPATH, "//input[contains(@name, 'phone') or contains(@name, 'tel')]"),
        (By.XPATH, "//label[contains(text(), 'Số điện thoại') or contains(text(), 'Phone')]/following-sibling::input"),
        (By.CSS_SELECTOR, "input[type='tel']"),
        (By.CSS_SELECTOR, "input[placeholder*='thoại'], input[placeholder*='phone'], input[placeholder*='sdt']"),
        (By.CSS_SELECTOR, "input[id*='phone'], input[id*='tel']"),
        (By.CSS_SELECTOR, "input[name*='phone'], input[name*='tel']"),
    ],
      # Cập nhật locators cho Ngày sinh
    "date_of_birth": [ # Dùng cho ô input ngày sinh duy nhất (dd/mm/yyyy)
        (By.ID, "dateOfBirth"), (By.NAME, "dateOfBirth"),
        (By.XPATH, "//input[contains(@placeholder, 'dd/mm/yyyy')]"),
        (By.XPATH, "//input[contains(@placeholder, 'ngày sinh') or contains(@placeholder, 'date of birth')]"),
        (By.XPATH, "//input[contains(@id, 'dob') or contains(@id, 'birth') or contains(@name, 'dob') or contains(@name, 'birth')]"),
        (By.XPATH, "//label[contains(text(), 'Ngày sinh') or contains(text(), 'Date of birth')]/following-sibling::input"),
        (By.CSS_SELECTOR, "input[placeholder*='dd/mm/yyyy'], input[placeholder*='ngày sinh'], input[placeholder*='date of birth']"),
    ],
    "date_of_birth_day": [ # Dùng cho ô input Ngày riêng biệt
        (By.XPATH, "//input[(@name='day' or contains(@id, 'day')) and (@placeholder='dd' or contains(@class, 'date-day'))]"),
        (By.XPATH, "//input[contains(@id, 'day') and @type='text' and @maxlength='2']"),
        (By.XPATH, "//input[contains(@name, 'day') and @type='text' and @maxlength='2']"),
        (By.XPATH, "//input[@placeholder='dd']"),
        (By.XPATH, "//label[contains(text(), 'Ngày')]/following-sibling::input"),
        (By.XPATH, "//label[contains(text(), 'Day')]/following-sibling::input"),
        (By.XPATH, "(//input[contains(@class, 'date-input') or contains(@class, 'form-control')])[1]"), # Giả định ô ngày là ô đầu tiên trong nhóm
        (By.XPATH, "//div[contains(@class, 'date-picker-group')]//input[1]"),
    ],
    "date_of_birth_month": [ # Dùng cho ô input Tháng riêng biệt
        (By.XPATH, "//input[(@name='month' or contains(@id, 'month')) and (@placeholder='mm' or contains(@class, 'date-month'))]"),
        (By.XPATH, "//input[contains(@id, 'month') and @type='text' and @maxlength='2']"),
        (By.XPATH, "//input[contains(@name, 'month') and @type='text' and @maxlength='2']"),
        (By.XPATH, "//input[@placeholder='mm']"),
        (By.XPATH, "//label[contains(text(), 'Tháng')]/following-sibling::input"),
        (By.XPATH, "//label[contains(text(), 'Month')]/following-sibling::input"),
        (By.XPATH, "(//input[contains(@class, 'date-input') or contains(@class, 'form-control')])[2]"), # Giả định ô tháng là ô thứ hai
        (By.XPATH, "//div[contains(@class, 'date-picker-group')]//input[2]"),
    ],
    "date_of_birth_year": [ # Dùng cho ô input Năm riêng biệt
        (By.XPATH, "//input[(@name='year' or contains(@id, 'year')) and (@placeholder='yyyy' or contains(@class, 'date-year'))]"),
        (By.XPATH, "//input[contains(@id, 'year') and @type='text' and @maxlength='4']"),
        (By.XPATH, "//input[contains(@name, 'year') and @type='text' and @maxlength='4']"),
        (By.XPATH, "//input[@placeholder='yyyy']"),
        (By.XPATH, "//label[contains(text(), 'Năm')]/following-sibling::input"),
        (By.XPATH, "//label[contains(text(), 'Year')]/following-sibling::input"),
        (By.XPATH, "(//input[contains(@class, 'date-input') or contains(@class, 'form-control')])[3]"), # Giả định ô năm là ô thứ ba
        (By.XPATH, "//div[contains(@class, 'date-picker-group')]//input[3]"),
    ],
    "id_card": [
        (By.ID, "idCard"), (By.NAME, "idCard"),
        (By.XPATH, "//input[contains(@placeholder, 'CCCD') or contains(@placeholder, 'ID Card') or contains(@placeholder, 'hộ chiếu') or contains(@placeholder, 'passport')]"),
        (By.XPATH, "//input[contains(@id, 'idcard') or contains(@id, 'passport') or contains(@id, 'identity')]"),
        (By.XPATH, "//input[contains(@name, 'idcard') or contains(@name, 'passport') or contains(@name, 'identity')]"),
        (By.XPATH, "//label[contains(text(), 'CCCD') or contains(text(), 'ID Card') or contains(text(), 'Hộ chiếu')]/following-sibling::input"),
        (By.CSS_SELECTOR, "input[placeholder*='CCCD'], input[placeholder*='ID Card'], input[placeholder*='hộ chiếu'], input[placeholder*='passport']"),
    ],
    # Cập nhật locator cho Sales Date, vì nó là một thẻ <select>
    "sales_date_select": [ 
        (By.NAME, "salesDate"),
        (By.ID, "salesDate"), # Nếu có ID
        (By.XPATH, "//select[contains(@name, 'salesDate')]"),
        (By.XPATH, "//select[contains(@id, 'salesDate')]"),
        (By.XPATH, "//label[contains(text(), 'Sales date') or contains(text(), 'Ngày mua hàng')]/following-sibling::select"),
        (By.CSS_SELECTOR, "select[name='salesDate']"),
    ],
    "session_dropdown": [
        (By.XPATH, "//select[contains(@class, 'form-select')]"),
        (By.XPATH, "//select[contains(@id, 'session')]"),
        (By.XPATH, "//select[contains(@name, 'session')]"),
        (By.XPATH, "//label[contains(text(), 'Session') or contains(text(), 'Phiên')]/following-sibling::select"),
        (By.CSS_SELECTOR, "select.form-select"),
        (By.CSS_SELECTOR, "select[id*='session']"),
    ],
    "submit_button": [
        (By.XPATH, "//button[contains(text(), 'GỬI ĐĂNG KÝ') or contains(text(), 'SUBMIT')]"),
        (By.XPATH, "//input[@type='submit' and (contains(@value, 'Gửi') or contains(@value, 'Submit'))]"),
        (By.CSS_SELECTOR, "button[type='submit']"),
        (By.CSS_SELECTOR, "input[type='submit']"),
    ]
}

def find_element_by_heuristics(driver_instance, field_key):
    locators = FIELD_LOCATORS.get(field_key, [])
    for by_type, locator_value in locators:
        try:
            # Bỏ qua các locator có '{}' vì chúng cần định dạng với giá trị cụ thể
            if '{}' in locator_value:
                continue 

            temp_wait = WebDriverWait(driver_instance, 1) 
            element = temp_wait.until(EC.presence_of_element_located((by_type, locator_value)))
            if element.is_displayed() and element.is_enabled():
                return element
        except TimeoutException:
            continue
        except Exception as e:
            continue
    return None

def fill_and_submit_process(task_data):
    url, driver_path_arg, chrome_options_arg, data, process_id, close_event, active_drivers_map = task_data
    driver = None
    success = False
    driver_id = str(uuid.uuid4()) 
    
    try:
        service = webdriver.chrome.service.Service(driver_path_arg)
        driver = webdriver.Chrome(service=service, options=chrome_options_arg)
        active_drivers_map[driver_id] = True 
        print(f"[Tiến trình {process_id}] WebDriver khởi tạo cho '{data.get('full_name', 'N/A')} - Sales Date: {data.get('sales_date', 'N/A')}', ID: {driver_id}")

        default_session = "10:00 - 12:00"

        driver.get(url)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        # Removed time.sleep(0.5) here as WebDriverWait for body is generally sufficient
        
        print(f"[Tiến trình {process_id}] Đang điền dữ liệu cho: {data.get('full_name', 'N/A')} (Sales Date: {data.get('sales_date', 'N/A')})")

        # --- Điền Full Name ---
        full_name_input = find_element_by_heuristics(driver, "full_name")
        if full_name_input: full_name_input.send_keys(data.get("full_name", ""))
        else: print(f"[Tiến trình {process_id}] Cảnh báo: Không thể điền Họ tên. Bỏ qua.")

        # --- Điền Date of Birth (Xử lý cả ô chung hoặc 3 ô riêng biệt) ---
        dob_data = data.get("date_of_birth", "")
        filled_dob = False 

        dob_input_combined = find_element_by_heuristics(driver, "date_of_birth")
        if dob_input_combined:
            try:
                dob_input_combined.send_keys(dob_data)
                print(f"[Tiến trình {process_id}] Đã điền Ngày sinh vào ô chung: {dob_data}")
                filled_dob = True
            except Exception as e:
                print(f"[Tiến trình {process_id}] Cảnh báo: Lỗi khi điền Ngày sinh vào ô chung: {e}. Thử các ô riêng biệt.")
        
        if not filled_dob:
            day_input = find_element_by_heuristics(driver, "date_of_birth_day")
            month_input = find_element_by_heuristics(driver, "date_of_birth_month")
            year_input = find_element_by_heuristics(driver, "date_of_birth_year")

            if day_input and month_input and year_input:
                try:
                    day, month, year = "", "", ""
                    if dob_data and '/' in dob_data:
                        parts = dob_data.split('/')
                        if len(parts) == 3:
                            day, month, year = parts[0], parts[1], parts[2]
                    
                    if day: day_input.send_keys(day)
                    if month: month_input.send_keys(month)
                    if year: year_input.send_keys(year)
                    print(f"[Tiến trình {process_id}] Đã điền Ngày sinh vào các ô riêng biệt (Ngày: {day}, Tháng: {month}, Năm: {year}).")
                    filled_dob = True
                except Exception as e:
                    print(f"[Tiến trình {process_id}] Cảnh báo: Lỗi khi điền Ngày sinh vào các ô riêng biệt: {e}")
            
            if not filled_dob:
                print(f"[Tiến trình {process_id}] Cảnh báo: Không thể điền Ngày sinh. Không tìm thấy ô chung hoặc bộ ba ô Ngày/Tháng/Năm.")

        # --- Điền Phone Number ---
        phone_input = find_element_by_heuristics(driver, "phone_number")
        if phone_input: phone_input.send_keys(data.get("phone_number", ""))
        else: print(f"[Tiến trình {process_id}] Cảnh báo: Không thể điền Số điện thoại. Bỏ qua.")

        # --- Điền Email ---
        email_input = find_element_by_heuristics(driver, "email")
        if email_input: email_input.send_keys(data.get("email", ""))
        else: print(f"[Tiến trình {process_id}] Cảnh báo: Không thể điền Email. Bỏ qua.")

        # --- Điền ID Card/Passport ---
        id_card_input = find_element_by_heuristics(driver, "id_card")
        if id_card_input: id_card_input.send_keys(str(data.get("id_card", "")))
        else: print(f"[Tiến trình {process_id}] Cảnh báo: Không thể điền CCCD/Hộ chiếu. Bỏ qua.")

        # --- Xử lý Sales date (Cải thiện logic chọn ngày cho <select>) ---
        sales_date_to_fill = data.get("sales_date", datetime.now().strftime("%d/%m/%Y"))
        sales_date_dropdown = find_element_by_heuristics(driver, "sales_date_select") # Sử dụng locator mới
        
        if sales_date_dropdown:
            try:
                select_object = Select(sales_date_dropdown)
                # Thử chọn bằng giá trị (value) trước
                select_object.select_by_value(sales_date_to_fill)
                print(f"[Tiến trình {process_id}] Đã chọn thành công ngày '{sales_date_to_fill}' trong dropdown Sales date.")
                # Removed time.sleep(0.3) here, as selection is usually immediate.
            except NoSuchElementException:
                # Nếu không tìm thấy bằng giá trị, thử bằng văn bản hiển thị
                try:
                    select_object.select_by_visible_text(sales_date_to_fill)
                    print(f"[Tiến trình {process_id}] Đã chọn thành công ngày '{sales_date_to_fill}' trong dropdown Sales date (bằng văn bản hiển thị).")
                    # Removed time.sleep(0.3) here, as selection is usually immediate.
                except NoSuchElementException:
                    print(f"[Tiến trình {process_id}] Cảnh báo: Không tìm thấy tùy chọn ngày '{sales_date_to_fill}' trong dropdown Sales date.")
                except Exception as e_select:
                    print(f"[Tiến trình {process_id}] Lỗi khi chọn ngày trong dropdown Sales date bằng văn bản hiển thị: {e_select}")
            except Exception as e:
                print(f"[Tiến trình {process_id}] Cảnh báo: Lỗi khi tương tác với dropdown Sales date: {e}")
        else: 
            print(f"[Tiến trình {process_id}] Cảnh báo: Không tìm thấy dropdown Sales date. Bỏ qua.")
            
        # --- Xử lý Session (Dropdown) ---
        session_value = data.get("session", default_session)
        session_dropdown = find_element_by_heuristics(driver, "session_dropdown")
        if session_dropdown:
            try:
                select_object = Select(session_dropdown)
                select_object.select_by_visible_text(str(session_value))
            except NoSuchElementException:
                print(f"[Tiến trình {process_id}] Cảnh báo: Giá trị session '{session_value}' không tìm thấy trong dropdown. Thử chọn giá trị đầu tiên.")
                try: select_object.select_by_index(1)
                except NoSuchElementException: print(f"[Tiến trình {process_id}] Cảnh báo: Không có option nào trong dropdown session ngoài giá trị mặc định.")
            except Exception as e: print(f"[Tiến trình {process_id}] Cảnh báo: Lỗi khi tương tác với dropdown session: {e}")
        else: print(f"[Tiến trình {process_id}] Cảnh báo: Không tìm thấy dropdown Session. Bỏ qua.")

        # --- Gửi form ---
        submit_button = find_element_by_heuristics(driver, "submit_button")
        if submit_button:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(submit_button))
            submit_button.click()
            print(f"[Tiến trình {process_id}] Đã gửi thành công dữ liệu cho: {data.get('full_name', 'N/A')} (Sales Date: {data.get('sales_date', 'N/A')})")
            time.sleep(random.uniform(0.5, 2)) 
            success = True
        else:
            print(f"[Tiến trình {process_id}] Cảnh báo: Không tìm thấy nút Gửi. Không thể gửi form.")
            raise NoSuchElementException("Submit button not found.")

    except (TimeoutException, NoSuchElementException) as e:
        error_message = e.msg if e.msg else str(e)
        print(f"[Tiến trình {process_id}] !!! Lỗi tìm kiếm phần tử hoặc timeout cho '{data.get('full_name', 'N/A')} - Sales Date: {data.get('sales_date', 'N/A')}': {error_message}")
        with open("failed_submissions.txt", "a", encoding="utf-8") as f:
            f.write(f"Failed: {data}\nError: {error_message}\n\n")
        success = False
    except Exception as e:
        error_message = str(e)
        print(f"[Tiến trình {process_id}] !!! Lỗi không xác định cho '{data.get('full_name', 'N/A')} - Sales Date: {data.get('sales_date', 'N/A')}': {error_message}")
        with open("failed_submissions.txt", "a", encoding="utf-8") as f:
            f.write(f"Failed: {data}\nError: {error_message}\n\n")
        success = False
    finally:
        if driver:
            print(f"[Tiến trình {process_id}] Hoàn tất xử lý cho '{data.get('full_name', 'N/A')} - Sales Date: {data.get('sales_date', 'N/A')}'.")
            
            wait_time = 30 * 60 
            check_interval = 1 
            start_time = time.time()
            
            print(f"[Tiến trình {process_id}] Trình duyệt sẽ mở trong tối đa 30 phút hoặc cho đến khi có lệnh đóng.")
            while (time.time() - start_time < wait_time) and not close_event.is_set():
                time.sleep(check_interval)
            
            if close_event.is_set():
                print(f"[Tiến trình {process_id}] Nhận lệnh đóng sớm. Đang đóng trình duyệt cho '{data.get('full_name', 'N/A')}'.")
            else:
                print(f"[Tiến trình {process_id}] Hết 30 phút. Đang đóng trình duyệt cho '{data.get('full_name', 'N/A')}'.")
            
            try:
                driver.quit()
                print(f"[Tiến trình {process_id}] Đã đóng trình duyệt cho '{data.get('full_name', 'N/A')}'.")
            except WebDriverException as e:
                print(f"[Tiến trình {process_id}] Cảnh báo: Lỗi khi đóng trình duyệt cho '{data.get('full_name', 'N/A')}': {e}")
            
            if driver_id in active_drivers_map:
                del active_drivers_map[driver_id]
        
        return (success, f"{data.get('full_name', 'N/A')} (Sales Date: {data.get('sales_date', 'N/A')})")


# --- Lớp ứng dụng GUI ---
class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Phần Mềm Điền Form Tự Động")
        self.geometry("700x600") 

        self.manager = multiprocessing.Manager()
        self.close_all_tabs_event = self.manager.Event() 
        self.active_drivers_map = self.manager.dict() 

        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing) 

    def create_widgets(self):
        input_frame = tk.Frame(self, padx=10, pady=10)
        input_frame.pack(pady=10, fill=tk.X)

        tk.Label(input_frame, text="URL Trang Web:").grid(row=0, column=0, sticky="w", pady=5)
        self.url_entry = tk.Entry(input_frame, width=60)
        self.url_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.url_entry.insert(0, "https://testpop-nine.vercel.app/registration") 

        tk.Label(input_frame, text="Số Lượng Tab (tối đa trình duyệt đồng thời):").grid(row=1, column=0, sticky="w", pady=5)
        self.num_tabs_entry = tk.Entry(input_frame, width=60)
        self.num_tabs_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.num_tabs_entry.insert(0, "1") 

        tk.Label(input_frame, text="File Excel Dữ Liệu:").grid(row=2, column=0, sticky="w", pady=5)
        self.excel_path_entry = tk.Entry(input_frame, width=50, state='readonly') 
        self.excel_path_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.browse_button = tk.Button(input_frame, text="Duyệt...", command=self.browse_excel_file)
        self.browse_button.grid(row=2, column=2, padx=5, pady=5)

        tk.Label(input_frame, text="Số Ngày Mua Hàng (từ hôm nay):").grid(row=3, column=0, sticky="w", pady=5)
        self.num_sales_days_entry = tk.Entry(input_frame, width=60)
        self.num_sales_days_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.num_sales_days_entry.insert(0, "1") 

        input_frame.grid_columnconfigure(1, weight=1) 

        button_frame = tk.Frame(self, padx=10, pady=5)
        button_frame.pack(pady=5)

        self.start_button = tk.Button(button_frame, text="Bắt Đầu Điền Form", command=self.start_automation_thread, height=2, font=('Arial', 11, 'bold'))
        self.start_button.pack(side=tk.LEFT, padx=10)

        self.close_all_button = tk.Button(button_frame, text="Đóng Tất Cả Trình Duyệt", command=self.confirm_close_all_tabs, height=2, font=('Arial', 11, 'bold'), state='disabled')
        self.close_all_button.pack(side=tk.LEFT, padx=10)

        tk.Label(self, text="Nhật Ký Hoạt Động:", font=('Arial', 10, 'bold')).pack(pady=5)
        self.log_text = scrolledtext.ScrolledText(self, width=80, height=18, state='disabled', wrap=tk.WORD)
        self.log_text.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

    def log_message(self, message):
        self.log_text.config(state='normal') 
        self.log_text.insert(tk.END, message + "\n") 
        self.log_text.see(tk.END) 
        self.log_text.config(state='disabled') 
        self.update_idletasks() 

    def browse_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Chọn File Excel Dữ Liệu",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path_entry.config(state='normal') 
            self.excel_path_entry.delete(0, tk.END) 
            self.excel_path_entry.insert(0, file_path) 
            self.excel_path_entry.config(state='readonly') 

    def start_automation_thread(self):
        self.start_button.config(state='disabled')
        self.close_all_button.config(state='disabled') 
        self.log_message("Bắt đầu quá trình tự động hóa...")
        
        self.close_all_tabs_event.clear() 
        self.active_drivers_map.clear()

        automation_thread = threading.Thread(target=self.run_automation)
        automation_thread.daemon = True 
        automation_thread.start()

    def run_automation(self):
        target_url = self.url_entry.get()
        num_tabs_str = self.num_tabs_entry.get()
        excel_file_path = self.excel_path_entry.get()
        num_sales_days_str = self.num_sales_days_entry.get() 

        if not target_url:
            messagebox.showerror("Lỗi Đầu Vào", "Vui lòng nhập URL.")
            self.start_button.config(state='normal')
            return
        if not excel_file_path:
            messagebox.showerror("Lỗi Đầu Vào", "Vui lòng chọn file Excel.")
            self.start_button.config(state='normal')
            return
        
        try:
            num_tabs_concurrent = int(num_tabs_str) # Đổi tên biến để rõ ràng hơn
            if num_tabs_concurrent <= 0:
                messagebox.showerror("Lỗi Đầu Vào", "Số lượng tab phải là số nguyên dương.")
                self.start_button.config(state='normal')
                return
        except ValueError:
            messagebox.showerror("Lỗi Đầu Vào", "Số lượng tab không hợp lệ. Vui lòng nhập một số nguyên.")
            self.start_button.config(state='normal')
            return

        try:
            num_sales_days = int(num_sales_days_str)
            if num_sales_days <= 0:
                messagebox.showerror("Lỗi Đầu Vào", "Số ngày mua hàng phải là số nguyên dương.")
                self.start_button.config(state='normal')
                return
        except ValueError:
            messagebox.showerror("Lỗi Đầu Vào", "Số ngày mua hàng không hợp lệ. Vui lòng nhập một số nguyên.")
            self.start_button.config(state='normal')
            return

        self.log_message(f"URL: {target_url}")
        self.log_message(f"Số lượng trình duyệt tối đa đồng thời: {num_tabs_concurrent}")
        self.log_message(f"Đường dẫn File Excel: {excel_file_path}")
        self.log_message(f"Số ngày mua hàng muốn điền (từ hôm nay): {num_sales_days}")

        try:
            df = pd.read_excel(excel_file_path)
            self.log_message(f"Đã đọc {len(df)} dòng dữ liệu từ file Excel.")
            
            if 'phone_number' in df.columns:
                df['phone_number'] = df['phone_number'].astype(str) 
                df['phone_number'] = df['phone_number'].apply(
                    lambda x: '0' + x if x and x.isdigit() and not x.startswith('0') and len(x) in [8, 9] else x
                )
                self.log_message("Đã xử lý định dạng số điện thoại (đảm bảo số 0 ở đầu nếu cần).")
                
            if 'id_card' in df.columns:
                df['id_card'] = df['id_card'].astype(str)
                df['id_card'] = df['id_card'].apply(
                     lambda x: '0' + x if x and x.isdigit() and not x.startswith('0') and len(x) == 11 else x
                )
                self.log_message("Đã đảm bảo cột ID Card/CCCD là định dạng văn bản.")

        except FileNotFoundError:
            messagebox.showerror("Lỗi Đọc File", f"Không tìm thấy file Excel tại đường dẫn '{excel_file_path}'.")
            self.start_button.config(state='normal')
            return
        except Exception as e:
            messagebox.showerror("Lỗi Đọc File", f"Lỗi khi đọc file Excel: {e}")
            self.start_button.config(state='normal')
            return

        sales_dates_to_fill = []
        today = datetime.now()
        for i in range(num_sales_days):
            date_to_add = today + timedelta(days=i)
            sales_dates_to_fill.append(date_to_add.strftime("%d/%m/%Y"))
        
        self.log_message(f"Các ngày mua hàng sẽ được điền: {', '.join(sales_dates_to_fill)}")

        all_tasks_data = []
        for excel_row_data in df.to_dict(orient='records'):
            for s_date in sales_dates_to_fill:
                task_data_copy = excel_row_data.copy()
                task_data_copy['sales_date'] = s_date 
                all_tasks_data.append(task_data_copy)

        num_total_tasks = len(all_tasks_data)
        
        if num_total_tasks == 0:
            self.log_message("Không có dữ liệu để điền hoặc số lượng yêu cầu là 0. Kết thúc.")
            self.start_button.config(state='normal')
            return

        chrome_options = get_chrome_options()
        
        tasks = [(target_url, driver_path, chrome_options, all_tasks_data[i], i+1, self.close_all_tabs_event, self.active_drivers_map) for i in range(num_total_tasks)]

        num_workers = min(num_tabs_concurrent, os.cpu_count() if os.cpu_count() else 4, num_total_tasks)
        self.log_message(f"\n--- Bắt đầu điền {num_total_tasks} form với {num_workers} tiến trình (trình duyệt) đồng thời ---")
        self.log_message(f"LƯU Ý: Mỗi trình duyệt sẽ mở trong tối đa 30 phút hoặc cho đến khi có lệnh đóng.")
        
        self.close_all_button.config(state='normal')

        pool = multiprocessing.Pool(processes=num_workers)
        
        async_results = [pool.apply_async(fill_and_submit_process, (task,)) for task in tasks]

        pool.close() 

        successful_submissions = 0
        failed_submissions = 0

        self.log_message("\n--- Đang chờ các tiến trình điền form hoàn tất và đóng trình duyệt ---")
        for i, res in enumerate(async_results):
            try:
                success, data_name = res.get() 
                if success:
                    successful_submissions += 1
                    self.log_message(f"[{i+1}/{num_total_tasks}] Xử lý thành công cho: {data_name}")
                else:
                    failed_submissions += 1
                    self.log_message(f"[{i+1}/{num_total_tasks}] Xử lý thất bại cho: {data_name}")
            except Exception as e:
                self.log_message(f"[{i+1}/{num_total_tasks}] Lỗi khi nhận kết quả từ tiến trình: {e}")
                failed_submissions += 1

        self.log_message("\n--- Báo cáo tổng kết quá trình ---")
        self.log_message(f"Tổng số form đã yêu cầu xử lý: {num_total_tasks}")
        self.log_message(f"Số form gửi thành công: {successful_submissions}")
        self.log_message(f"Số form gửi thất bại: {failed_submissions}")
        if failed_submissions > 0:
            self.log_message(f"Vui lòng kiểm tra file 'failed_submissions.txt' trong cùng thư mục để xem chi tiết các lỗi.")
        
        pool.join() 
        self.log_message("\n--- Tất cả các trình duyệt đã được đóng và quá trình tự động hóa đã kết thúc hoàn toàn ---")
        
        self.start_button.config(state='normal')
        self.close_all_button.config(state='disabled')

    def confirm_close_all_tabs(self):
        if self.active_drivers_map: 
            response = messagebox.askyesno(
                "Xác Nhận Đóng Tất Cả",
                f"Bạn có chắc chắn muốn đóng tất cả {len(self.active_drivers_map)} trình duyệt đang hoạt động ngay lập tức không? "
                "Quá trình điền form trên các tab đó sẽ bị dừng lại."
            )
            if response:
                self.close_all_tabs()
        else:
            messagebox.showinfo("Thông báo", "Không có trình duyệt nào đang hoạt động để đóng.")
            self.close_all_button.config(state='disabled') 

    def close_all_tabs(self):
        if not self.close_all_tabs_event.is_set():
            self.close_all_tabs_event.set() 
            self.log_message("\n>>> Đã gửi lệnh đóng tất cả các trình duyệt. Các trình duyệt sẽ đóng dần.")
            self.close_all_button.config(state='disabled') 

    def on_closing(self):
        if messagebox.askokcancel("Thoát Ứng Dụng", "Bạn có muốn thoát ứng dụng không? Các trình duyệt đang chạy sẽ bị đóng."):
            self.close_all_tabs() 
            self.manager.shutdown() 
            self.destroy() 

if __name__ == "__main__":
    multiprocessing.freeze_support() 
    app = Application()
    app.mainloop()