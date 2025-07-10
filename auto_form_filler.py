import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from bs4 import BeautifulSoup
import pandas as pd
import multiprocessing
import threading
import time
import random
import uuid
import re
import json
from selenium import webdriver
from fuzzywuzzy import fuzz
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    WebDriverException,
    InvalidElementStateException,
)
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain.prompts import PromptTemplate
import os

# Cấu hình Gemini API Key (thay thế bằng API của bạn)
GOOGLE_API_KEY = "AIzaSyA4xzYIVE6GWnKCrMIcY55MdLpS9eWzGRU"  # Thay bằng API key thực tế


def get_chrome_options(headless=True):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")  # Chạy ẩn nếu headless là True
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--log-level=3")
    options.add_argument("--disable-extensions")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_argument("--disable-javascript-animations")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    return options


def format_date(date_str):
    """Tách ngày tháng năm từ nhiều định dạng ngày hoặc trả về nguyên gốc"""
    from datetime import datetime

    try:
        if isinstance(date_str, pd.Timestamp):
            date_str = date_str.strftime("%d/%m/%Y")
        elif isinstance(date_str, str):
            for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    return {
                        "day": parsed_date.strftime("%d"),
                        "month": parsed_date.strftime("%m"),
                        "year": parsed_date.strftime("%Y"),
                        "full": parsed_date.strftime("%d/%m/%Y"),
                    }
                except ValueError:
                    continue
        return {"day": "", "month": "", "year": "", "full": date_str}
    except Exception as e:
        print(f"Lỗi khi tách ngày tháng: {e}")
        return {"day": "", "month": "", "year": "", "full": date_str}


def get_sales_dates(driver, form_mapping):
    """Lấy tất cả giá trị từ dropdown sales_date"""
    sales_dates = []
    sales_date_element = find_element_by_heuristics(driver, "sales_date", form_mapping)

    if sales_date_element and sales_date_element.tag_name == "select":
        try:
            select = Select(sales_date_element)
            sales_dates = [
                option.text.strip()
                for option in select.options
                if option.text.strip() and option.text != "-- Chọn ngày --"
            ]
            return sales_dates
        except Exception as e:
            print(f"Lỗi khi lấy danh sách sales_date: {e}")
    else:
        print("Không tìm thấy dropdown sales_date")
    return []


def preprocess_html(html_content):
    """Chuẩn hóa HTML để tăng tính nhất quán cho phân tích"""
    soup = BeautifulSoup(html_content, "html.parser")
    # Chuẩn hóa các element input/select không có name
    labels = soup.find_all("label")
    for label in labels:
        for_id = label.get("for")
        if for_id:
            input_elem = soup.find("input", id=for_id) or soup.find("select", id=for_id)
            if input_elem and not input_elem.get("name"):
                # Tạo name từ label text
                name = (
                    label.text.lower()
                    .strip()
                    .replace(" ", "_")
                    .replace(":", "")
                    .replace("(", "")
                    .replace(")", "")
                )
                input_elem["name"] = name
        else:
            # Nếu label không có for, tìm input/select gần nhất
            input_elem = (
                label.find_next("input")
                or label.find_next("select")
                or label.find("input")
                or label.find("select")
            )
            if input_elem and not input_elem.get("name"):
                name = (
                    label.text.lower()
                    .strip()
                    .replace(" ", "_")
                    .replace(":", "")
                    .replace("(", "")
                    .replace(")", "")
                )
                input_elem["name"] = name
    # Đảm bảo mọi input/select có name hoặc id
    for elem in soup.find_all(["input", "select"]):
        if not elem.get("name") and not elem.get("id"):
            elem["name"] = f"unnamed_{str(uuid.uuid4())[:8]}"
    return str(soup)


def handle_calendar_date(driver, locator_info, date_parts):
    """Xử lý trường date_of_birth kiểu calendar"""
    try:
        by_type = getattr(By, locator_info["by"])
        calendar_input = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((by_type, locator_info["value"]))
        )
        if calendar_input.get_attribute("disabled"):
            print("Trường calendar bị vô hiệu hóa")
            return False
        driver.execute_script("arguments[0].scrollIntoView(true);", calendar_input)
        calendar_input.click()
        driver.execute_script(
            f"arguments[0].value = '{date_parts['full'].replace('/', '-')}'; arguments[0].dispatchEvent(new Event('input'));",
            calendar_input,
        )
        print(f"Đã điền ngày sinh: {date_parts['full']}")
        return True
    except (
        TimeoutException,
        NoSuchElementException,
        InvalidElementStateException,
    ) as e:
        print(f"Lỗi khi điền calendar: {e}")
        try:
            calendar_input.clear()
            calendar_input.send_keys(date_parts["full"])
            print(f"Đã thử điền bằng send_keys: {date_parts['full']}")
            return True
        except Exception as e2:
            print(f"Lỗi khi thử send_keys: {e2}")
            return False


def analyze_html_and_map_columns(
    driver, html_content, excel_columns, excel_data_sample
):
    """
    Phân tích HTML bằng AI để ánh xạ cột Excel và lấy các giá trị hợp lệ từ dropdown.
    Hàm này được cải tiến để tập trung vào việc hiểu ngữ nghĩa của form.
    """
    llm = ChatGoogleGenerativeAI(
        model="gemini-2.0-flash", google_api_key=GOOGLE_API_KEY
    )
    html_content = preprocess_html(html_content)

    # Lấy danh sách option từ dropdown sales_date nếu có
    sales_date_options = []
    try:
        sales_date_element = find_element_by_heuristics(driver, "sales_date", {})
        if sales_date_element and sales_date_element.tag_name == "select":
            select = Select(sales_date_element)
            sales_date_options = [
                option.text.strip() for option in select.options if option.text.strip()
            ]
    except Exception as e:
        print(f"Lỗi khi lấy option từ dropdown sales_date: {e}")

    # Prompt được thiết kế lại để tập trung vào việc hiểu ngữ nghĩa
    prompt_template = PromptTemplate(
        input_variables=[
            "html",
            "excel_columns",
            "excel_data_sample",
            "sales_date_options",
        ],
        template="""\
        **Mục tiêu chính**: Hoạt động như một công cụ ánh xạ thông minh, chuyển đổi HTML của form và dữ liệu Excel thành một cấu trúc JSON để tự động điền form. Phải ưu tiên **hiểu ngữ nghĩa** hơn là chỉ khớp các thuộc tính tĩnh.

        **Bối cảnh**:
        - **HTML**: Mã nguồn HTML của trang web chứa form. HTML có thể phức tạp và không nhất quán.
        - **Excel Columns**: Tên các cột từ file Excel.
        - **Excel Data Sample**: Ba dòng dữ liệu đầu tiên từ Excel để giúp bạn hiểu ý nghĩa của từng cột (ví dụ: cột chứa "Nguyễn Văn A" là tên).
        - **Sales Date Dropdown Options**: Các giá trị `text` từ các `option` trong dropdown `sales_date` (nếu có).

        **Nhiệm vụ**:

        1.  **Phân tích và Ánh xạ Form (`form_mapping`)**:
            - **HIỂU NGỮ NGHĨA**: Phân tích sâu sắc từng element (`input`, `select`, `button`). Đừng chỉ dựa vào `id` hoặc `name`. Hãy xem xét tất cả các gợi ý:
                - `label` đi kèm: Đây là manh mối quan trọng nhất. Ví dụ, `<label>Họ và tên:</label><input name="param1">` rõ ràng là trường `full_name`.
                - `placeholder`: Ví dụ, `placeholder="Nhập email của bạn"`.
                - `name`, `id`, `class`: Sử dụng chúng nhưng phải linh hoạt. `fullName`, `user_name`, `Name` đều có thể là `full_name`.
                - Văn bản xung quanh: Chữ trong thẻ `p`, `span`, `div` gần element cũng chứa thông tin.
            - **TẠO LOCATOR BỀN VỮNG**: Với mỗi trường đã xác định, hãy tạo một locator (bộ định vị) để Selenium có thể tìm thấy nó. Ưu tiên theo thứ tự: ID > NAME > CSS_SELECTOR > XPATH.
            - **CHUẨN HÓA TÊN TRƯỜNG**: Ánh xạ các biến thể về một tên chuẩn. Ví dụ:
                - "Họ tên", "Họ và tên", "Full Name", "Name" -> `full_name`
                - "Email", "Thư điện tử" -> `email`
                - "Số điện thoại", "SĐT", "Phone" -> `phone_number`
                - "CCCD", "CMND", "ID Card", "Hộ chiếu" -> `id_card`
                - "Ngày sinh", "Date of Birth", "DOB" -> `date_of_birth` (Nếu là một trường duy nhất). Nếu chia nhỏ, hãy xác định các trường `day`, `month`, `year`.
                - Nút bấm chính (Gửi, Đăng ký, Submit) -> `submit_button`
            - **ÁNH XẠ CỘT EXCEL**: Trong mỗi mục của `form_mapping`, đặt giá trị của `excel_column` là tên cột tương ứng từ file Excel. Nếu một trường form (như `sales_date`, `session`, `submit_button`) không có dữ liệu từ Excel, đặt `excel_column: null`.

        2.  **Ánh xạ Cột (`column_mapping`)**:
            - Tạo một đối tượng map giữa tên cột gốc trong Excel và tên trường đã chuẩn hóa. Ví dụ: `{{"Họ Tên Khách Hàng": "full_name"}}`.

        3.  **Lọc Ngày Hợp Lệ (`valid_sales_dates`)**:
            - Từ `Sales Date Dropdown Options` được cung cấp, chỉ trích xuất các chuỗi là ngày tháng hợp lệ (ví dụ: `dd/mm/yyyy`, `yyyy-mm-dd`).
            - Loại bỏ hoàn toàn các giá trị không phải là ngày như "-- Chọn ngày --", "Select Date", v.v.

        **HTML**:
        {html}

        **Excel Columns**:
        {excel_columns}

        **Excel Data Sample (ba hàng đầu tiên)**:
        {excel_data_sample}

        **Sales Date Dropdown Options**:
        {sales_date_options}

        **Định dạng đầu ra (chỉ trả về JSON)**:
        Ví dụ:
        {{
            "form_mapping": {{
                "full_name": {{"by": "NAME", "value": "Name", "excel_column": "full_names"}},
                "email": {{"by": "CSS_SELECTOR", "value": "input[type='email']", "excel_column": "email"}},
                "day": {{"by": "NAME", "value": "day", "excel_column": "date_of_birth"}},
                "month": {{"by": "NAME", "value": "month", "excel_column": "date_of_birth"}},
                "year": {{"by": "NAME", "value": "year", "excel_column": "date_of_birth"}},
                "phone_number": {{"by": "NAME", "value": "phoneNumber", "excel_column": "phone_number"}},
                "id_card": {{"by": "NAME", "value": "idCard", "excel_column": "id_card"}},
                "sales_date": {{"by": "NAME", "value": "salesDate", "excel_column": null}},
                "session": {{"by": "NAME", "value": "session", "excel_column": null}},
                "submit_button": {{"by": "CSS_SELECTOR", "value": ".submit-btn", "excel_column": null}}
            }},
            "column_mapping": {{
                "full_names": "full_name",
                "email": "email",
                "phone_number": "phone_number",
                "id_card": "id_card",
                "date_of_birth": "date_of_birth"
            }},
            "valid_sales_dates": ["09/07/2025", "10/07/2025", "11/07/2025"]
        }}
        """,
    )
    chain = prompt_template | llm
    response = chain.invoke(
        {
            "html": html_content,
            "excel_columns": json.dumps(excel_columns),
            "excel_data_sample": json.dumps(excel_data_sample),
            "sales_date_options": json.dumps(sales_date_options),
        }
    )
    try:
        # Dọn dẹp output từ AI để đảm bảo là JSON hợp lệ
        cleaned_content = re.sub(r"```json\n|\n```", "", response.content).strip()
        result = json.loads(cleaned_content)
        print(f"Phân tích form thành công: {json.dumps(result, indent=2)}")
        return (
            result.get("form_mapping", {}),
            result.get("column_mapping", {}),
            result.get("valid_sales_dates", []),
        )
    except json.JSONDecodeError as e:
        print(f"Lỗi khi parse JSON từ AI: {e}")
        return {}, {}, []


def find_element_by_heuristics(driver, form_col, form_mapping):
    """
    Tìm element bằng nhiều tiêu chí (name, id, class, label, placeholder) với fuzzy matching.
    """
    if form_col not in form_mapping:
        return None

    mapping = form_mapping[form_col]
    by = mapping["by"]
    value = mapping["value"]

    # Thử tìm bằng tiêu chí chính từ form_mapping
    try:
        if by == "NAME":
            return WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.NAME, value))
            )
        elif by == "ID":
            return WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, value))
            )
        elif by == "CSS_SELECTOR":
            return WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, value))
            )
    except (TimeoutException, NoSuchElementException):
        print(f"[{form_col}] Không tìm thấy element bằng {by}={value}")

    # Thử tìm bằng các tiêu chí khác với fuzzy matching
    heuristics = [
        (By.NAME, form_col, 90),
        (By.ID, form_col, 90),
        (
            By.XPATH,
            f"//input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{form_col.lower()}')]",
            80,
        ),
        (
            By.XPATH,
            f"//input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{form_col.lower()}')]",
            80,
        ),
        (
            By.XPATH,
            f"//input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{form_col.lower()}')]",
            80,
        ),
        (
            By.XPATH,
            f"//select[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{form_col.lower()}')]",
            80,
        ),
        (
            By.XPATH,
            f"//select[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{form_col.lower()}')]",
            80,
        ),
        (
            By.XPATH,
            f"//label[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{form_col.lower()}')]/following::input[1]",
            80,
        ),
        (
            By.XPATH,
            f"//label[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{form_col.lower()}')]/following::select[1]",
            80,
        ),
    ]

    for by, value, threshold in heuristics:
        try:
            elements = driver.find_elements(by, value)
            for element in elements:
                if by == By.NAME or by == By.ID:
                    if (
                        fuzz.ratio(
                            value.lower(), element.get_attribute(by.lower()).lower()
                        )
                        >= threshold
                    ):
                        return element
                else:
                    return element
        except (TimeoutException, NoSuchElementException):
            continue

    print(f"[{form_col}] Không tìm thấy element bằng bất kỳ tiêu chí nào")
    return None


# Dán và thay thế toàn bộ hàm fill_and_submit_process trong file của bạn
# Dán và thay thế toàn bộ hàm fill_and_submit_process trong file của bạn
def fill_and_submit_process(task_data):
    (
        url,
        chrome_options,
        data,
        process_id,
        close_event,
        active_drivers_map,
        form_mapping,
        column_mapping,
        valid_sales_dates,
    ) = task_data
    driver = None
    success = False
    driver_id = str(uuid.uuid4())
    mapped_data = {"full_name": "N/A"}

    try:
        service = webdriver.chrome.service.Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        active_drivers_map[driver_id] = True

        driver.get(url)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "form"))
        )
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script("return document.readyState === 'complete'")
        )

        initial_url = driver.current_url
        print(f"[{process_id}] URL ban đầu: {initial_url}")

        mapped_data = {}
        for excel_col, form_field_key in column_mapping.items():
            if excel_col in data:
                if form_field_key == "date_of_birth":
                    date_parts = format_date(data.get(excel_col, ""))
                    if "day" in form_mapping:
                        mapped_data["day"] = date_parts["day"].lstrip("0")
                    if "month" in form_mapping:
                        mapped_data["month"] = date_parts["month"].lstrip("0")
                    if "year" in form_mapping:
                        mapped_data["year"] = date_parts["year"]
                    if "date_of_birth" in form_mapping:
                        mapped_data["date_of_birth"] = date_parts["full"]
                else:
                    mapped_data[form_field_key] = str(data.get(excel_col, "")).strip()

        mapped_data["sales_date"] = data.get(
            "sales_date", random.choice(valid_sales_dates) if valid_sales_dates else ""
        )
        mapped_data["session"] = data.get("session", "")
        # LƯU Ý: Sửa lại giá trị session nếu cần thiết để khớp với thuộc tính 'value' trong HTML
        # Ví dụ: nếu text là "13:30 - 15:30" thì value là "12:00 - 14:00"
        if mapped_data["session"] == "13:30 - 15:30":
            mapped_data["session"] = "12:00 - 14:00"

        print(f"[{process_id}] Dữ liệu đã ánh xạ: {mapped_data}")

        priority_fields = ["sales_date", "session"]
        fill_order = priority_fields + [
            k for k in mapped_data.keys() if k not in priority_fields
        ]

        for form_col in fill_order:
            if not mapped_data.get(form_col):
                print(f"[{process_id}] Giá trị trống cho trường {form_col}, bỏ qua")
                continue

            value = mapped_data[form_col]
            element = find_element_by_heuristics(driver, form_col, form_mapping)
            if element:
                try:
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                        element,
                    )
                    time.sleep(0.2)

                    # === GIẢI PHÁP CUỐI CÙNG: SỬ DỤNG JAVASCRIPT ===
                    if element.tag_name == "select":
                        print(
                            f"[{process_id}] DEBUG: Đang xử lý trường SELECT: '{form_col}'"
                        )
                        print(
                            f"[{process_id}] DEBUG: Giá trị được truyền vào JS: '{value}'"
                        )
                        try:
                            js_code = f"""
                                var selectElement = arguments[0];
                                var valueToSelect = arguments[1];
                                console.log('Selenium JS: Đang cố gắng đặt giá trị cho select:', selectElement.name, 'với giá trị:', valueToSelect);
                                selectElement.value = valueToSelect;
                                var event = new Event('change', {{ 'bubbles': true }});
                                selectElement.dispatchEvent(event);
                                console.log('Selenium JS: Đã đặt giá trị và kích hoạt sự kiện change.');
                            """
                            driver.execute_script(js_code, element, value)
                            print(
                                f"[{process_id}] Đã đặt giá trị '{value}' cho dropdown '{form_col}' bằng JavaScript"
                            )
                        except Exception as e:
                            print(
                                f"[{process_id}] LỖI khi đặt giá trị dropdown bằng JS cho '{form_col}': {e}"
                            )
                            return False, mapped_data.get("full_name", "N/A")
                    # === KẾT THÚC PHẦN SỬA ĐỔI ===
                    else:  # Xử lý các trường input như cũ
                        driver.execute_script(
                            f"""
                            var elem = arguments[0];
                            elem.value = '{value}';
                            elem.dispatchEvent(new Event('input', {{ bubbles: true }}));
                            elem.dispatchEvent(new Event('change', {{ bubbles: true }}));
                            elem.dispatchEvent(new Event('blur', {{ bubbles: true }}));
                            """,
                            element,
                        )
                        print(f"[{process_id}] Đã điền '{value}' vào '{form_col}'")

                    time.sleep(0.2)

                except Exception as e:
                    print(f"[{process_id}] Lỗi khi điền '{form_col}': {e}")
                    with open("failed_submissions.txt", "a", encoding="utf-8") as f:
                        f.write(
                            f"Failed to fill '{form_col}' for {mapped_data.get('full_name', 'N/A')}: {e}\n"
                        )
                    return False, mapped_data.get("full_name", "N/A")
            else:
                print(f"[{process_id}] Lỗi: Không tìm thấy trường '{form_col}'")
                with open("failed_submissions.txt", "a", encoding="utf-8") as f:
                    f.write(
                        f"Field not found: '{form_col}' for {mapped_data.get('full_name', 'N/A')}\n"
                    )
                return False, mapped_data.get("full_name", "N/A")

        # Log dữ liệu cuối cùng để kiểm tra
        final_form_data = driver.execute_script(
            "return Object.fromEntries(new FormData(document.querySelector('form')))"
        )
        print(
            f"[{process_id}] Dữ liệu form cuối cùng trước khi submit: {final_form_data}"
        )

        submit_element = find_element_by_heuristics(
            driver, "submit_button", form_mapping
        )
        if submit_element:
            driver.execute_script("arguments[0].click();", submit_element)
            print(
                f"[{process_id}] Đã nhấn nút submit cho: {mapped_data.get('full_name', 'N/A')}"
            )
        else:
            print(f"[{process_id}] Không tìm thấy nút submit")
            return False, mapped_data.get("full_name", "N/A")

        # Kiểm tra kết quả
        try:
            success_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'ĐĂNG KÝ THÀNH CÔNG')]")
                )
            )
            if success_element.is_displayed():
                print(f"[{process_id}] Form nộp thành công: Tìm thấy thông báo.")
                success = True
        except TimeoutException:
            print(
                f"[{process_id}] Form nộp thất bại: Không tìm thấy thông báo thành công sau 10 giây."
            )

        if not success:
            with open("failed_submissions.txt", "a", encoding="utf-8") as f:
                f.write(
                    f"Submission verification failed for {mapped_data.get('full_name', 'N/A')} with data {final_form_data}\n"
                )

        return success, mapped_data.get("full_name", "N/A")

    except Exception as e:
        print(f"[{process_id}] Lỗi nghiêm trọng: {e}")
        with open("failed_submissions.txt", "a", encoding="utf-8") as f:
            f.write(f"Critical error for {mapped_data.get('full_name', 'N/A')}: {e}\n")
        return False, mapped_data.get("full_name", "N/A")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
            if driver_id in active_drivers_map:
                del active_drivers_map[driver_id]


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tool Điền Form Tự Động")
        self.geometry("800x600")

        self.manager = multiprocessing.Manager()
        self.close_all_tabs_event = self.manager.Event()
        self.active_drivers_map = self.manager.dict()
        self.form_mapping = self.manager.dict()
        self.column_mapping = self.manager.dict()

        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        input_frame = tk.Frame(self, padx=10, pady=10)
        input_frame.pack(fill=tk.X)

        tk.Label(input_frame, text="URL:").grid(row=0, column=0, sticky="w")
        self.url_entry = tk.Entry(input_frame, width=50)
        self.url_entry.grid(row=0, column=1, padx=5, sticky="ew")
        self.url_entry.insert(0, "https://testpop-nine.vercel.app/registration")

        tk.Label(input_frame, text="File Excel:").grid(row=1, column=0, sticky="w")
        self.excel_path_entry = tk.Entry(input_frame, width=40, state="readonly")
        self.excel_path_entry.grid(row=1, column=1, padx=5, sticky="ew")
        tk.Button(input_frame, text="Duyệt...", command=self.browse_excel).grid(
            row=1, column=2
        )

        tk.Label(input_frame, text="Phiên Mặc Định:").grid(row=2, column=0, sticky="w")
        self.session_var = tk.StringVar(value="10:00 - 12:00")
        session_options = ["10:00 - 12:00", "13:30 - 15:30"]
        self.session_menu = tk.OptionMenu(
            input_frame, self.session_var, *session_options
        )
        self.session_menu.config(width=20)
        self.session_menu.grid(row=2, column=1, sticky="w", padx=5)

        # Thêm checkbox cho chế độ headless
        self.headless_var = tk.BooleanVar(value=True)  # Mặc định bật headless
        tk.Checkbutton(
            input_frame, text="Chạy ở chế độ ẩn (headless)", variable=self.headless_var
        ).grid(row=3, column=0, columnspan=2, sticky="w", padx=5, pady=5)

        input_frame.grid_columnconfigure(1, weight=1)

        button_frame = tk.Frame(self, padx=10, pady=5)
        button_frame.pack(fill=tk.X)

        self.start_button = tk.Button(
            button_frame, text="Bắt Đầu", command=self.start_automation, width=15
        )
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.close_button = tk.Button(
            button_frame,
            text="Đóng Tất Cả",
            command=self.close_all_tabs,
            width=15,
            state="disabled",
        )
        self.close_button.pack(side=tk.LEFT, padx=5)

        tk.Label(self, text="Nhật Ký:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(self, height=15, state="disabled")
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    def log_message(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(
            tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n"
        )
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.update_idletasks()

    def browse_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_path_entry.config(state="normal")
            self.excel_path_entry.delete(0, tk.END)
            self.excel_path_entry.insert(0, file_path)
            self.excel_path_entry.config(state="readonly")
            self.log_message(f"Đã chọn file Excel: {file_path}")

    def start_automation(self):
        self.start_button.config(state="disabled")
        self.close_button.config(state="normal")
        threading.Thread(target=self.run_automation, daemon=True).start()

    def run_automation(self):
        url = self.url_entry.get()
        excel_file = self.excel_path_entry.get()
        headless = self.headless_var.get()
        if not url or not excel_file:
            self.log_message("Lỗi: Vui lòng nhập URL và chọn file Excel")
            self.start_button.config(state="normal")
            self.close_button.config(state="disabled")
            return
        try:
            df = pd.read_excel(excel_file)
            excel_columns = df.columns.tolist()
            self.log_message(
                f"Đã đọc {len(df)} dòng dữ liệu từ Excel với các cột: {', '.join(excel_columns)}"
            )
            # Lấy ba hàng đầu tiên làm dữ liệu mẫu
            excel_data_sample = df.head(3).to_dict(orient="records")
            # Chuẩn hóa dữ liệu
            for col in excel_columns:
                df[col] = df[col].astype(str).str.strip()
                if any(
                    keyword in col.lower()
                    for keyword in ["phone", "số điện thoại", "sđt"]
                ):
                    df[col] = df[col].apply(
                        lambda x: (
                            "0" + x
                            if x
                            and x.isdigit()
                            and not x.startswith("0")
                            and len(x) in [8, 9]
                            else x
                        )
                    )
                if any(
                    keyword in col.lower()
                    for keyword in ["cccd", "id_card", "cmnd", "hộ chiếu"]
                ):
                    df[col] = df[col].apply(
                        lambda x: (
                            "0" + x
                            if x
                            and x.isdigit()
                            and not x.startswith("0")
                            and len(x) == 11
                            else x
                        )
                    )
            # Phân tích HTML và lấy sales_dates
            self.log_message("Đang phân tích form và lấy danh sách sales_date...")
            try:
                driver = webdriver.Chrome(
                    service=webdriver.chrome.service.Service(
                        ChromeDriverManager().install()
                    ),
                    options=get_chrome_options(headless),
                )
                driver.get(url)
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                driver.execute_script("return document.readyState === 'complete'")
                time.sleep(1)
                html_content = driver.page_source
                form_mapping, column_mapping, valid_sales_dates = (
                    analyze_html_and_map_columns(
                        driver, html_content, excel_columns, excel_data_sample
                    )
                )
                if not form_mapping or not column_mapping:
                    self.log_message(
                        "Lỗi: Không thể phân tích form hoặc ánh xạ cột Excel"
                    )
                    driver.quit()
                    self.start_button.config(state="normal")
                    self.close_button.config(state="disabled")
                    return
                self.form_mapping.update(form_mapping)
                self.column_mapping.update(column_mapping)
                self.log_message(
                    f"Phân tích form thành công.\nÁnh xạ form: {json.dumps(dict(form_mapping), indent=2)}\nÁnh xạ cột: {json.dumps(dict(column_mapping), indent=2)}\nDanh sách sales_date hợp lệ: {valid_sales_dates}"
                )
                driver.quit()
                if not valid_sales_dates:
                    self.log_message(
                        "Lỗi: Không tìm thấy ngày hợp lệ nào trong dropdown sales_date"
                    )
                    self.start_button.config(state="normal")
                    self.close_button.config(state="disabled")
                    return
                default_session = self.session_var.get()
                self.log_message(f"Phiên mặc định được chọn: {default_session}")
                tasks = []
                for _, row in df.iterrows():
                    for s_date in valid_sales_dates:
                        task_data = row.to_dict()
                        task_data["sales_date"] = s_date
                        task_data["session"] = task_data.get("session", default_session)
                        tasks.append(
                            (
                                url,
                                get_chrome_options(headless),
                                task_data,
                                len(tasks) + 1,
                                self.close_all_tabs_event,
                                self.active_drivers_map,
                                self.form_mapping,
                                self.column_mapping,
                                valid_sales_dates,  # Thêm valid_sales_dates vào task_data
                            )
                        )
                num_workers = min(os.cpu_count() if os.cpu_count() else 4, len(tasks))
                self.log_message(
                    f"Bắt đầu điền {len(tasks)} form với {num_workers} tiến trình"
                )
                pool = multiprocessing.Pool(processes=num_workers)
                async_results = [
                    pool.apply_async(fill_and_submit_process, (task,)) for task in tasks
                ]
                pool.close()
                success_count = 0
                for i, res in enumerate(async_results):
                    try:
                        success, data_name = res.get()
                        self.log_message(
                            f"[{i+1}/{len(tasks)}] {'Thành công' if success else 'Thất bại'}: {data_name} (Ngày: {tasks[i][2]['sales_date']})"
                        )
                        if success:
                            success_count += 1
                    except Exception as e:
                        self.log_message(
                            f"[{i+1}/{len(tasks)}] Lỗi khi nhận kết quả: {e}"
                        )
                pool.join()
                self.log_message(
                    f"\nKết quả: {success_count}/{len(tasks)} form thành công"
                )
                if len(tasks) - success_count > 0:
                    self.log_message(
                        "Kiểm tra file 'failed_submissions.txt' để xem chi tiết lỗi"
                    )
                self.start_button.config(state="normal")
                self.close_button.config(state="disabled")
                self.log_message("Hoàn tất quá trình")
            except WebDriverException as e:
                self.log_message(f"Lỗi WebDriver khi phân tích HTML: {e}")
                self.start_button.config(state="normal")
                self.close_button.config(state="disabled")
                return
            except Exception as e:
                self.log_message(f"Lỗi khi phân tích HTML hoặc lấy ngày: {e}")
                self.start_button.config(state="normal")
                self.close_button.config(state="disabled")
                return
        except FileNotFoundError:
            self.log_message("Lỗi: Không tìm thấy file Excel. Vui lòng chọn lại.")
            self.start_button.config(state="normal")
            self.close_button.config(state="disabled")

    def close_all_tabs(self):
        if not self.close_all_tabs_event.is_set():
            self.close_all_tabs_event.set()
            self.log_message("Đã gửi lệnh đóng tất cả trình duyệt")
            self.close_button.config(state="disabled")

    def on_closing(self):
        if messagebox.askokcancel("Thoát", "Bạn có muốn thoát ứng dụng?"):
            self.close_all_tabs()
            self.manager.shutdown()
            self.destroy()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = Application()
    app.mainloop()
