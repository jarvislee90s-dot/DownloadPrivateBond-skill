import argparse
import os
import re
import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent if SCRIPT_DIR.name == "scripts" else SCRIPT_DIR


def find_project_root(start_path):
    for path in [start_path, *start_path.parents]:
        if (path / "data").exists() and (path / "reference").exists():
            return path
    return Path.cwd()


ROOT = find_project_root(SCRIPT_DIR)
DEFAULT_DOWNLOAD_DIR = ROOT / "Download"
DEFAULT_OUTPUT_DIR = ROOT / "output"
DEFAULT_EXCEL_PATTERN = "信评需求私募债_*.xlsx"


def load_skill_env(env_path=SKILL_DIR / ".env"):
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip().lstrip(":")
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def _import_selenium():
    from selenium import webdriver
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    return webdriver, TimeoutException, Options, By, EC, WebDriverWait


def extract_year_from_bond_full_name(value):
    match = re.search(r"(20\d{2})年", str(value))
    if not match:
        raise ValueError(f"债券全称中未找到年份：{value}")
    return int(match.group(1))


def normalize_for_match(value):
    text = str(value).strip()
    replacements = {
        "（": "(",
        "）": ")",
        "“": "",
        "”": "",
        "\"": "",
        " ": "",
        "\u3000": "",
        "\n": "",
        "\t": "",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def title_matches_bond(title, bond_full_name):
    return normalize_for_match(bond_full_name) in normalize_for_match(title)


def is_target_prospectus_title(title, bond_full_name):
    return "募集说明书" in str(title) and title_matches_bond(title, bond_full_name)


def group_consecutive_issuers(rows):
    groups = []
    current = []
    current_company = None
    for row in rows:
        company = row["company_name"]
        if current and company != current_company:
            groups.append(current)
            current = []
        current.append(row)
        current_company = company
    if current:
        groups.append(current)
    return groups


def read_bond_rows(excel_path):
    workbook = load_workbook(excel_path, data_only=True)
    sheet = workbook.active
    rows = []
    for row_index in range(2, sheet.max_row + 1):
        company_name = sheet.cell(row=row_index, column=1).value
        bond_short_name = sheet.cell(row=row_index, column=2).value
        bond_code = sheet.cell(row=row_index, column=3).value
        bond_full_name = sheet.cell(row=row_index, column=4).value
        issue_method = sheet.cell(row=row_index, column=5).value
        if not company_name or not bond_full_name:
            continue
        rows.append(
            {
                "company_name": str(company_name).strip().replace("（", "(").replace("）", ")"),
                "bond_short_name": str(bond_short_name or "").strip(),
                "bond_code": str(bond_code or "").strip(),
                "bond_full_name": str(bond_full_name).strip(),
                "issue_method": str(issue_method or "").strip(),
            }
        )
    return rows


def find_latest_prepared_excel(output_dir=DEFAULT_OUTPUT_DIR):
    output_dir = Path(output_dir)
    candidates = [
        path for path in output_dir.glob(DEFAULT_EXCEL_PATTERN)
        if path.is_file()
    ]
    if not candidates:
        raise FileNotFoundError(
            f"未在 {output_dir} 找到 {DEFAULT_EXCEL_PATTERN}，请先运行 prepare_bond_excel.py 或手动传入 --excel"
        )
    return max(candidates, key=lambda path: path.stat().st_mtime)


def resolve_excel_path(excel_path, output_dir=DEFAULT_OUTPUT_DIR):
    if excel_path:
        return Path(excel_path)
    return find_latest_prepared_excel(output_dir)


def create_driver(download_dir):
    webdriver, _, Options, _, _, _ = _import_selenium()
    download_dir = Path(download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)
    prefs = {
        "download.default_directory": str(download_dir.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "safebrowsing.enabled": False,
    }
    options = Options()
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    driver.download_dir = download_dir
    driver.set_window_size(1400, 900)
    return driver


def find_visible(driver, by, locator, timeout=20):
    _, TimeoutException, _, _, _, WebDriverWait = _import_selenium()
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        try:
            elements = driver.find_elements(by, locator)
            for element in elements:
                if element.is_displayed():
                    return element
        except Exception as exc:
            last_error = exc
        time.sleep(0.2)
    raise TimeoutException(f"未找到可见元素：{locator}") from last_error


def visible_input(driver, placeholder, timeout=20):
    _, _, _, By, _, _ = _import_selenium()
    return find_visible(driver, By.XPATH, f"//input[@placeholder='{placeholder}']", timeout=timeout)


def click_if_visible(driver, xpath, timeout=3):
    _, TimeoutException, _, By, _, _ = _import_selenium()
    try:
        element = find_visible(driver, By.XPATH, xpath, timeout=timeout)
        driver.execute_script("arguments[0].click();", element)
        return True
    except TimeoutException:
        return False


def login(driver, username, password):
    _, _, _, By, EC, WebDriverWait = _import_selenium()
    wait = WebDriverWait(driver, 30)
    for attempt in range(2):
        driver.get("https://www.ratingdog.cn/login")
        click_if_visible(driver, "//*[text()='手机密码登录']", timeout=5)
        phone_input, password_input = visible_login_inputs(driver, timeout=60)
        phone_input.clear()
        password_input.clear()
        phone_input.send_keys(username)
        password_input.send_keys(password)
        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='登录']]")))
        driver.execute_script("arguments[0].click();", login_button)
        wait_for_login_submit(driver)
        driver.get("https://www.ratingdog.cn/researchIssuer/yyRating")
        try:
            visible_input(driver, "代码、简称、发行人", timeout=30)
            return
        except Exception:
            if attempt == 1:
                raise
            time.sleep(2)


def visible_login_inputs(driver, timeout=60):
    _, TimeoutException, _, By, _, _ = _import_selenium()
    deadline = time.time() + timeout
    while time.time() < deadline:
        inputs = [
            element for element in driver.find_elements(By.TAG_NAME, "input")
            if element.is_displayed()
        ]
        phone_inputs = [
            element for element in inputs
            if (element.get_attribute("type") or "").lower() in ("text", "tel")
        ]
        password_inputs = [
            element for element in inputs
            if (element.get_attribute("type") or "").lower() == "password"
        ]
        if phone_inputs and password_inputs:
            return phone_inputs[0], password_inputs[0]
        time.sleep(0.2)
    raise TimeoutException("未找到登录页手机号和密码输入框")


def wait_for_login_submit(driver, timeout=60):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if "login" not in driver.current_url:
            return
        logged_in = driver.execute_script(
            """
            const keys = Object.keys(window.localStorage || {});
            return keys.some(key => /token|auth/i.test(key) && localStorage.getItem(key));
            """
        )
        if logged_in:
            time.sleep(2)
            return
        time.sleep(0.5)
    return


def _has_login_state(driver):
    if "login" not in driver.current_url:
        return True
    try:
        return bool(
            driver.execute_script(
                """
                const keys = Object.keys(window.localStorage || {});
                return keys.some(key => /token|auth|user/i.test(key) && localStorage.getItem(key));
                """
            )
        )
    except Exception:
        return False


def search_issuer(driver, company_name):
    _, _, _, By, EC, WebDriverWait = _import_selenium()
    wait = WebDriverWait(driver, 30)
    search_input = visible_input(driver, "代码、简称、发行人")
    search_input.clear()
    search_input.send_keys(company_name)
    search_button = search_input.find_element(
        By.XPATH,
        "./ancestor::form[1]//button[contains(@class,'yyep-button--default')]",
    )
    driver.execute_script("arguments[0].click();", search_button)
    first_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[1]/td[1]//a[contains(@class,'details')]")))
    driver.execute_script("arguments[0].click();", first_link)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@role='tab' and .//span[text()='相关公告']]")))


def open_announcements_tab(driver):
    _, TimeoutException, _, By, _, _ = _import_selenium()
    deadline = time.time() + 30
    last_error = None

    while time.time() < deadline:
        try:
            tab = find_visible(
                driver,
                By.XPATH,
                "//*[@role='tab' and .//span[text()='相关公告']]",
                timeout=2,
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", tab)
            try:
                tab.click()
            except Exception:
                driver.execute_script("arguments[0].click();", tab)

            visible_input(driver, "标题", timeout=2)
            return
        except Exception as exc:
            last_error = exc
            time.sleep(0.3)

    raise TimeoutException("点击相关公告后未出现标题筛选框") from last_error


def set_private_checkbox(driver, checked):
    _, TimeoutException, _, By, _, _ = _import_selenium()
    try:
        label = find_visible(
            driver,
            By.XPATH,
            "//label[contains(@class,'yyep-checkbox') and .//*[contains(text(),'只看私募')]]",
            timeout=10,
        )
    except TimeoutException:
        label = driver.execute_script(
            """
            return [...document.querySelectorAll('label.yyep-checkbox')]
              .find(el => el.innerText && el.innerText.includes('只看私募')) || null;
            """
        )
        if label is None:
            raise
    checkbox = label.find_element(By.XPATH, ".//input[@type='checkbox']")
    if checkbox.is_selected() != checked:
        driver.execute_script("arguments[0].click();", label)
        time.sleep(1)


def clear_and_type(element, value):
    element.clear()
    if value:
        element.send_keys(value)


def dismiss_date_picker(driver, fallback_element=None):
    from selenium.webdriver.common.keys import Keys

    active_element = driver.switch_to.active_element
    active_element.send_keys(Keys.ESCAPE)
    time.sleep(0.3)
    if fallback_element is not None:
        driver.execute_script("arguments[0].click();", fallback_element)
    else:
        driver.execute_script("document.body.click();")
    time.sleep(0.5)


def set_date_range(driver, start_input, end_input, year):
    from selenium.webdriver.common.keys import Keys

    start_value = f"{year}-01-01" if year else ""
    end_value = f"{year}-12-31" if year else ""

    def set_input_value(element, value):
        element.click()
        element.send_keys(Keys.CONTROL, "a")
        element.send_keys(Keys.BACKSPACE)
        if value:
            element.send_keys(value)
        driver.execute_script(
            """
            const el = arguments[0];
            const value = arguments[1];
            const setter = Object.getOwnPropertyDescriptor(
              window.HTMLInputElement.prototype,
              'value'
            ).set;
            setter.call(el, value);
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new Event('blur', { bubbles: true }));
            """,
            element,
            value,
        )

    set_input_value(start_input, start_value)
    set_input_value(end_input, end_value)
    end_input.send_keys(Keys.ENTER)
    dismiss_date_picker(driver, fallback_element=start_input)

    actual_start = start_input.get_attribute("value") or ""
    actual_end = end_input.get_attribute("value") or ""
    if start_value and actual_start != start_value:
        driver.execute_script(
            """
            const el = arguments[0];
            el.value = arguments[1];
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            """,
            start_input,
            start_value,
        )
    if end_value and actual_end != end_value:
        driver.execute_script(
            """
            const el = arguments[0];
            el.value = arguments[1];
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            """,
            end_input,
            end_value,
        )
    dismiss_date_picker(driver, fallback_element=start_input)


def search_announcements(driver, year=None, private_only=True):
    _, _, _, By, _, _ = _import_selenium()
    title_input = visible_input(driver, "标题")
    clear_and_type(title_input, "")
    start_input = visible_input(driver, "开始日期")
    end_input = visible_input(driver, "结束日期")
    set_date_range(driver, start_input, end_input, year)
    set_private_checkbox(driver, private_only)
    search_button = title_input.find_element(By.XPATH, "./ancestor::form[1]//button[contains(@class,'yyep-button--default')]")
    driver.execute_script("arguments[0].click();", search_button)
    wait_for_announcement_results(driver)


def wait_for_announcement_results(driver, timeout=30):
    _, TimeoutException, _, By, _, _ = _import_selenium()
    deadline = time.time() + timeout
    while time.time() < deadline:
        rows = driver.find_elements(By.XPATH, "//tbody/tr")
        if any(row.is_displayed() for row in rows):
            return
        empty_hints = driver.find_elements(By.XPATH, "//*[contains(text(),'暂无数据') or contains(text(),'无数据')]")
        if any(element.is_displayed() for element in empty_hints):
            return
        time.sleep(0.2)
    raise TimeoutException("公告搜索后未出现结果行或空结果提示")


def close_download_dialog(driver):
    _, TimeoutException, _, By, EC, WebDriverWait = _import_selenium()
    try:
        button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@role='dialog' and not(contains(@style,'display: none'))]//button[.//*[text()='确定']]"))
        )
        driver.execute_script("arguments[0].click();", button)
    except TimeoutException:
        pass


def find_and_download(driver, bond_full_name):
    _, _, _, By, _, _ = _import_selenium()
    rows = driver.find_elements(By.XPATH, "//tbody/tr")
    for row in rows:
        if not row.is_displayed():
            continue
        title_cells = row.find_elements(By.XPATH, "./td[1]")
        if not title_cells:
            continue
        title_text = title_cells[0].text.strip()
        if is_target_prospectus_title(title_text, bond_full_name):
            download_elements = row.find_elements(By.XPATH, ".//span[contains(@class,'file-desc') and contains(.,'下载')]")
            if not download_elements:
                return False, f"找到公告但无附件下载：{title_text}"
            before_files = snapshot_download_files(driver)
            driver.execute_script("arguments[0].click();", download_elements[0])
            close_download_dialog(driver)
            downloaded = wait_for_download_complete(driver, before_files)
            if not downloaded:
                return False, f"下载未在限定时间内完成：{title_text}"
            return True, title_text
    return False, "未找到标题匹配的募集说明书公告"


def snapshot_download_files(driver):
    download_dir = Path(getattr(driver, "download_dir", DEFAULT_DOWNLOAD_DIR))
    download_dir.mkdir(parents=True, exist_ok=True)
    return {path.name for path in download_dir.iterdir() if path.is_file()}


def wait_for_download_complete(driver, before_files, timeout=120):
    download_dir = Path(getattr(driver, "download_dir", DEFAULT_DOWNLOAD_DIR))
    deadline = time.time() + timeout
    while time.time() < deadline:
        files = [path for path in download_dir.iterdir() if path.is_file()]
        new_files = [path for path in files if path.name not in before_files]
        partial_files = [path for path in new_files if path.name.endswith(".crdownload")]
        complete_pdfs = [path for path in new_files if path.suffix.lower() == ".pdf"]
        if complete_pdfs and not partial_files:
            return complete_pdfs[0]
        time.sleep(0.5)
    return None


def close_active_issuer_tab(driver):
    _, _, _, By, _, _ = _import_selenium()
    icons = driver.find_elements(By.XPATH, "//li[contains(@class,'tags-li') and contains(@class,'active')]//span[contains(@class,'tags-li-icon')]")
    if icons:
        driver.execute_script("arguments[0].click();", icons[0])
        time.sleep(1)


def switch_to_rating_tab(driver):
    _, _, _, By, _, _ = _import_selenium()
    tab = find_visible(
        driver,
        By.XPATH,
        "//li[contains(@class,'tags-li') and .//*[text()='主体评级']]",
        timeout=20,
    )
    driver.execute_script("arguments[0].click();", tab)
    visible_input(driver, "代码、简称、发行人", timeout=20)


def write_log(log_path, lines):
    Path(log_path).parent.mkdir(parents=True, exist_ok=True)
    Path(log_path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def process_rows(driver, rows):
    log_lines = []
    valid_rows = []
    for row in rows:
        try:
            extract_year_from_bond_full_name(row["bond_full_name"])
        except ValueError as exc:
            log_lines.append(
                f"{row['company_name']}\t{row['bond_short_name']}\t{row['bond_full_name']}\t{exc}"
            )
            continue
        valid_rows.append(row)

    for group in group_consecutive_issuers(valid_rows):
        company_name = group[0]["company_name"]
        switch_to_rating_tab(driver)
        search_issuer(driver, company_name)
        open_announcements_tab(driver)
        for row in group:
            bond_full_name = row["bond_full_name"]
            year = extract_year_from_bond_full_name(bond_full_name)
            search_announcements(driver, year=year, private_only=True)
            ok, message = find_and_download(driver, bond_full_name)
            if not ok:
                search_announcements(driver, year=None, private_only=False)
                ok, message = find_and_download(driver, bond_full_name)
            if not ok:
                log_lines.append(
                    f"{company_name}\t{row['bond_short_name']}\t{bond_full_name}\t{message}"
                )
        close_active_issuer_tab(driver)
    return log_lines


def main():
    load_skill_env()
    parser = argparse.ArgumentParser(description="从 Ratingdog 下载私募债相关公告附件")
    parser.add_argument("--excel", required=True, help="prepare_bond_excel.py 输出的 Excel")
    parser.add_argument("--download-dir", default=str(DEFAULT_DOWNLOAD_DIR), help="PDF 下载目录")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="日志输出目录")
    parser.add_argument("--username", default=os.environ.get("RATINGDOG_USERNAME", ""), help="Ratingdog 手机号")
    parser.add_argument("--password", default=os.environ.get("RATINGDOG_PASSWORD", ""), help="Ratingdog 密码")
    for action in parser._actions:
        if action.dest == "excel":
            action.required = False
            action.help = "prepare_bond_excel.py 输出的 Excel；未提供时自动使用 output 目录下最新的 信评需求私募债_*.xlsx"
            break
    args = parser.parse_args()

    if not args.username or not args.password:
        raise ValueError("请通过参数或环境变量提供 RATINGDOG_USERNAME/RATINGDOG_PASSWORD")

    excel_path = resolve_excel_path(args.excel, args.output_dir)
    print(f"Excel：{excel_path}")

    rows = read_bond_rows(excel_path)
    driver = create_driver(args.download_dir)
    try:
        login(driver, args.username, args.password)
        log_lines = process_rows(driver, rows)
    finally:
        driver.quit()

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = Path(args.output_dir) / f"download_log_{stamp}.txt"
    write_log(log_path, log_lines or ["全部债券均已处理，未记录失败项"])
    print(f"日志：{log_path}")


if __name__ == "__main__":
    main()
