from openpyxl import load_workbook
from playwright.sync_api import sync_playwright
from selector import selector


def test_demo():
    with sync_playwright() as p:
        # Config browser and context page
        browser = p.chromium.launch(headless=False, slow_mo=1000)
        context = browser.new_context(
            record_video_dir="screen-record/",
            record_video_size={"width": 1920, "height": 1080},
        )
        page = context.new_page()
        page.set_viewport_size({"width": 1920, "height": 1080})
        page.goto("https://katalon-demo-cura.herokuapp.com/")

        # Homepage
        page.click(selector.MENU_TOGGLE)
        page.screenshot(path="screenshot/001-homepage.png", full_page=True)
        page.click(selector.LOGIN_MENU_TEXT)

        # Login page
        username = page.input_value(selector.USERNAME_TEXT)
        password = page.input_value(selector.PASSWORD_TEXT)
        page.fill(selector.USERNAME_INPUT_TEXT, username)
        page.fill(selector.PASSWORD_INPUT_TEXT, password)
        page.screenshot(path="screenshot/002-loginpage.png", full_page=True)
        page.click(selector.LOGIN_BTN)

        # Appointment page
        page.select_option(
            selector.FACILITY_DROPDOWNLIST, value="Seoul CURA Healthcare Center"
        )
        page.click(selector.APPLY_FOR_HOSTPITAL_READMISSION_CHECKBOX)
        page.click(selector.HEALTHCARE_PROGRAM_MEDICAID_RB)
        page.type(selector.VISIT_DATE_CALENDAR, "02/11/2025")
        page.keyboard.press("Tab")
        page.fill(selector.COMMENT_INPUT_TEXT, "Apping Ganteng :P")
        page.screenshot(path="screenshot/003-appointment-form.png", full_page=True)
        page.click(selector.BOOK_APPOINTMENT_BTN)

        page.screenshot(path="screenshot/004-appointment-success.png", full_page=True)

        # Save appointment to excel file
        wb = load_workbook("appointment.xlsx")
        ws = wb["appointment"]
        last_row = ws.max_row + 1
        ws["A" + str(last_row)] = page.inner_text(selector.FACILITY_TEXT)
        ws["B" + str(last_row)] = page.inner_text(
            selector.APPLY_FOR_HOSTPITAL_READMISSION_TEXT
        )
        ws["C" + str(last_row)] = page.inner_text(selector.HEALTHCARE_PROGRAM_TEXT)
        ws["D" + str(last_row)] = page.inner_text(selector.VISIT_DATE_TEXT)
        ws["E" + str(last_row)] = page.inner_text(selector.COMMENT_TEXT)

        wb.save("appointment.xlsx")

        context.close()
        browser.close()
