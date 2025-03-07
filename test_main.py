import os
from openpyxl import load_workbook, Workbook
from playwright.sync_api import sync_playwright
from selector import select


def test_demo():
    print("\nRunning automation test...")

    with sync_playwright() as p:
        # Config browser and context page
        browser = p.chromium.launch(headless=True, slow_mo=1000)
        context = browser.new_context(
            record_video_dir="screen-record/",
            record_video_size={"width": 1920, "height": 1080},
        )

        # Trace activity
        context.tracing.start(screenshots=True, snapshots=True, sources=True)

        page = context.new_page()
        page.set_viewport_size({"width": 1920, "height": 1080})
        page.goto("https://katalon-demo-cura.herokuapp.com/")

        # Homepage
        page.click(select.MENU_TOGGLE)
        page.screenshot(path="screenshot/001-homepage.png", full_page=True)
        page.click(select.LOGIN_MENU_TEXT)

        # Login page
        username = page.input_value(select.USERNAME_TEXT)
        password = page.input_value(select.PASSWORD_TEXT)
        page.fill(select.USERNAME_INPUT_TEXT, username)
        page.fill(select.PASSWORD_INPUT_TEXT, password)
        page.screenshot(path="screenshot/002-loginpage.png", full_page=True)
        page.click(select.LOGIN_BTN)

        # Appointment page
        page.select_option(select.FACILITY_DROPDOWNLIST, value="Seoul CURA Healthcare Center")
        page.click(select.APPLY_FOR_HOSTPITAL_READMISSION_CHECKBOX)
        page.click(select.HEALTHCARE_PROGRAM_MEDICAID_RB)
        page.type(select.VISIT_DATE_CALENDAR, "02/11/2025")
        page.keyboard.press("Tab")
        page.fill(select.COMMENT_INPUT_TEXT, "Apping Ganteng :P")
        page.screenshot(path="screenshot/003-appointment-form.png", full_page=True)
        page.click(select.BOOK_APPOINTMENT_BTN)
        page.screenshot(path="screenshot/004-appointment-success.png", full_page=True)

        # Save appointment to excel file
        if not os.path.exists("appointment.xlsx"):
            wb = Workbook()
            ws = wb.active
            ws.title = "appointment"
            ws["A1"] = "Facility"
            ws["B1"] = "Apply for hospital readmission"
            ws["C1"] = "Healthcare Program"
            ws["D1"] = "Visit Date"
            ws["E1"] = "Comment"
            ws["F1"] = "Status"
            wb.save("appointment.xlsx")
        else:
            wb = load_workbook("appointment.xlsx")
            ws = wb["appointment"]

        last_row = ws.max_row + 1
        ws["A" + str(last_row)] = page.inner_text(select.FACILITY_TEXT)
        ws["B" + str(last_row)] = page.inner_text(select.APPLY_FOR_HOSTPITAL_READMISSION_TEXT)
        ws["C" + str(last_row)] = page.inner_text(select.HEALTHCARE_PROGRAM_TEXT)
        ws["D" + str(last_row)] = page.inner_text(select.VISIT_DATE_TEXT)
        ws["E" + str(last_row)] = page.inner_text(select.COMMENT_TEXT)
        ws["F" + str(last_row)] = "Appointment Success"
        wb.save("appointment.xlsx")
        wb.close()

        context.tracing.stop(path="trace.zip")

        context.close()
        browser.close()

    print("Done")
