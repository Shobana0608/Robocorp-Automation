from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.HTTP import HTTP
import json
from datetime import datetime

time_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
import logging

logging.info("Automation started")
@task
def step_one_open_example():
    browser = Selenium()

    browser.open_available_browser("https://example.com")

    browser.capture_page_screenshot("output/example.png")

    browser.close_browser()




@task
def step_two_google_search():

    browser = Selenium()

    try:

        browser.open_available_browser(
            "https://robotsparebinindustries.com",
            headless=True
        )

        browser.wait_until_page_contains(
            "RobotSpareBin Industries"
        )

        browser.capture_page_screenshot(
            "output/robotsparebin.png"
        )

        print("Website opened successfully")

        return "Website opened successfully"

    except Exception as e:

        print(f"Error occurred: {e}")

        return "Failed"

    finally:

        browser.close_all_browsers()
@task
def step_three_fetch_api():

    http = HTTP()

    http.download(
        "https://api.github.com/repos/robocorp/robocorp",
        "output/repo.json"
    )

    with open("output/repo.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    return data["description"]


@task
def step_four_save_excel():

    search_result = step_two_google_search()

    repo_description = step_three_fetch_api()

    excel = Files()

    excel.create_workbook("output/report.xlsx")

    excel.append_rows_to_worksheet(
        [[search_result, repo_description]],
        header=["Google Search Result", "Repo Description"]
    )

    excel.save_workbook()

    excel.close_workbook()

    print("Excel report created successfully")

