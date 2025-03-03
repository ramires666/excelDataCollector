import re
import time

from playwright.sync_api import Playwright, sync_playwright, expect
from datetime import datetime, timedelta
import locale

# Устанавливаем локаль на русскую
try:
    locale.setlocale(locale.LC_TIME, 'ru_KZ')  # Linux/macOS
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'ru_RU')  # Альтернативная форма
    except locale.Error:
        locale.setlocale(locale.LC_TIME, 'Russian')  # Windows

def login(headless):
    browser = playwright.chromium.launch(headless=headless)
    # browser = playwright.webkit.launch(headless=headless)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://hosting.wialon.com/?lang=ru")
    time.sleep(2)
    page.get_by_role("textbox", name="Пользователь").click()
    page.get_by_role("textbox", name="Пользователь").fill("Гром Алексей")
    page.get_by_role("textbox", name="Пароль").click()
    page.get_by_role("textbox", name="Пароль").fill("xi2tPbtT&v2#oo")
    page.get_by_role("button", name="Войти").click()
    return browser, context, page


def remove_objects_from_group(browser, context, page) -> None:
    # Read the file and prepare the list
    with open('list2remove2.txt', 'r') as file:
        # Read all lines and split by commas, stripping whitespace
        content = file.read()
        lines = [line.strip() for line in content.split('\n') if line.strip()]

    # Debug: Print the lines to verify the data
    print("Lines to process:", lines)

    # Navigate to the group section
    page.get_by_role("cell", name="").locator("#user_notification_info_icon_id").click()
    page.locator("#controls_list_view_devices_list_target_wrapper").get_by_text("Группы").click()
    page.locator("#devices_list_ed_g_3382 label").click()
    page.locator("#groups_obj_container_1_filter_right_mask").click()

    count = 0
    field_count=1
    total=len(lines)
    for line in lines:
        print(f"Processing: {line}")
        page.locator(f"#groups_obj_container_{field_count}_filter_right_mask").fill(line)
        time.sleep(1)
        page.locator(f"#groups_obj_container_{field_count}_select_all_right").click()
        # time.sleep(0.3)
        page.get_by_title("Удалить").click()
        time.sleep(0.3)
        # page.get_by_title("Удалить").click()
        # time.sleep(0.5)
        page.locator(f"#groups_obj_container_{field_count}_filter_right_mask_clear_label").click()
        # time.sleep(0.5)
        count +=1
        ready = round(count/total,2)
        print(f"{count} {ready} objects removed")
        # if count%10==0:
        #     page.locator(f"#groups_obj_container_{field_count}_filter_right_mask_clear_label").click()
        #     time.sleep(2)
        #     page.get_by_role("button", name="OK").click()
        #     field_count +=1
        #     time.sleep(1)
        #     page.locator("#devices_list_ed_g_3382 label").click()
        #     page.locator(f"#groups_obj_container_{field_count}_filter_right_mask").click()
        #     time.sleep(1)
    time.sleep(0.3)

    page.locator(f"#groups_obj_container_{field_count}_filter_right_mask_clear_label").click()
    page.get_by_role("button", name="OK").click()


def close(browser, context):
    context.close()
    browser.close()



def get_report(browser, context, page, start_date, end_date):
    page.get_by_role("button", name=" Отчеты").click()

    page.locator("#report_templates_filter_reports").click()
    page.locator("div").filter(has_text=re.compile(r"^сводка$")).first.click()

    page.get_by_role("cell", name="").locator("#user_notification_info_icon_id").click()

    page.get_by_role("combobox").nth(1).click()
    page.locator("#report_templates_filter_units").dblclick()
    time.sleep(1)
    page.locator("#report_templates_filter_units").press("End")
    time.sleep(1)
    page.locator("#report_templates_filter_units").press("Shift+Home")
    time.sleep(1)
    page.locator("#report_templates_filter_units").fill("[ВСЕ_прс]")
    time.sleep(1)
    page.locator("[data-test=\"passed\"]").press("ArrowDown")
    time.sleep(1)

    page.locator("#report_templates_filter_units").press("Shift+Home")
    time.sleep(1)
    page.locator("#report_templates_filter_units").fill("[ВСЕ_прс]")
    time.sleep(1)
    page.locator("[data-test=\"passed\"]").press("ArrowDown")
    time.sleep(1)


    def downloadHalfMonth(page, start_date, end_date):

        start_date_str = start_date.strftime("%d %B %Y 00:00")
        end_date_str = end_date.strftime("%d %B %Y 23:59")

        print("-" * 50)
        print(f"First Half: {start_date_str} - {end_date_str}")

        def capmonth(datestr):
            return datestr.replace(datestr.split()[1], datestr.split()[1].capitalize())

        start_date_str_cap = capmonth(start_date_str)
        end_date_str_cap = capmonth(end_date_str)

        page.locator("#time_from_report_templates_filter_time").click()
        page.locator("#time_from_report_templates_filter_time").press("End")
        page.locator("#time_from_report_templates_filter_time").press("Shift+Home")
        page.locator("#time_from_report_templates_filter_time").fill(start_date_str_cap)
        page.locator("#time_from_report_templates_filter_time").press("Enter")


        page.locator("#time_to_report_templates_filter_time").click()
        page.locator("#time_to_report_templates_filter_time").press("End")
        page.locator("#time_to_report_templates_filter_time").press("Shift+Home")
        page.locator("#time_to_report_templates_filter_time").fill(end_date_str_cap)
        page.locator("#time_to_report_templates_filter_time").press("Enter")

        import time

        max_wait_before_execution = 70  # Максимальное время на начало выполнения
        elapsed = 0

        waiting_locator = page.locator(".waiting-text[data-translate-phrase='execution']")
        execute_button = page.get_by_role("button", name="Выполнить")

        # Пытаемся запустить выполнение до 70 секунд
        while elapsed < max_wait_before_execution:
            execute_button.click()
            time.sleep(5)  # Ждём 5 секунд после клика
            elapsed += 5

            if waiting_locator.is_visible():
                print("Запрос начал выполняться!")
                break  # Выходим из цикла, если выполнение началось
        else:
            raise Exception("Сервер не начал выполнять запрос в течение 70 секунд!")

        # Ждем завершения выполнения (исчезновения индикатора)
        waiting_locator.wait_for(state="hidden", timeout=65000)  # Ждём максимум 10 минут
        print("Запрос успешно выполнен!")

        locator = page.locator("#report_result_buttons_contanier > label:nth-child(4) > ._m-0")
        locator.wait_for(state="visible", timeout=60000)

        with page.expect_download() as download_info:
            page.locator("#report_result_buttons_contanier > label:nth-child(4) > ._m-0").click()
        download = download_info.value
        download.save_as(fr"C:\Users\delxps\PycharmProjects\excelCollector\_mch_prob_own\{download.suggested_filename}")


    year = 2024
##########################
    # Iterate through each month (1 to 12)
    for month in range(1, 13):
        # First half of the month: 1st to 15th
        start_first_half = datetime(year, month, 1)
        end_first_half = datetime(year, month, 15)

        # Second half of the month: 16th to last day
        start_second_half = datetime(year, month, 16)

        # Calculate the last day of the month
        if month == 12:
            end_second_half = datetime(year, month, 31)
        else:
            end_second_half = datetime(year, month + 1, 1) - timedelta(days=1)

        downloadHalfMonth(page, start_first_half, end_first_half)

        downloadHalfMonth(page, start_second_half, end_second_half)


with sync_playwright() as playwright:
    headless = False
    browser, context, page = login(headless)
    # remove_objects_from_group(browser, context, page)


    start_date = datetime(2024, 1, 1)  # 01.01.2024
    end_date = datetime(2024, 12, 31)  # 31.12.2024
    get_report(browser, context, page, start_date, end_date)

    close(browser, context)