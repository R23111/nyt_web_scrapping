import re
from datetime import datetime

from RPA.Browser.Selenium import Selenium
from SeleniumLibrary.errors import ElementNotFound
from dateutil.relativedelta import relativedelta
from selenium.common import StaleElementReferenceException, ElementClickInterceptedException
from RPA.Excel.Files import Files
import urllib.request

SEARCH_PHRASE = "brazil"
NEWS_CATEGORY = "World"
NUMBER_OF_MONTHS = 1

browser_lib = Selenium()


def open_the_website(url):
    browser_lib.open_available_browser(url)


def accept_cookies():
    try:
        browser_lib.wait_until_element_is_visible('//button[@data-testid="GDPR-accept"]')
        browser_lib.wait_and_click_button('//button[@data-testid="GDPR-accept"]')
        browser_lib.wait_until_element_is_not_visible('//button[@data-testid="GDPR-accept"]')
        browser_lib.wait_until_element_is_not_visible('//div[@data-testid="response-snackbar"]')
    except AssertionError as e:
        if "still visible after 5 seconds." in e.args[0]:
            return
    except Exception as e:
        print(e)
        return accept_cookies()


def click_search_button():
    data_test_id = "//button[@data-test-id='search-button']"
    browser_lib.click_button(data_test_id)


def search_for(term):
    data_test_id = "//input[@data-testid='search-input']"
    browser_lib.input_text(data_test_id, term)


def submit_search():
    data_test_id = "//button[@data-test-id='search-submit']"
    browser_lib.click_button(data_test_id)


def set_section(section):
    section_selection_data_test_id = '//div[@data-testid="section"] //button[@data-testid="search-multiselect-button"]'
    browser_lib.click_button(section_selection_data_test_id)
    dropdown_data_test_id = "//div[@data-testid='section'] //ul[@data-testid='multi-select-dropdown-list'] //li //label[@data-testid='DropdownLabel']"
    labels = browser_lib.find_elements(dropdown_data_test_id)
    d = {}
    for label in labels:
        section_name = re.sub(r"[^A-Za-z.]", "", label.text).lower()
        d[section_name] = label

    browser_lib.click_button(d[section.lower()])


def set_date_span(number_of_months):
    today = datetime.today()
    today_date = today.strftime("%m/%d/%Y")

    if number_of_months:
        number_of_months -= 1
    target = today - relativedelta(months=number_of_months)
    target_date = target.strftime("%m/01/%Y")

    date_range_data_test_id = "//button[@data-testid='search-date-dropdown-a']"
    browser_lib.click_button(date_range_data_test_id)

    specific_date_value = "//button[@value='Specific Dates']"
    browser_lib.click_button(specific_date_value)

    start_date_data_test_id = "//input[@data-testid='DateRange-startDate']"
    browser_lib.input_text(start_date_data_test_id, target_date)

    end_date_data_test_id = "//input[@data-testid='DateRange-endDate']"
    browser_lib.input_text(end_date_data_test_id, today_date)

    browser_lib.click_button(date_range_data_test_id)
    print(today_date, target_date)


def get_article_data(web_element):
    browser_lib.wait_until_page_contains_element(web_element)
    date = web_element.find_element(by="xpath", value=".//span[@data-testid='todays-date']").text
    title = web_element.find_element(by="xpath", value=".//h4").text
    description = web_element.find_elements(by="xpath", value=".//p")[1].text
    picture_file_name = web_element.find_element(by="xpath", value=".//img").get_property("src").split('?')[0]
    n_of_phrases = description.lower().count(SEARCH_PHRASE.lower()) + title.lower().count(SEARCH_PHRASE.lower())
    money_regex = r"(\$)(([1-9]\d{0,2}(\,\d{3})*)|([1-9]\d*)|(0))(\.\d{2})?|[0-9]+ (\bdollars\b|\busd\b|\bdollar\b)+"
    has_money = bool(re.search(money_regex, title.lower())) or bool(
        re.search(money_regex, description.lower()))
    article_url = web_element.find_element(by="xpath", value=".//a").get_property("href").split('?')[0]
    return {'date': date, 'title': title, 'description': description, 'picture_file_name': picture_file_name,
            'url': article_url, 'number_of_phrase_occurrences': n_of_phrases,
            'has_amount_of_money': has_money}


def get_search_result():
    show_more_data_test_id = '//button[@data-testid="search-show-more-button"]'
    d = {}
    idx = 0
    while browser_lib.does_page_contain_element(show_more_data_test_id):
        try:
            browser_lib.press_key(show_more_data_test_id, key="end")
            browser_lib.find_element(show_more_data_test_id).click()
        except ElementClickInterceptedException:
            continue
        except ElementNotFound:
            break

    search_result_data_test_id = '//ol[@data-testid="search-results"] //li[@data-testid="search-bodega-result"]'
    search_result = browser_lib.find_elements(search_result_data_test_id)

    while idx < len(search_result):
        result = search_result[idx]
        try:
            article_data = get_article_data(result)
            # try:
            #     if d[article_data['url']] is None:
            #         idx += 1
            #         continue
            # except KeyError:
            #     d[article_data['url']] = article_data
            if article_data['url'] not in d.keys():
                d[article_data['url']] = article_data
            idx += 1
        except StaleElementReferenceException:
            search_result = browser_lib.find_elements(search_result_data_test_id)

    return list(d.values())


def save_to_excel(result):
    lib = Files()
    lib.create_workbook("")
    try:
        lib.open_workbook("./result.xlsx")
        lib.append_rows_to_worksheet(content=result, header=True)
        lib.save_workbook()
    except FileNotFoundError:
        lib.create_workbook("./result.xlsx")
        lib.save_workbook()
        lib.open_workbook("./result.xlsx")
        lib.append_rows_to_worksheet(content=result, header=True)
        lib.save_workbook()
    finally:
        lib.close_workbook()


def get_images(search_result):
    for result in search_result:
        file_name = result['picture_file_name'].split('/')[-1]
        urllib.request.urlretrieve(result['picture_file_name'], file_name)
        result['picture_file_name'] = file_name
        result.pop('url')



def main():
    try:
        open_the_website("https://www.nytimes.com/")
        accept_cookies()
        click_search_button()
        search_for(SEARCH_PHRASE)
        submit_search()
        set_section(NEWS_CATEGORY)
        set_date_span(NUMBER_OF_MONTHS)
        browser_lib.wait_until_page_contains(SEARCH_PHRASE)
        search_result = get_search_result()
        get_images(search_result)
        save_to_excel(search_result)
        print(len(search_result))
        print(search_result)
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
