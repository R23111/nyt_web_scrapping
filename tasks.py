"""
Script to get New York Times articles from a search within a date range, and a section type
"""

import re
from datetime import datetime
from urllib import request

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Robocorp.WorkItems import WorkItems
from SeleniumLibrary.errors import ElementNotFound
from dateutil.relativedelta import relativedelta
from selenium.common import StaleElementReferenceException, ElementClickInterceptedException
from selenium.webdriver.remote.webelement import WebElement


class ArticleData:
    """
    Article data from the web page.
    """
    __money_regex = r"(\$)(([1-9]\d{0,2}(\,\d{3})*)|([1-9]\d*)|" \
                    r"(0))(\.\d{2})?|\d+ (\bdollars\b|\busd\b|" \
                    r"\bdollar\b)+"

    def __init__(self, web_element: WebElement, search_phrase: str) -> None:
        """
        :param web_element: WebElement of the article
        :param search_phrase: The phrase used to search
        """
        self.date = web_element.find_element(
            by="xpath", value=".//span[@data-testid='todays-date']").text
        self.title = web_element.find_element(by="xpath", value=".//h4").text
        self.description = web_element.find_elements(
            by="xpath",
            value=".//p"
        )[1].text
        self.picture_url = web_element.find_element(
            by="xpath",
            value=".//img"
        ).get_property("src").split('?')[0]

        description_lower = self.description.lower()
        title_lower = self.title.lower()

        self.n_of_phrases = description_lower.count(search_phrase) + \
                            title_lower.count(search_phrase)

        self.has_money = bool(re.search(self.__money_regex, title_lower)) or bool(
            re.search(self.__money_regex, description_lower))
        self.url = web_element.find_element(
            by="xpath",
            value=".//a"
        ).get_property("href").split('?')[0]

    def to_dict(self) -> dict:
        """
        :return: dictionary of the article data
        """
        return {
            "date": self.date,
            "title": self.title,
            "description": self.description,
            "picture_file_name": self.picture_url,
            "n_of_phrases": self.n_of_phrases,
            "has_money": self.has_money
        }

    def __str__(self):
        return str(self.to_dict())

    def __repr__(self):
        return str(self.to_dict())


wi = WorkItems()
wi.get_input_work_item()
SEARCH_PHRASE = wi.get_work_item_variable("search_phrase").lower()
NEWS_CATEGORY = wi.get_work_item_variable("news_category")
NUMBER_OF_MONTHS = wi.get_work_item_variable("number_of_months")

browser_lib = Selenium()


def open_the_website(url: str) -> None:
    """
    Open the website on the given url
    :param url: site url
    """
    browser_lib.open_available_browser(url)


def accept_cookies() -> None:
    """
    Accept cookies
    """
    try:
        accept_cookies_test_data_id = "//button[@data-testid='GDPR-accept']"
        browser_lib.wait_until_element_is_visible(accept_cookies_test_data_id)
        browser_lib.wait_and_click_button(accept_cookies_test_data_id)
        browser_lib.wait_until_element_is_not_visible(accept_cookies_test_data_id)
        browser_lib.wait_until_element_is_not_visible('//div[@data-testid="response-snackbar"]')
    except AssertionError as error:
        if "still visible after 5 seconds." in error.args[0]:
            return


def click_search_button() -> None:
    """
    Click the search button
    """
    data_test_id = "//button[@data-test-id='search-button']"
    browser_lib.click_button(data_test_id)


def search_for(term: str) -> None:
    """
    Search for the given term
    :param term: phrase to be searched
    """
    data_test_id = "//input[@data-testid='search-input']"
    browser_lib.input_text(data_test_id, term)


def submit_search() -> None:
    """
    Submits the search
    :return:
    """
    data_test_id = "//button[@data-test-id='search-submit']"
    browser_lib.click_button(data_test_id)


def set_section(section: str) -> None:
    """
    Select the news section to filter the search result
    :param section: news section to be selected
    """
    section_selection_data_test_id = '//div[@data-testid="section"]' \
                                     ' //button[@data-testid="search-multiselect-button"]'
    browser_lib.click_button(section_selection_data_test_id)
    dropdown_data_test_id = "//div[@data-testid='section']" \
                            " //ul[@data-testid='multi-select-dropdown-list']" \
                            " //li //label[@data-testid='DropdownLabel']"
    labels = browser_lib.find_elements(dropdown_data_test_id)
    labels_dict = {}
    for label in labels:
        section_name = re.sub(r"[^A-Za-z.]", "", label.text).lower()
        labels_dict[section_name] = label

    browser_lib.click_button(labels_dict[section.lower()])


def set_date_span(number_of_months: int) -> None:
    """
    Set the date span to filter the search result
    :param number_of_months: number of months including the current one
    """
    number_of_months = max(number_of_months, 0)

    today = datetime.today()
    today_date = today.strftime("%m/%d/%Y")

    if number_of_months:
        number_of_months -= 1
    target = today - relativedelta(months=number_of_months)
    target_date = target.strftime("%m/01/%Y")

    date_range_data_test_id = "//button[@data-testid='search-date-dropdown-a']"
    browser_lib.click_button(date_range_data_test_id)

    browser_lib.click_button("//button[@value='Specific Dates']")
    browser_lib.input_text("//input[@data-testid='DateRange-startDate']", target_date)
    browser_lib.input_text("//input[@data-testid='DateRange-endDate']", today_date)

    browser_lib.click_button(date_range_data_test_id)


def get_articles_web_element() -> list[WebElement]:
    """
    Loads the articles from the search result and store them in a list of web elements
    :return: list of found web elements to the corresponding articles
    """
    show_more_data_test_id = '//button[@data-testid="search-show-more-button"]'

    while browser_lib.does_page_contain_element(show_more_data_test_id):
        try:
            browser_lib.press_key(show_more_data_test_id, key="end")
            browser_lib.find_element(show_more_data_test_id).click()
        except (ElementClickInterceptedException, StaleElementReferenceException):
            continue
        except ElementNotFound:
            break

    return browser_lib.find_elements('//ol[@data-testid="search-results"] '
                                     '//li[@data-testid="search-bodega-result"]')


def get_articles_data(articles_web_element: list[WebElement]) -> list[ArticleData]:
    """
    Loads the articles from the search result web element and store them in a list of ArticleData
    :param articles_web_element:
    :return:
    """
    article_dict = {}
    idx = 0
    while idx < len(articles_web_element):
        web_element = articles_web_element[idx]
        try:
            browser_lib.wait_until_page_contains_element(web_element)
            article_data = ArticleData(web_element, SEARCH_PHRASE)
            if article_data.url not in article_dict:
                article_dict[article_data.url] = article_data
            idx += 1
        except StaleElementReferenceException:
            articles_web_element = browser_lib.find_elements(
                '//ol[@data-testid="search-results"] '
                '//li[@data-testid="search-bodega-result"]')

    return list(article_dict.values())


def save_to_excel(articles: list[ArticleData]) -> None:
    """
    Save the articles to excel, creating the file if it doesn't exist, and appending data if it does
    :param articles: list of articles to be saved
    """
    lib = Files()
    lib.create_workbook("")
    excel_file_path = "./result.xlsx"
    content = [article.to_dict() for article in articles]
    try:
        lib.open_workbook(excel_file_path)
    except FileNotFoundError:
        lib.create_workbook(excel_file_path)
        lib.save_workbook()
        lib.open_workbook(excel_file_path)
    finally:
        lib.append_rows_to_worksheet(content=content, header=True)
        lib.save_workbook()
        lib.close_workbook()


def get_images(articles: list[ArticleData]) -> None:
    """
    Download the images of the articles
    :param articles: list of Articles to have the image downloaded
    """
    for article in articles:
        file_name = article.picture_url.split('/')[-1]
        request.urlretrieve(article.picture_url, file_name)
        article.picture_url = file_name


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
        articles_web_element = get_articles_web_element()
        articles_data = get_articles_data(articles_web_element)
        get_images(articles_data)
        save_to_excel(articles_data)
        print(len(articles_data))
        print(articles_data)
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
