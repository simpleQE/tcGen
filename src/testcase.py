import re
import urllib.request as urllib2
from urllib.parse import urlparse

import click
from bs4 import BeautifulSoup

from src.view import TestsSheet

testshe = TestsSheet()


class ManualTestCases(object):
    """Class to generate test cases and its structure"""

    def __init__(self):
        self.row_count = 1

    def test_case_generator(self, url):
        """generates tests"""
        home = urlparse(url).netloc
        page = urllib2.urlopen(url)
        soup = BeautifulSoup(page, "lxml")
        soup = soup.find("body")
        self.parse_anchor_tags(soup, home)
        self.parse_button_tags(soup, home)
        self.parse_input_tags(soup, home)
        testshe.test_case_show()

    def parse_anchor_tags(self, soup, home):

        untitled_count = 0
        anchors_list = soup.find_all("a")
        for i, div in enumerate(anchors_list):
            self.row_count += 1
            non_alnum_chars = [
                ":",
                "/",
                "?",
                "#",
                "[",
                "]",
                "@",
                "!",
                "$",
                "&",
                "'",
                "(",
                ")",
                "*",
                "+",
                ",",
                ";",
                "=",
                "-",
                ".",
                "_",
                "~",
                "<",
                ">",
                "#",
                "%",
                "{",
                "}",
                "|",
                "\\",
                "^",
                "[",
                "]",
                "`",
                '"',
                " ",
            ]
            link_text = " ".join(
                (
                    "".join(
                        ch
                        for ch in div.text
                        if ch.isalnum() or ch in non_alnum_chars
                    )
                ).split()
            )
            link_url = div.get("href")
            if link_url is None or link_url.startswith("#"):
                self.row_count -= 1
                continue
            elif link_url == "/":
                link_text = "home"
            elif link_text == "":
                link_text = (
                    div.get("text")
                    or div.get("aria-label")
                    or div.get("id")
                    or div.get("aria-labelledby")
                )
                if link_text is None:
                    if div.img is not None:
                        link_text = div.img.get("alt")
                    else:
                        if link_url.startswith("http"):
                            path = urlparse(link_url).path
                        else:
                            path = link_url
                        link_text = " ".join(re.findall(r"\w+", path))
                    if link_text is None:
                        link_text = "untitled" + str(untitled_count)
                        untitled_count += 1
                case_name = link_text
            case_name = link_text.replace(" ", "_")
            rows_count = self.row_count
            testshe.write_anchor_test_case(
                link_text, case_name, div, home, rows_count
            )
            if i > 100000:
                break

    def parse_button_tags(self, soup, home):
        untitled_count = 0
        buttons_list = soup.find_all("button")
        for i, div in enumerate(buttons_list):
            self.row_count += 1
            button_text = " ".join(
                (
                    "".join(ch for ch in div.text if ch.isalnum() or ch == " ")
                ).split()
            )
            if button_text == "":
                button_text = (
                    div.get("text")
                    or div.get("name")
                    or div.get("id")
                    or div.get("title")
                    or div.get("aria-label")
                    or div.get("aria-labelledby")
                )
                if button_text is None:
                    button_text = "untitled" + str(untitled_count)
                    case_name = button_text
                    untitled_count += 1
            case_name = button_text.replace(" ", "_")
            rows_count = self.row_count
            testshe.write_button_test_case(
                button_text, case_name, div, home, rows_count
            )
            if i > 100000:
                break

    def parse_input_tags(self, soup, home):
        untitled_count = 0
        input_boxes_list = soup.find_all("input")
        for i, div in enumerate(input_boxes_list):
            if div.get("type") == "hidden":
                continue
            self.row_count += 1
            input_box_name = (
                div.get("value")
                or div.get("title")
                or div.get("placeholder")
                or div.get("alt")
                or div.get("id")
                or div.get("name")
                or div.get("aria-label")
                or div.get("aria-labelledby")
            )
            if input_box_name is None:
                input_box_name = "untitled" + str(untitled_count)
                untitled_count += 1
            case_name = input_box_name.replace(" ", "_")
            rows_count = self.row_count

            input_type = div.get("type")
            if input_type == "submit":
                testshe.write_input_submit_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "button":
                testshe.write_input_button_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "checkbox":
                testshe.write_input_checkbox_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "color":
                testshe.write_input_color_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "date":
                testshe.write_input_date_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "datetime-local":
                testshe.write_input_datetime_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "email":
                testshe.write_input_email_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "file":
                testshe.write_input_file_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "image":
                testshe.write_input_image_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "month":
                testshe.write_input_month_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "number":
                testshe.write_input_number_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "password":
                testshe.write_input_password_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "radio":
                testshe.write_input_radio_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "range":
                testshe.write_input_range_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "reset":
                testshe.write_input_reset_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "search":
                testshe.write_input_search_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "tel":
                testshe.write_input_tel_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "time":
                testshe.write_input_time_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "url":
                testshe.write_input_url_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            elif input_type == "week":
                testshe.write_input_week_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            else:
                testshe.write_input_test_case(
                    input_box_name, case_name, div, home, rows_count
                )
            if i > 100000:
                break


@click.command(help="Provide url")
@click.option("-u", "--url", default=None, help="URL to generate test cases")
def generate(url):
    """tests"""
    test = ManualTestCases()
    test.test_case_generator(url)
