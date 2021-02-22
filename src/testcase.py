import os
import re
import urllib.request as urllib2
from urllib.parse import urlparse

import click
import xlsxwriter
from bs4 import BeautifulSoup


class ManualTestCases(object):
    """Class to generate test cases and its structure"""

    def __init__(self):

        self.config_path = os.path.expanduser("~")
        self.testcasefile = os.path.join(self.config_path, "output.xlsx")

        self._check_file()
        self.row_count = 1
        self.workbook = xlsxwriter.Workbook(self.testcasefile)
        self.workbook.window_width = 1920
        self.workbook.window_height = 720
        self.cell_format = self.workbook.add_format()
        self.workbook.formats[0].set_font_size(16)
        self.worksheet = self.workbook.add_worksheet()
        cell_format = self.workbook.add_format(
            {"bold": True, "font_size": 18, "bg_color": "cyan", "border": 1}
        )
        cell_format.set_align("center")
        self.workbook.formats[0].set_align("vcenter")
        self.worksheet.set_row(0, 30, cell_format)
        self.worksheet.set_column(0, 10, 30)
        self.worksheet.set_column(2, 2, 20)
        self.worksheet.set_column(5, 5, 40)
        self.worksheet.set_column(7, 8, 15)
        self.worksheet.set_column(10, 10, 150)
        self.workbook.formats[0].set_text_wrap()
        self.worksheet.write("A1", "Use Case Name")
        self.worksheet.write("B1", "Test Case Name")
        self.worksheet.write("C1", "Scenario")
        self.worksheet.write("D1", "Use Case")
        self.worksheet.write("E1", "Test Case Title")
        self.worksheet.write("F1", "Test Case Description")
        self.worksheet.write("G1", "Expected Results")
        self.worksheet.write("H1", "Test Case Type")
        self.worksheet.write("I1", "Status")
        self.worksheet.write("J1", "Comments")
        self.worksheet.write("K1", "Reference")

    def _check_file(self):
        """Remove output file if exists"""
        if os.path.isfile(self.testcasefile):
            os.remove(self.testcasefile)

    def test_case_generator(self, url):
        """generates tests"""
        home = urlparse(url).netloc
        page = urllib2.urlopen(url)
        soup = BeautifulSoup(page, "lxml")
        soup = soup.find("body")
        self.parse_anchor_tags(soup, home)
        self.parse_button_tags(soup, home)
        self.parse_input_tags(soup, home)
        self.workbook.close()
        print("User can see generated test cases in file:", self.testcasefile)

    def write_anchor_test_case(self, link_text, case_name, div, home):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_click",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_click",
        )
        self.worksheet.write("C" + str(self.row_count), link_text)
        self.worksheet.write(
            "D" + str(self.row_count), "Validating " + link_text + " link"
        )
        self.worksheet.write(
            "E" + str(self.row_count), "[" + home + "][" + link_text + "]"
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate opening of "
            + link_text
            + " link. \nPre-requisite - User should have \
                    desired access to the "
            + home
            + " . \nTest steps: \n1. Go to "
            + home
            + " .\n2. Click on "
            + link_text
            + " link.",
        )
        self.worksheet.write(
            "G" + str(self.row_count),
            "1. " + link_text + " link should open.",
        )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

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
            self.write_anchor_test_case(link_text, case_name, div, home)
            if i > 100000:
                break

    def write_button_test_case(self, button_text, case_name, div, home):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_button_click",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_button_click",
        )
        self.worksheet.write("C" + str(self.row_count), button_text)
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + button_text + " button",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + button_text + "]",
        )
        if div.name == "input":
            self.worksheet.write(
                "F" + str(self.row_count),
                "Objective: To Validate clicking "
                + button_text
                + " button."
                + "\nPre-requisite - User should have desired access to the "
                + home
                + "."
                + "\nTest steps: "
                + "\n1. Go to "
                + home
                + "."
                + "\n2. Give relevant input in input box corresponding to "
                + button_text
                + " button."
                + "\n3. Click on "
                + button_text
                + " button.",
            )
        else:
            self.worksheet.write(
                "F" + str(self.row_count),
                "Objective: To Validate clicking "
                + button_text
                + " button."
                + "\nPre-requisite - User should have desired access to the "
                + home
                + "."
                + "\nTest steps: "
                + "\n1. Go to "
                + home
                + "."
                + "\n2. Click on "
                + button_text
                + " button.",
            )
        button_type = div.get("type")
        button_onclick = div.get("onclick")
        if button_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + button_text
                + " button click should activate respective \
                    onClick function.",
            )
        elif button_type is not None:
            if button_type.lower() == "submit":
                self.worksheet.write(
                    "G" + str(self.row_count),
                    "1. "
                    + button_text
                    + " button click should activate submit \
                        action for respective input field.",
                )
            elif button_type.lower() == "reset":
                self.worksheet.write(
                    "G" + str(self.row_count),
                    "1. "
                    + button_text
                    + " button click should reset all \
                        input fields to default.",
                )
            elif button_type.lower() == "button":
                self.worksheet.write(
                    "G" + str(self.row_count),
                    "1. " + button_text + " button should get clicked.",
                )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. " + button_text + " button click should do nothing.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

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
            self.write_button_test_case(button_text, case_name, div, home)
            if i > 100000:
                break

    def write_input_test_case(self, input_box_name, case_name, div, home):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_input_check",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_input_check",
        )
        self.worksheet.write("C" + str(self.row_count), input_box_name)
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " input box",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "] [" + input_box_name + " input]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate "
            + input_box_name
            + " input box. \nPre-requisite - "
            + "User should have desired access to the "
            + home
            + " . \nTest steps: \n1. Go to "
            + home
            + " .\n2. Click on "
            + input_box_name
            + " input box.\n3. Type relevant input "
            + "in already clicked input box.",
        )
        self.worksheet.write(
            "G" + str(self.row_count),
            "1. "
            + input_box_name
            + " input box should be clickable.\n2. "
            + input_box_name
            + " input box should reflect typed characters.",
        )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

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
                or div.get("id")
                or div.get("name")
                or div.get("aria-label")
                or div.get("aria-labelledby")
            )
            if input_box_name is None:
                input_box_name = "untitled" + str(untitled_count)
                untitled_count += 1
            case_name = input_box_name.replace(" ", "_")
            if div.get("type") == "submit":
                self.write_button_test_case(
                    input_box_name, case_name, div, home
                )
            else:
                self.write_input_test_case(
                    input_box_name, case_name, div, home
                )
            if i > 100000:
                break


@click.command(help="Provide url")
@click.option("-u", "--url", default=None, help="URL to generate test cases")
def generate(url):
    """tests"""
    tests = ManualTestCases()
    tests.test_case_generator(url)
