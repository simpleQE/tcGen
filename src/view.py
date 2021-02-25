import os

import xlsxwriter


class TestsSheet(object):
    """Class to generate test cases and its structure"""

    def __init__(self):
        self.row_count = 0
        self.config_path = os.path.expanduser("~")
        self.testcasefile = os.path.join(self.config_path, "output.xlsx")

        self._check_file()
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

    def test_case_show(self):
        """show generated tests"""
        self.workbook.close()
        print("User can see generated test cases in file:", self.testcasefile)

    def write_anchor_test_case(
        self, link_text, case_name, div, home, rows_count
    ):
        self.row_count = rows_count
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

    def write_button_test_case(
        self, button_text, case_name, div, home, rows_count
    ):
        self.row_count = rows_count

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

    def write_input_submit_test_case(
        self, input_box_name, case_name, div, home
    ):
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
        self.worksheet.write("C" + str(self.row_count), input_box_name)
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " button",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate clicking "
            + input_box_name
            + " button."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Give relevant input in input box corresponding to "
            + input_box_name
            + " button."
            + "\n3. Click on "
            + input_box_name
            + " button.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " button click should activate respective \
                    onClick function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " button click should activate submit \
                    action for respective input field.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_button_test_case(
        self, input_box_name, case_name, div, home
    ):
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
        self.worksheet.write("C" + str(self.row_count), input_box_name)
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " button",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate clicking "
            + input_box_name
            + " button."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Do relevant action corresponding to "
            + input_box_name
            + " button, if any."
            + "\n3. Click on "
            + input_box_name
            + " button.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " button click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " button click should activate relevant action.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_checkbox_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_checkbox_click",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_checkbox_click",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " checkbox"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " checkbox",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate clicking "
            + input_box_name
            + " checkbox."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Click on "
            + input_box_name
            + " checkbox.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " checkbox click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " checkbox should be clickable."
                + "\n2. It should activate relevant action, if any.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_color_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_color_selection_click",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_color_selection_click",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " color selection"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " color selection",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate clicking "
            + input_box_name
            + " color selection input."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Click on "
            + input_box_name
            + " color selection.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " color selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " color selection should be clickable."
                + "\n2. It should open color pallete."
                + "\n3. User should be able to select color"
                + " from opened color pallete.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_date_test_case(self, input_box_name, case_name, div, home):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_date_selection_click",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_date_selection_click",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " date selection"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " date selection",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate clicking "
            + input_box_name
            + " date selection input."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Click on "
            + input_box_name
            + " date selection.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " date selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " date selection should be clickable."
                + "\n2. It should open date dropdown with calendar."
                + "\n3. User should be able to select date"
                + " from opened date dropdown.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_datetime_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_datetime_selection_click",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_datetime_selection_click",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " datetime selection"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " datetime selection",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate clicking "
            + input_box_name
            + " datetime selection input."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Click on "
            + input_box_name
            + " datetime selection.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " datetime selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " datetime selection should be clickable."
                + "\n2. It should open datetime dropdown with calendar."
                + "\n3. User should be able to select date and time"
                + " from opened datetime dropdown.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_email_test_case(
        self, input_box_name, case_name, div, home
    ):
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
            + " input box.\n3. Type email id "
            + "in already clicked input box.",
        )
        self.worksheet.write(
            "G" + str(self.row_count),
            "1. "
            + input_box_name
            + " input box should be clickable.\n2. "
            + input_box_name
            + " input box should reflect typed characters of email id.",
        )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_file_test_case(self, input_box_name, case_name, div, home):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_input_file_check",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_input_file_check",
        )
        self.worksheet.write("C" + str(self.row_count), input_box_name)
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " file input button",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "] [" + input_box_name + " file input]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate "
            + input_box_name
            + "file input button."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Click on "
            + input_box_name
            + " input button."
            + "\n3. Choose file to upload from file manager prompt.",
        )
        self.worksheet.write(
            "G" + str(self.row_count),
            "1. "
            + input_box_name
            + " file input button should be clickable."
            + "\n2. "
            + input_box_name
            + " input button should open a file manager prompt to select file."
            + "\n3. File should get uploaded and "
            + "it should indicate file uploaded successfully or not.",
        )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_image_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_image_input_check",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_image_input_check",
        )
        self.worksheet.write("C" + str(self.row_count), input_box_name)
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " input image",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "] [" + input_box_name + " image]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate "
            + input_box_name
            + " image button."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Click on "
            + input_box_name
            + " image button.",
        )
        self.worksheet.write(
            "G" + str(self.row_count),
            "1. "
            + input_box_name
            + " image button should be clickable."
            + "\n2. "
            + input_box_name
            + " input button should activate click action"
            + " corresponding to it.",
        )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_month_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_month_selection_click",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_month_selection_click",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " month selection"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " month selection",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate clicking "
            + input_box_name
            + " month selection input."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Click on "
            + input_box_name
            + " month selection.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " month selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " month selection should be clickable."
                + "\n2. It should open month dropdown."
                + "\n3. User should be able to select month"
                + " from opened month dropdown.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_number_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_number_input_check",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_number_input_check",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " number selection"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " number selection",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate clicking "
            + input_box_name
            + " number selection input."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Click on up/down arrow in"
            + input_box_name
            + " number selection and set desired number.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " number selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " number input should be clickable."
                + "\n2. User should be able to select desired number"
                + " using up/down arrows.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_password_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_password_input_check",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_password_input_check",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " password input"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " password input",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate "
            + input_box_name
            + " password input box. \nPre-requisite - "
            + "User should have desired access to the "
            + home
            + " . \nTest steps: \n1. Go to "
            + home
            + " .\n2. Click on "
            + input_box_name
            + " password input box.\n3. Type password "
            + "in already clicked input box.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " password selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " password input box should be clickable.\n2. "
                + input_box_name
                + " password input box should reflect "
                + "typed characters of password."
                + "\n 3. typed characters should be masked.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_radio_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_radio_input_check",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_radio_input_check",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " radio input"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " radio input",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate "
            + input_box_name
            + " radio input. \nPre-requisite - "
            + "User should have desired access to the "
            + home
            + " . \nTest steps: \n1. Go to "
            + home
            + " .\n2. Click on "
            + input_box_name
            + " radio input.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " radio selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " radio input button should be clickable."
                + "\n2. User should able to select corresponding input.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_range_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_range_input_check",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_range_input_check",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " range input"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " range input",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate "
            + input_box_name
            + " range input. \nPre-requisite - "
            + "User should have desired access to the "
            + home
            + " . \nTest steps: \n1. Go to "
            + home
            + " .\n2. Click, hold and swipe "
            + input_box_name
            + " range input.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " range selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " range input button should be clickable."
                + "\n2. User should able to select corresponding input.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_reset_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_reset_button_click",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_reset_button_click",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " reset button"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " reset button",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate "
            + input_box_name
            + " reset button. \nPre-requisite - "
            + "User should have desired access to the "
            + home
            + " . \nTest steps: \n1. Go to "
            + home
            + " ."
            + "\n2. Provide inputs in input boxes corresponding"
            + " to reset button."
            + "\n3. Click on "
            + input_box_name
            + " reset button.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " reset selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " reset button should be clickable."
                + "\n2. All corresponding inputs should reset after "
                + "clicking on reset button.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_search_test_case(
        self, input_box_name, case_name, div, home
    ):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_search_input_check",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_search_input_check",
        )
        self.worksheet.write("C" + str(self.row_count), input_box_name)
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " search input box",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "] [" + input_box_name + " search input]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate "
            + input_box_name
            + " search input box. \nPre-requisite - "
            + "User should have desired access to the "
            + home
            + " . \nTest steps: \n1. Go to "
            + home
            + " .\n2. Click on "
            + input_box_name
            + " search input box.\n3. Type relevant input "
            + "in already clicked search input box.",
        )
        self.worksheet.write(
            "G" + str(self.row_count),
            "1. "
            + input_box_name
            + " search input box should be clickable.\n2. "
            + input_box_name
            + " search input box should reflect typed characters of email id."
            + "\n3. "
            + "search input box should show suggestions dropdown.",
        )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_tel_test_case(self, input_box_name, case_name, div, home):
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
            + " input box.\n3. Type telephone number "
            + "in already clicked input box.",
        )
        self.worksheet.write(
            "G" + str(self.row_count),
            "1. "
            + input_box_name
            + " input box should be clickable.\n2. "
            + input_box_name
            + " input box should reflect "
            + "typed characters of telephone number.",
        )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_time_test_case(self, input_box_name, case_name, div, home):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_time_selection_click",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_time_selection_click",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " time selection"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " time selection",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate clicking "
            + input_box_name
            + " time selection input."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Click on "
            + input_box_name
            + " time selection.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " time selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " time selection should be clickable."
                + "\n2. It should open time dropdown."
                + "\n3. User should be able to select time"
                + " from opened time dropdown.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_url_test_case(self, input_box_name, case_name, div, home):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_url_input_check",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_url_input_check",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " url input"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " url input",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate "
            + input_box_name
            + " url input box. \nPre-requisite - "
            + "User should have desired access to the "
            + home
            + " . \nTest steps: \n1. Go to "
            + home
            + " .\n2. Click on "
            + input_box_name
            + " url input box.\n3. Type url "
            + "in already clicked input box.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " url selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " url input box should be clickable.\n2. "
                + input_box_name
                + " url input box should reflect "
                + "typed characters of url."
                + "\n 3. Invalid url input should show invalid message.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))

    def write_input_week_test_case(self, input_box_name, case_name, div, home):
        self.worksheet.write(
            "A" + str(self.row_count),
            "UC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_week_selection_click",
        )
        self.worksheet.write(
            "B" + str(self.row_count),
            "TC"
            + str(self.row_count - 1)
            + "_"
            + case_name.lower()
            + "_week_selection_click",
        )
        self.worksheet.write(
            "C" + str(self.row_count), input_box_name + " week selection"
        )
        self.worksheet.write(
            "D" + str(self.row_count),
            "Validating " + input_box_name + " week selection",
        )
        self.worksheet.write(
            "E" + str(self.row_count),
            "[" + home + "][" + input_box_name + "]",
        )
        self.worksheet.write(
            "F" + str(self.row_count),
            "Objective: To Validate clicking "
            + input_box_name
            + " week selection input."
            + "\nPre-requisite - User should have desired access to the "
            + home
            + "."
            + "\nTest steps: "
            + "\n1. Go to "
            + home
            + "."
            + "\n2. Click on "
            + input_box_name
            + " week selection.",
        )
        input_onclick = div.get("onclick")
        if input_onclick is not None:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " week selection click should activate respective "
                + input_onclick
                + " function.",
            )
        else:
            self.worksheet.write(
                "G" + str(self.row_count),
                "1. "
                + input_box_name
                + " week selection should be clickable."
                + "\n2. It should open week dropdown."
                + "\n3. User should be able to select week"
                + " from opened week dropdown.",
            )
        self.worksheet.write("H" + str(self.row_count), "Smoke")
        self.worksheet.write_string("K" + str(self.row_count), str(div))
