import os
from typing import List
from xml.dom import minidom
from datetime import datetime
import openpyxl


class TestCaseDetailsParser:
    """
    Class to extract test case details from XML file and write in excel file(.xlsx).

    """

    def __init__(self, file_name: str) -> None:
        """
        Initialization for the TestCaseDetailsParser class

        Parameters
        ----------
        file_name: str
            xml file to parse

        """
        self.file_name = file_name
        self.file_path = os.path.join(
            os.getcwd().rsplit("\\", 1)[0], f"inputs\{file_name}"
        )

    def parse_xml(self) -> List:
        """
        Method to parse xml file to extract test case details.


        Returns
        -------
        `List`
            Data list consists of test case information.

        """
        xml_obj = minidom.parse(self.file_path)
        info_type_list = ["header", "uut", "test_information"]
        data_list = []
        for info_type in info_type_list:
            tag_name = xml_obj.getElementsByTagName(info_type)
            for node in tag_name:
                test_case_details_list = node.getElementsByTagName("br")
                for i in range(0, len(test_case_details_list), 2):
                    try:
                        data_list.append(
                            [
                                info_type,
                                test_case_details_list[i].childNodes[0].nodeValue,
                                test_case_details_list[i + 1].childNodes[0].nodeValue,
                            ]
                        )
                    except IndexError:
                        data_list.append(
                            [
                                info_type,
                                test_case_details_list[i].childNodes[0].nodeValue,
                                "' '",
                            ]
                        )
        self._write_to_excel(data_list)

    def _write_to_excel(self, data_list: List) -> None:
        """
        Method to write the test case details in the excel file(.xlsx).

        Parameters
        ----------
        `data_list`
            Data list consists of test case information.

        """
        wb = openpyxl.Workbook()
        sheet = wb.active
        row = 1
        for each_data_list in data_list:
            sheet.cell(row=row, column=1, value=each_data_list[0])
            sheet.cell(row=row, column=2, value=each_data_list[1])
            sheet.cell(row=row, column=3, value=each_data_list[2])
            row += 1
        time = datetime.now().strftime("%Y-%m-%d-%H%M%S")
        wb.save(
            os.path.join(os.getcwd().rsplit("\\", 1)[0], "outputs\extracted_xml{time}.xlsx")
        )
