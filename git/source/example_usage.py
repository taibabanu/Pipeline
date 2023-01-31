from xml_to_excel import TestCaseDetailsParser


test_case_details_parser_obj = TestCaseDetailsParser(
    "scenario_template_2023-01-25_15h26m46s.xml"
)
test_case_details_parser_obj.parse_xml()
