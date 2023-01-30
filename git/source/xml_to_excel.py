from xml.dom import minidom
import openpyxl
p1 = minidom.parse(r"C:\Users\z0175000\Downloads\scenario_template_2023-01-25_15h26m46s (1).xml")
wb = openpyxl.Workbook()
sheet = wb.active
info_type_list = ["header","uut","test_information"]
for info_type in info_type_list:
    tag_name = p1.getElementsByTagName(info_type)
    #print(tag_name)
    for node in tag_name:
        alist = node.getElementsByTagName('br')
        for i in range(0, len(alist), 2):
            try:
                sheet.append([info_type, alist[i].childNodes[0].nodeValue, alist[i + 1].childNodes[0].nodeValue])
                print(alist[i].childNodes[0].nodeValue, ':', alist[i+1].childNodes[0].nodeValue)
            except IndexError:
                sheet.append([info_type,alist[i].childNodes[0].nodeValue, "''"])
wb.save("exmpl_xml.xlsx")

