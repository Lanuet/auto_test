import codecs
from typing import Any

import xlrd
from lxml import etree

start_row = 0
end_row = 113
sheet_id = 0
title_col = 1
preconditon_col = 2
action_col = 3
expected_result_col = 4
priority_col = 6
excutedtype_col = 7


def convertToXml(excelFileName, xmlFileName, directoryInputFile, sheetId):
    xml_file = directoryInputFile + xmlFileName
    # file = codecs.open(xml_file,"w")
    # Open ExcelFile
    wb = xlrd.open_workbook(excelFileName)
    # Read by sheet index
    sh = wb.sheet_by_index(sheetId)

    # Lấy dữ liệu cột Title
    name_list = sh.col_values(title_col, start_row, end_row)
    # Lấy dữ liệu cột Pre-preconditon_
    pre_condition_list = sh.col_values(preconditon_col, start_row, end_row)
    # Lấy dữ liệu cột Action
    action_list = sh.col_values(action_col, start_row, end_row)
    # Lấy dữ liệu cột Expected Result
    expected_result_list = sh.col_values(expected_result_col, start_row, end_row)
    # Lấy dữ liệu cột Excuted Type
    excuted_type_list = sh.col_values(excutedtype_col, start_row, end_row)
    # Lấy dữ liệu cột Priority
    priority_list = sh.col_values(priority_col, start_row, end_row)

    num_cases = []
    for i in range(len(action_list)):
        if action_list[i].startswith("1."):
            num_cases.append(i)

    root = etree.Element('testcases')
    for i in range(len(num_cases)):
        index = num_cases[i]
        step_num = 1
        name = name_list[index]
        print(index)
        testcase = etree.SubElement(root, 'testcase')
        testcase.set('name', name)
        testcase.set('internalid', '')

        # data = {
        #     'node_order': '',
        #     'externalid': '',
        #     'version': '',
        #     'summary': '',
        #     'preconditions': pre_condition_list[index],
        #     'execution_type': excuted_type_list[index],
        #     'importance': priority_list[index],
        #     'steps': steps_data
        # }
        # steps_data = {
        #     'step_number': step_num,
        #     'actions': ''

        # }
        # for k, v in data.items():
        #     n = etree.SubElement(testcase, k)
        #     n.text = etree.CDATA(v)

        node_order = etree.SubElement(testcase, 'node_order')
        node_order.text = etree.CDATA('')

        externalid = etree.SubElement(testcase, 'externalid')
        externalid.text = etree.CDATA('')

        version = etree.SubElement(testcase, 'version')
        version.text = etree.CDATA('')

        summary = etree.SubElement(testcase, 'summary')
        summary.text = etree.CDATA('')

        preconditions = etree.SubElement(testcase, 'preconditions')
        preconditions.text = etree.CDATA(pre_condition_list[index])

        execution_type = etree.SubElement(testcase, 'execution_type')
        execution_type.text = etree.CDATA(excuted_type_list[index])

        importance = etree.SubElement(testcase, 'importance')
        importance.text = etree.CDATA(priority_list[index])

        steps = etree.SubElement(testcase, 'steps')

        if i < len(num_cases) - 1:
            create_steps(action_list, expected_result_list, excuted_type_list, steps, step_num, index, num_cases[i + 1])
        else:
            create_steps(action_list, expected_result_list, excuted_type_list, steps, step_num, index, len(action_list))
    # Write to file
    tree = etree.ElementTree(root)
    tree.write(xml_file)


def create_steps(action_list, expected_result_list, excuted_type_list, steps, step_num, start_index, end_index):
    """

    :param action_list: list action of each test case
    :param expected_result_list:
    :type excuted_type_list: object
    :param steps:
    :param step_num:
    :param start_index:
    :param end_index:
    """
    for j in range(start_index, end_index):
        step = etree.SubElement(steps, 'step')
        step_number = etree.SubElement(step, 'step_number')
        step_number.text = etree.CDATA(str(step_number))
        actions = etree.SubElement(step, 'actions')
        actions.text = etree.CDATA(action_list[j])
        expectedresults = etree.SubElement(step, 'expectedresults')
        expectedresults.text = etree.CDATA(expected_result_list[j])
        execution_type = etree.SubElement(step, 'execution_type')
        execution_type.text = etree.CDATA(excuted_type_list[j])
        step_num += 1


def main():
    print("Start converting....")
    directory_input_file = "C:/Users/Lannt/Data logistic/"
    excel_file_name = "Test.xlsx"
    xml_file_name = "test.xml"
    convertToXml(excel_file_name, xml_file_name, directory_input_file, sheet_id)
    print("Finished converting .....")


if __name__ == '__main__':
    main()
