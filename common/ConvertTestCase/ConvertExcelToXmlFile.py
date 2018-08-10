from typing import Any

import xlrd
from lxml import etree


def convertToXml(excelFileName, xmlFileName, directoryInputFile, sheetId):
    # root = etree.Element('testcases')
    # root.set('name', 'Network')
    # tree = etree.ElementTree(root)
    # name = etree.Element('nodes')
    # root.append(name)
    wb = xlrd.open_workbook(excelFileName)

    sh = wb.sheet_by_index(sheetId)

    # Lấy dữ liệu cột Title
    nameList = sh.col_values(colx=1, start_rowx=0, end_rowx=113)
    # Lấy dữ liệu cột Pre-condition
    preConditionList = sh.col_values(colx=2, start_rowx=0, end_rowx=113)
    # Lấy dữ liệu cột Action
    actionList = sh.col_values(colx=3, start_rowx=0, end_rowx=113)
    # Lấy dữ liệu cột Expected Result
    expectedResultList = sh.col_values(colx=4, start_rowx=0, end_rowx=113)
    # Lấy dữ liệu cột Excuted Type
    excutedTypeList = sh.col_values(colx=7, start_rowx=0, end_rowx=113)
    # Lấy dữ liệu cột Priority
    priorityList = sh.col_values(colx=6, start_rowx=0, end_rowx=113)

    numCases = []
    for i in range(len(actionList)):
        if (actionList[i].startswith("1.")):
            numCases.append(i)

    root = etree.Element('testcases')
    for i in range(len(numCases)):
        index = numCases[i]
        name = nameList[index]
        print(index)
        testcase = etree.SubElement(root, 'testcase')
        testcase.set('name', name)
        testcase.set('internalid', '')

        data = {
            'node_order': '',
            'externalid': '',
            # ...
        }

        for k, v in data.items():
            n = etree.SubElement(testcase, k)
            n.text = etree.CDATA(v)

        node_order = etree.SubElement(testcase, 'node_order')
        node_order.text = etree.CDATA('')

        externalid = etree.SubElement(testcase, 'externalid')
        externalid.text = etree.CDATA('')

        version = etree.SubElement(testcase ,'version')
        version.text = etree.CDATA('')

        summary = etree.SubElement(testcase,'summary')
        summary.text = etree.CDATA('')

        preconditions = etree.SubElement(testcase,'preconditions')
        preconditions.text = etree.CDATA(preConditionList[index])

        execution_type = etree.SubElement(testcase,'execution_type')
        execution_type.text = etree.CDATA(excutedTypeList[index])

        importance = etree.SubElement(testcase, 'importance')
        importance.text = etree.CDATA(priorityList[index])

        steps = etree.SubElement(testcase, 'steps')
        stepNum = 1
        if (i < len(numCases) - 1):
            TestCaseFormat.createSteps(actionList, expectedResultList, excutedTypeList, steps, stepNum, index, numCases[i+1])
            # for j in range(index, numCases[i+1]):
            #     step = etree.SubElement(steps, 'step')
            #     step_number =  etree.SubElement(step, 'step_number')
            #     step_number.text = etree.CDATA(str(stepNum))
            #     actions = etree.SubElement(step, 'actions')
            #     actions.text = etree.CDATA(actionList[j])
            #     expectedresults = etree.SubElement(step, 'expectedresults')
            #     expectedresults.text = etree.CDATA(expectedResultList[j])
            #     execution_type = etree.SubElement(step, 'execution_type')
            #     execution_type.text = etree.CDATA(excutedTypeList[j])
            #     stepNum +=1
        else:
            TestCaseFormat.createSteps(actionList, expectedResultList, excutedTypeList, steps,stepNum, index, len(actionList) )
            # for j in range (index, len(actionList)):
            #     step = etree.SubElement(steps, 'step')
            #     step_number = etree.SubElement(step, 'step_number')
            #     step_number.text = etree.CDATA(str(stepNum))
            #     actions = etree.SubElement(step, 'actions')
            #     actions.text = etree.CDATA(actionList[j])
            #     expectedresults = etree.SubElement(step, 'expectedresults')
            #     expectedresults.text = etree.CDATA(expectedResultList[j])
            #     execution_type = etree.SubElement(step, 'execution_type')
            #     execution_type.text = etree.CDATA(excutedTypeList[j])
            #     stepNum += 1
    print(etree.tostring(root))


# for row in range(1, sh.nrows):
#     val = sh.row_values(row)
#
#     element = etree.SubElement(name, 'node')
#     element.set('id', str(val[0]))
#     element.set('x', str(val[1]))
#     element.set('y', str(val[2]))
# print(etree.tostring(root, pretty_print=True))
# xml = open(xmlFileName, "w")
# xml.write(etree.tostring(root, pretty_print=True))

class TestCaseFormat:

    def __init__(self) -> None:
        super().__init__()
        self.node_oder = etree.Element('node_order')
        self.node_oder = etree.CDATA
        self.externalid = etree.Element('externalid')
        self.externalid = etree.CDATA

    def __set_node_order__(self, value: str) -> None:
        self.node_oder.text = etree.CDATA(value)

    @staticmethod
    def createSteps(actionList, expectedResultList,excutedTypeList, steps, stepNum, startIndex, endIndex):
        for j in range(startIndex, endIndex):
            step = etree.SubElement(steps, 'step')
            step_number = etree.SubElement(step, 'step_number')
            step_number.text = etree.CDATA(str(stepNum))
            actions = etree.SubElement(step, 'actions')
            actions.text = etree.CDATA(actionList[j])
            expectedresults = etree.SubElement(step, 'expectedresults')
            expectedresults.text = etree.CDATA(expectedResultList[j])
            execution_type = etree.SubElement(step, 'execution_type')
            execution_type.text = etree.CDATA(excutedTypeList[j])
            stepNum += 1


def main():
    print("Start converting....")
    directoryInputFile = ""
    excelFileName = "Test.xlsx"
    xmlFileName = "test.xml"
    sheetId = 0
    convertToXml(excelFileName, xmlFileName, directoryInputFile, sheetId)
    print("Finished converting .....")


if __name__ == '__main__':
    main()
