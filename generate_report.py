from openpyxl import Workbook, load_workbook
from types import SimpleNamespace as Namespace
import json
import re

#   GenerateReport
# In the template can be defined path to the data in the JSON object using . delimeter e.g. prop1.prop2.prop3
# Complete array can be mapped to the result excel statrting from the cell with mapping and going down in the table using syntax prop1.arrayprop[:]
# One property from the arrays of objects can be mapped to the result excel statrting from the cell with mapping and going down in the table 
# using syntax prop1.arrayprop[:]/objectprop
def GenerateReport(Template, Data):

    # this is a prefix what defines that cell value contains path in the json object
    dataBindingPrefix = "##"

    for sheet in Template.worksheets:
        cells = sheet[sheet.dimensions]

        templateMap = {}

    # build a map with all cells what need to be replaced
        for row in cells:
            for cell in row:
                if cell.value == None or type(cell.value) is not str:
                    continue
                if cell.value[:2] != dataBindingPrefix:
                    continue
                templateMap[cell.coordinate] = cell.value[2:]

    # for all cells what need to be replaced try to find value in the data
        for key, binding in templateMap.items():

            pathProperty = binding
            arrayProperty = None
            # look if the bindgin contains mapping to some property from the array of objects 
            propDelimeter = pathProperty.find("/")
            if propDelimeter >= 0:
                arrayProperty = pathProperty[propDelimeter + 1:]
                pathProperty = pathProperty[:propDelimeter]

            # calculate the value of the expression
            try:
                expression = "Data." + pathProperty
                val = eval(expression)
            except:
                continue

            if type(val) is str or type(val) is int or type(val) is float:
                # if the calculated value is of type "simple" then we just use it
                sheet[key] = val
            elif type(val) is list:
                # if the calculated value is of type list then we need to handle it specially
                # find the starting point
                letter = re.compile("[A-Z]+").findall(key)[0]
                number = int(re.compile("[0-9]+").findall(key)[0])
                increment = 0
                for v in val:
                    if arrayProperty is not None:
                        try:
                            subArrayVal = eval("v." + arrayProperty)
                        except:
                            continue
                        v = subArrayVal
                    if type(v) is str or type(v) is int or type(v) is float:
                        sheet[letter + str(number + increment)] = v
                    increment += 1
    #-------------------------------------------------------------------------------------------------------

# load template and object
workbook = load_workbook("data\\template.xlsx", data_only=True)
with open('data\\data.json') as f:
    data = json.load(f, object_hook = lambda d: Namespace(**d))

GenerateReport(workbook, data)

# save result
workbook.save("data\\result.xlsx")
