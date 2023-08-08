import pandas as pd
import json
from datetime import datetime
import os
import openpyxl


def convert_to_string(value):
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d %H:%M:%S')
    elif isinstance(value, float):
        return round(value, 4)
    else:
        return value


def getSheet(fileName, sheetName, key):
    if key == 'col':
        data = pd.read_excel(fileName, sheet_name=sheetName)
        data = data.applymap(convert_to_string)
        list = data.to_dict(orient='records')
    else:
        data = pd.read_excel(fileName, sheet_name=sheetName, header=None)
        data.reset_index(col_level=0, drop=True)
        data = data.applymap(convert_to_string)
        rowNames = data.iloc[:, 0].tolist()
        ncols = data.count(axis="columns", numeric_only=False)
        list = []
        for i in range(1, ncols[0]):
            colContent = data.iloc[:, i].tolist()
            if colContent:
                app = {}
                for index in range(len(rowNames)):
                    app[rowNames[index]] = colContent[index]
                list.append(app)
    return list


#取一个产品
def getProduct(filepath):
    key = ['row', 'col', 'row', 'col']  # 代表每张工作表是以行还是列名为关键字
    file = pd.ExcelFile(filepath)
    sheetNames = file.sheet_names
    app={}
    for i in range(len(sheetNames)):
        app[sheetNames[i]] = getSheet(file, sheetNames[i], key[i])
    return app

if __name__ == '__main__':
    folderName = ['多策略', '相对价值', '市场中性']
    jsonName = "test2.json"
    list=[]
    app={}
    for foldername in folderName:
        app["name"] = foldername
        filepath = r'' + foldername + '\\'
        app['value'] = []
        productNameList = os.listdir(filepath)
        for name in productNameList:
            app['value'].append(getProduct(filepath + name))
        list.append(app)
        app={}

    with open(jsonName, 'w', encoding='utf-8') as f:
        f.write(json.dumps(list, ensure_ascii=False))



