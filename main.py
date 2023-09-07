import pandas as pd
import numpy as np
import math
from Loger import get_logger
import ProductAttribute
import SoldProductItem
import os

productAttrMap = {}
outData = []
global logger


def read_excel(path):
    raw_data = pd.read_excel(path)
    # print(raw_data)
    return raw_data


def getCodeAndSize(raw_data):
    shape = np.shape(raw_data)
    col = shape[1]
    productTypeCodeCol = -1
    productSizeStrCol = -1
    for j in range(col):
        content = str(raw_data[0][j])
        if content.startswith("产品型号"):
            productTypeCodeCol = j
            break
    for j in range(col):
        content = str(raw_data[0][j])
        if content.startswith("规格"):
            productSizeStrCol = j
            break
    # if productTypeCodeCol == -1 or productSizeStrCol == -1:
    #     logger.warning("找不到产品型号或者规格所在列!")

    return raw_data[1:, productTypeCodeCol], raw_data[1:, productSizeStrCol]


def Data2Object(raw_data):
    productCode, productSize = getCodeAndSize(raw_data)

    codeShape = np.shape(productCode)
    sizeShape = np.shape(productSize)
    if codeShape != sizeShape:
        logger.error("产品型号和规格行数不一致！")
        return
    for i in range(codeShape[0]):
        item = ProductAttribute.ProductAttr()
        item.productCode = str(productCode[i])
        item.productSizeStr = str(productSize[i])
        if item.productCode == "nan":
            continue
        size = item.productSizeStr.split("*")
        if len(size) < 3:
            continue
        item.productLength = int(size[0])
        item.productWidth = int(size[1])
        item.scale = (float(size[0]) / 1000.0) * (float(size[1]) / 1000.0)
        if productAttrMap.get(item.productCode) is not None:
            it = productAttrMap.get(item.productCode)
            if it.ProductAttr.productSizeStr is not item.productSizeStr:
                logger.error("请注意这个产品代码有重复，并且规格不一致，请人工检查！")
                logger.error("产品型号: %s", item.productCode)
                logger.error("产品规格： %s", item.productSizeStr)
        productAttrMap[item.productCode] = item


def getSoldCodeAndSoldNum(raw_data):
    shape = np.shape(raw_data)
    col = shape[1]
    productTypeCol = -1
    SoldNumCol = -1
    soldNum = -1
    for j in range(col):
        content = str(raw_data[0][j])
        if content.startswith("型号"):
            productTypeCol = j
            break
    for j in range(col):
        content = str(raw_data[0][j])
        if content.startswith("入库数量"):
            SoldNumCol = j
            break
    if productTypeCol == -1 or SoldNumCol == -1:
        logger.warning("找不到型号或者入库数量所在列!")
    return raw_data[1:, productTypeCol], raw_data[1:, SoldNumCol]


def decodeType(productType):
    str_len = len(productType)
    code = ""
    for i in range(str_len):
        c = productType[i]
        if c.isdigit() or c.encode().isalpha() or c == '-':
            if c.isdigit() or c.encode().isalpha():
                code += c
            elif 0 < i < str_len - 1:
                before = productType[i - 1]
                after = productType[i + 1]
                if (before.isdigit() or before.encode().isalpha()) and (after.isdigit() or after.encode().isalpha()):
                    code += c
        else:
            if len(code) > 2:
                break
    return code


def generateSoldInfo(raw_data):
    product_type, sold_num = getSoldCodeAndSoldNum(raw_data)

    typeShape = np.shape(product_type)
    soldShape = np.shape(sold_num)
    if typeShape != soldShape:
        logger.error("型号和入库数量行数不一致！")
        return
    for i in range(typeShape[0]):
        item = SoldProductItem.SoldProductItem()
        item.productType = str(product_type[i])
        if not math.isnan(sold_num[i]):
            item.sold_num = int(sold_num[i])
        code = decodeType(item.productType)
        if code == "":
            logger.error("无法解析出产品代码！请手动处理！产品类型： %s", item.productType)
        else:
            item.productCode = code
            if productAttrMap.get(code) is None:
                keys = productAttrMap.keys()
                isFound = False
                for key in keys:
                    key = str(key)
                    if key.endswith(code):
                        item.scale = productAttrMap.get(key).scale
                        isFound = True
                        break
                if not isFound:
                    logger.error("没有对应的产品数据，请手动处理！产品类型: %s, 产品代码：%s", item.productType, code)
            else:
                item.scale = productAttrMap.get(code).scale
        item.sold_scale_total = item.scale * item.sold_num
        outData.append(item)
        # print(round(item.scale, 8))


def generateProductMap(path):
    excel_file = pd.ExcelFile(path)

    for file in excel_file.sheet_names:
        raw_data = np.array(pd.read_excel(path, file))
        Data2Object(raw_data)


def getAllExcels():
    map_files = []
    data_files = []
    for root, dirs, files in os.walk("./", topdown=True):
        for file in files:
            file = str(file)
            if file.endswith(".xlsx"):
                if file.find("产品清单") > -1:
                    map_files.append(file)
                elif file.find("暂估单") > -1:
                    data_files.append(file)
    return map_files, data_files


if __name__ == '__main__':
    logger = get_logger()
    map_files, data_files = getAllExcels()
    if len(map_files) < 1 or len(data_files) < 1:
        logger.error("文件缺失，产品清单和暂估单都要大于1个！")

    for map_file in map_files:
        generateProductMap(map_file)


    # product_size_path = "金螳螂新增产品清单-联丰2023-7.xlsx"
    # product_size_path2 = "金螳螂新增产品清单-联丰2023-3-30最终定稿版本.xlsx"
    predict_sheet = data_files[0]#"应付暂估单_202308280127531369474(1)(1).xlsx"
    # product_size_raw = np.array(read_excel(product_size_path))
    # generateProductMap(product_size_path2)
    predict_price_excel = read_excel(predict_sheet)
    predict_price_raw = np.array(predict_price_excel)
    # Data2Object(product_size_raw)
    generateSoldInfo(predict_price_raw)
    scale_list = ["单片面积"]
    sold_scale_list = ["售出总面积"]

    for item in outData:
        scale_list.append(round(item.scale, 9))
        sold_scale_list.append(round(item.sold_scale_total, 9))
    predict_price_excel["面积统计"] = scale_list
    predict_price_excel["总面积统计"] = sold_scale_list
    predict_price_excel.to_excel("out.xlsx", index=False)
    print("程序运行结束！请按回车键退出.")
    input()