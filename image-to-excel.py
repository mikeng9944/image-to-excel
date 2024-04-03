import os
from azure.ai.vision.imageanalysis import ImageAnalysisClient
from azure.ai.vision.imageanalysis.models import VisualFeatures
from azure.core.credentials import AzureKeyCredential

from openpyxl import Workbook

import re

# Set the values of your computer vision endpoint and computer vision key
# as environment variables:
try:
    endpoint = "https://mikeai.cognitiveservices.azure.com/"
    key = "73b20d0d33ee42ada24eb120e0f252ac"
except KeyError:
    print("Missing environment variable 'VISION_ENDPOINT' or 'VISION_KEY'")
    print("Set them before running this sample.")
    exit()

# Create an Image Analysis client
client = ImageAnalysisClient(
    endpoint=endpoint,
    credential=AzureKeyCredential(key)
)

with open("account_data.png", "rb") as f:
    image = f.read()

# Get a caption for the image. This will be a synchronously (blocking) call.
result = client.analyze(
    image_data=image,
    visual_features=[VisualFeatures.CAPTION, VisualFeatures.READ],
    gender_neutral_caption=True,  # Optional (default is False)
)

wb = Workbook()

ws1 = wb["Sheet"]
ws1["A1"] = "Bank"
ws1["B1"] = "Invoice Date"
ws1["C1"] = "Receipt Date"
ws1["D1"] = "Invoice No"
ws1["E1"] = "Customer Name"
ws1["F1"] = "Purpose"
ws1["G1"] = "In"
ws1["H1"] = "Out"
ws1["I1"] = "Amount"

row_start = 2

startRecord = False
firstRow = False
secondDate = False
date_pattern = re.compile(r'\d{4}/\d{2}/\d{2}')
amount_pattern = re.compile(r'\d*\.\d{2}')

totalAmount = 0
amount_calculation = [0.0, 0.0]
storeCount = 0

if result.read is not None:
    for line in result.read.blocks[0].lines:
        if (line.text == "原幣結餘/(結欠)"):
            startRecord = True
            firstRow == True
            continue
        elif (startRecord == True and firstRow == True):
            if (date_pattern.search(line.text)):
                ws1.cell(row_start, 2).value = line.text
                continue
            elif (line.text == "承前結餘"):
                ws1.cell(row_start, 6).value = line.text
                continue
            elif (amount_pattern.search(line.text)):
                totalAmount = float(line.text.replace(",", ""))
                ws1.cell(row_start, 9).value = totalAmount
                firstRow = False
                row_start += 1
                continue
        elif (startRecord == True and firstRow == False):
            if (date_pattern.search(line.text) and secondDate == False):
                ws1.cell(row_start, 2).value = line.text
                secondDate = True
                continue
            elif (date_pattern.search(line.text) and secondDate == True):
                ws1.cell(row_start, 3).value = line.text
                secondDate = False
                continue
            elif (amount_pattern.search(line.text)):
                amount_calculation[storeCount] = float(line.text.replace(",", ""))
                if (storeCount == 0):
                    storeCount += 1
                    continue
                else:
                    amount_calculation.sort()
                    print(str(amount_calculation[0]) + ", " + str(amount_calculation[1]))
                    if (amount_calculation[0] + amount_calculation[1] == totalAmount):
                        ws1.cell(row_start, 7).value = amount_calculation[0]
                    else:
                        ws1.cell(row_start, 8).value = amount_calculation[0]
                    ws1.cell(row_start, 9).value = amount_calculation[1]
                    totalAmount = amount_calculation[1]
                    storeCount = 0
                    row_start += 1
                    continue

wb.save("book_eg.xlsx")
        # print(f"   Line: '{line.text}'")