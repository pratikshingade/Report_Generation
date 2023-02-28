FONT = "Calibri"
TITLE_FONTSIZE = 18
HEADER_FONTSIZE = 12
FONTSIZE = 11
ROW_HEIGHT = 17
COLUMN_WIDTH = 23
BOLD_HEADER = True
TITLE_ROW_HEIGHT = 24

BORDER_STYLE = "thin"
CENTER_ALIGNMENT = "center"
CUSTOM_PAGE_SCALE = 102  # in percentage

HEADER_ROW_NUMBERS = list(range(1, 10))

TALUKA_NAME_COLUMN_WIDTH = 19
VILLAGE_NAME_COLUMN_WIDTH = 17.5

PDF_GENERATION_APPLICATION = "Excel.Application"
PDF_FORMAT = ".pdf"

COLUMN_MAP = {
    1: "A",
    2: "B",
    3: "C",
    4: "D",
    5: "E",
    6: "F",
    7: "G",
    8: "H",
    9: "I",
    10: "J",
    11: "K",
}
#     12: "L",
#     13: "M",
#     14: "N",
#     15: "O",
# }

COLUMN_NUMBERS = [v for v in COLUMN_MAP.keys()]

PRINT_MARGINS = {
    "LEFT": 0.25,
    "RIGHT": 0.25,
    "TOP": 0.75,
    "BOTTOM": 0.75,
    "HEADER": 0.29,
    "FOOTER": 0.29
}

HEADER_TALUKA_NAME = {
    "K": "कवठे महांकाळ",
    "T": "तासगांव",
    "M": "मिरज"
}

HEADER_CANAL_NAME = {
    "KVT-1-39": ["कवठे महांकाळ कालवा कि मी १ ते कि मी ३९", "म्हैशाळ पंपगृह उपविभाग क्र. ५, खंडेराजुरी"],
    "KVT-40-56": ["कवठे महांकाळ कालवा कि मी ४० ते कि मी ५६", "म्हैशाळ उपसासिंचन व्यवस्थापन उपविभाग क्र. २, बेळंकी"],
    "KVT-56-73": ["कवठे महांकाळ कालवा कि मी ५६ ते कि मी ७३", "म्हैशाळ भुविकास उपविभाग क्र. २, म्हैशाळ"],
    "SBC-11-30": ["सलगरे शाखा कालवा कि मी ११ ते कि मी ३०", "म्हैशाळ पंपगृह उपविभाग क्र. ५, खंडेराजुरी"],
    "BAN": ["बनेवाडी उपसा सिंचन योजना", "म्हैशाळ पंपगृह उपविभाग क्र. ५, खंडेराजुरी"],
    "AGC-1-7": ["आरग कालवा कि मी १ ते कि मी ७", "म्हैशाळ पंपगृह उपविभाग क्र. २, म्हैशाळ"]
}
