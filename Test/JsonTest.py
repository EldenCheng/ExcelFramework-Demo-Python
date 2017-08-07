import json
from Common.CONST import *



def Json_Write(filename):
    json_file = open(filename,'w')
    json_file.write (json.dumps(data))

    json_file.close()

def Json_Read(filename):
    json_file = open(filename)

    json_data = json.load(json_file)

    print(json_data['URL'])

data = {'URL': r"http://www-uat.kerrylogistics.com/kerriervbo-demo/dispatcher/index-flow?execution=e4s1",
        'EXCELPATH': r".\TestCaseData\Kerry_data_2c.xlsx",
        'CHROMEDRIVERPATH': r".\Webdrivers\chromedriver.exe",
        'IEDRIVERPATH': r".\Webdrivers\IEDriverServer.exe",
        'TESTREPORTPATH': r".\TestReport",
        'BROWSER': "Chrome"}

Json_Write('Env.json')

Json_Read('Env.json')