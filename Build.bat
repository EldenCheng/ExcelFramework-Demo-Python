@echo off

set /p bdno=Please input the current build No. : 

python setup.py build

for /f %%i in ('dir /ad/b ".\build"') do set yourvar=%%i

cd .\build

md TestCaseData

md TestReport

md WebDrivers

copy ..\Env.json .\

copy ..\Common\RunTest.bat .\

copy ..\TestCaseData .\TestCaseData\

copy ..\TestReport .\TestReport\

copy ..\WebDrivers .\WebDrivers\

ren %yourvar% Bin

cd..

ren build KerryDemo_Build(%bdno%)_x64