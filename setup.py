# coding: utf-8

from cx_Freeze import setup, Executable

base = None


executables = [Executable(r"Framework.py", base=base)]

packages = ["Common", "TestCase_Scripts"]
options = {'build_exe': { 'packages':packages }}

setup(
    name = "Framework",
    options = options,
    version = "0.0.1.2",
    description = 'An Excel Framework Demo for Kerry',
    executables = executables
)