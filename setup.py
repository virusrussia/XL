'''
Created on 16 июн. 2019 г.

@author: alex

Скрипт для создания отдельного исполняемого файла с использованием библиотеки cx_Freeze
'''


from cx_Freeze import setup, Executable
import sys
base = None
if sys.platform == "win32":
    base = "Win32GUI"
    
executables = [Executable('XL.py', base = base)]

setup(name='XL',
      version='0.0.1',
      description='Автоматизация',
      executables=executables)