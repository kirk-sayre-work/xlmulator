# xlmulator
This is a Python emulator for Excel XLM macros. It runs under Linux
(developed in Ubuntu). It works in Python 2, Python 3, and PyPy.

## Running

The simplest way to run XLMulator is to run it out of a prebuilt
Docker container contained in the global Docker container repository. 

## Installation

XLMulator uses LibreOffice to dump the data cells of the Excel
workbook being analyzed. Install LibreOffice and the LibreOffice
Python UNO programmatic bridge.

`
apt-get update && apt-get install -y \
	build-essential \
	file\
      	libimage-exiftool-perl\
	libpython-dev\
        libreoffice\
        libreoffice-script-provider-python\
        uno-libs3\
        python3-uno\
        python3\
        unzip\
`

Install the Python packages used by XLMulator (change the pip commands
based on whether you are using Python 2, Python 3, or PyPy).

`pip install lark-parser prettytable`
