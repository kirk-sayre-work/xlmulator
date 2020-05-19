# XLMulator

This is a Python emulator for Excel XLM macros. It runs under Linux
(developed in Ubuntu). It works in Python 2, Python 3, and PyPy.

Excel 97 and 2007+ file with XLM macros are supported.

## Running

The simplest way to run XLMulator is to run it out of a prebuilt
Docker container contained in the global Docker container
repository. If you already have Docker installed you can just use the
docker-xlmulator.sh script to start analyzing Excel workbooks with XLM
macros. Note that depending on how Docker is set up on your machine you
may need to run this script with sudo.

`docker-xlmulator.sh FILE.xls`

The Docker script does the following:

1. Pulls down the most recent Docker container.
2. Turns off networking in the Docker container.
3. Does a git pull on the XLMulator GitHub repository to ensure that the
   latest version is used.
4. Analyzes FILE.xls with XLMulator.
5. Analysis results are printed to stdout and (if desired) saved to a
   JSON file.

To analyze a file and save the analysis results to a JSON file run:

`docker-xlmulator.sh FILE.xls MY_OUT_FILE.json`

Analysis results will be saved in MY_OUT_FILE.json.

If you don't want to use the Docker script you will need to do a local
install of XLMulator. This is currently a manual process.

## Installation

Make local clone of the XLMulator GitHub repository.

`git clone https://github.com/kirk-sayre-work/xlmulator.git`

XLMulator uses LibreOffice to dump the data cells of the Excel
workbook being analyzed. Install LibreOffice and the LibreOffice
Python UNO programmatic bridge.

`
apt-get update && apt-get install -y 
	build-essential 
	file
      	libimage-exiftool-perl
	libpython-dev
        libreoffice
        libreoffice-script-provider-python
        uno-libs3
        python3-uno
        python3
        unzip
`

Install the Python packages used by XLMulator (change the pip commands
based on whether you are using Python 2, Python 3, or PyPy).

`pip install lark-parser prettytable`

XLMulator needs the most recent version of oletools, not what is
installed via pip. Do the following to install the current version of
oletools:

```
git clone https://github.com/decalage2/oletools.git
cd oletools
sudo python3 ./setup.py
```

### Install office_dumper

XLMulator uses another package to dump all sheets of an Excel workbook
to CSV files. As with XLMulator this is currently also a manual
install.

Clone the office_dumper GitHub repository.

`git clone https://github.com/kirk-sayre-work/office_dumper`

Install the psutil package.

`pip3 install psutil`

Add the office_dumper install directory to your PYTHONPATH environment
variable so that XLMulator can import it.

## Source Code

Doxygen generated documentation for the source code of XLMulator is
available in the code_docs/ directory. Load code_docs/html/index.html in a
web browser to view the code documentation.

## Acknowledgements

Thanks to @DissectMalware for allowing their MS XLM grammar to be used in
XLMulator. Check out their XLM emulator at
[XLMMacroDeobfuscator](https://github.com/DissectMalware/XLMMacroDeobfuscator/).
