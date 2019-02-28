# LinkedinSpider

## Environment Setting

OS: Windows 10

Web Browser: Chrome

Python: python anaconda 3.7

packages: selenium(webdriver), xlsxwriter, xlrd, xlwt, xlutils, pandas

## Running Arguments
1. Switch to the directory e.g. D:\Documents\python\LinkedinSpider and you have all the files including LinJDSpider.py. 
2. argument one: username for Linkedin Account
3. argument two: password for Linkedin Account
4. argument three: input file. This file should be .xlsx and it must be existing since it contains the jobs you want to crawl.
5. argument four: output file. This file should also be .xlsx and it isn't necessarily an existing file. 
For example:
python LinJDSpider.py username@gmail.com 123456 input.xlsx output.xlsx     

## Others
1. Chrome Browser will be open while running program. Don't do any operations on the browser and the operations are done by the programs automatically.
2. To make a safe crawler, I slow down the program to avoid being detected as a crawler. If that happens, you may need to manually enter your username and password and run the program again.
