from fetcher_dir.fetcher import fetch, file_fetch
from utils.link_assambler import assamble
from utils.excel_writer import writer
import sys
import pyfiglet

ascii_banner = pyfiglet.figlet_format("Stock  Updater")
print(ascii_banner)
print(" ")
print("****************************************************************")
print("Provide date for NSE In Format d m YYYY. example => 15-7-2020")
print("****************************************************************")
argument_data = input("Input date and press Enter...\n")
arr = argument_data.split("-")
date = int(arr[0])
month = int(arr[1])
year = int(arr[2])
print(" ")
print("Date is {}".format(date))
print("Month is {}".format(month))
print("Year is {}".format(year))
print(" ")


# Download data
file_fetch(assamble('equity-CM-CW-turnover', date,
                    month, year), 'equity-CM-CW-turnover')
fetch(assamble('cdsl-FII-DD'), 'cdsl-FII-DD')
fetch(assamble('bse-equity-CW-TO'), 'bse-equity-CW-TO')


# write to excel
writer()
