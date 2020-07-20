from fetcher_dir.fetcher import fetch, file_fetch
from utils.link_assambler import assamble

print(assamble('cdsl-FII-DD'))
fetch(assamble('cdsl-FII-DD'), 'cdsl-FII-DD')
print(assamble('bse-equity-CW-TO'))
fetch(assamble('bse-equity-CW-TO'), 'bse-equity-CW-TO')
print(assamble('equity-CM-CW-turnover', 9, 7, 2020))
file_fetch(assamble('equity-CM-CW-turnover', 9, 7, 2020),
           'equity-CM-CW-turnover')
