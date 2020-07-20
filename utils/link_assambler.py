def assamble(fileName='undefined.json', date=2, month=7, year=2020):
    MONTHS = {
        '01': 'JAN',
        '02': 'FEB',
        '03': 'MAR',
        '04': 'APR',
        '05': 'MAY',
        '06': 'JUN',
        '07': 'JUL',
        '08': 'AUG',
        '09': 'SEP',
        '10': 'OCT',
        '11': 'NOV',
        '12': 'DEC'
    }
    LINKS = {
        'equity-bhavcopy': 'https://archives.nseindia.com/content/historical/EQUITIES/',
        'equity-CM-CW-turnover': 'https://archives.nseindia.com/archives/equities/cat/cat_turnover_',
        'derivative-bhavcopy': 'https://archives.nseindia.com/content/historical/DERIVATIVES/',
        'derivative-PW-TV': 'https://archives.nseindia.com/content/nsccl/fao_participant_vol_',
        'derivative-PW-OI': 'https://archives.nseindia.com/content/nsccl/fao_participant_oi_',
        'derivative-FII-DS': 'https://archives.nseindia.com/content/fo/fii_stats_',
        'derivative-COI-AE': 'https://archives.nseindia.com/archives/nsccl/mwpl/combineoi_',
        'cdsl-FII-DD': 'https://www.cdslindia.com/Publications/FIIDailyData.aspx',
        'bse-equity-CW-TO': 'https://www.bseindia.com/markets/equity/EQReports/categorywise_turnover.aspx'
    }

    try:
        monthString = MONTHS.get(
            str(f"{month:02d}"), 'ERROR: Entered month is wrong ')
        if(fileName == 'equity-bhavcopy'):
            link = LINKS.get(fileName) + str(year) + '/' + monthString + \
                '/cm' + f"{date:02d}" + monthString + \
                str(year) + 'bhav.csv.zip'
        if(fileName == 'equity-CM-CW-turnover'):
            link = LINKS.get(fileName)+f"{date:02d}" + \
                f"{month:02d}"+str(year)[:-2]+".xls"
        if(fileName == 'derivative-bhavcopy'):
            link = LINKS.get(fileName) + str(year) + '/' + monthString + \
                '/fo' + f"{date:02d}" + monthString + \
                str(year) + 'bhav.csv.zip'
        if(fileName == 'derivative-PW-TV'):
            link = LINKS.get(fileName) + \
                f"{date:02d}" + f"{month:02d}" + str(year) + '.csv'
        if(fileName == 'derivative-PW-OI'):
            link = LINKS.get(fileName) + \
                f"{date:02d}" + f"{month:02d}" + str(year) + '.csv'
        if(fileName == 'derivative-FII-DS'):
            link = LINKS.get(fileName) + f"{date:02d}" + "-" + \
                monthString.capitalize() + "-" + str(year) + '.xls'
        if(fileName == 'derivative-COI-AE'):
            link = LINKS.get(fileName) + \
                f"{date:02d}" + f"{month:02d}" + str(year) + '.zip'
        if(fileName == 'cdsl-FII-DD'):
            link = LINKS.get(fileName)
        if(fileName == 'bse-equity-CW-TO'):
            link = LINKS.get(fileName)
        return link
    except Exception as e:
        print(' ERROR: link generation error ' + str(e))
        return ' ERROR: link generation error ' + str(e)
