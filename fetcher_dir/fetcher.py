from requests_html import HTMLSession
import os
from os import path
import json
from bs4 import BeautifulSoup as bs
import xlrd
from datetime import datetime
# s = "15-Jul-2020"
# f = "%d-%b-%Y"
# out = datetime.strptime(date_string, "%d/%m/%Y").strftime("%d-%b-%Y")
# print(out)


def fetch(url, filename):
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36 Edg/83.0.478.45',
                   'accept-encoding': 'gzip, deflate, br', 'accept-language': 'en-US,en;q=0.9'}

        # initialize the session
        session = HTMLSession()

        try:
            # make the HTTP request and retrieve response
            response = session.get(url)
            if(response.status_code == 200):
                # execute Javascript
                response.html.render(timeout=200000)

                # construct the soup parser
                soup = bs(response.html.html, "html.parser")
                print('Data Fetched from this ' + url)

                try:
                    scrap(soup, filename)
                except Exception as e:
                    print('Error - Data Scraping Failed ' +
                          url + ' {}'.format(e))
            else:
                print(url)
                print('Site is not responding or internet is not availabe {}'.format(
                    response.status_code))

        except Exception as e:
            print('Error - fetch Failed for ' + url + ' {}'.format(e))
    except Exception as e:
        print('Error - fetcher Failed for ' + url + ' {}'.format(e))


def file_fetch(url, filename):
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36 Edg/83.0.478.45',
                   'accept-encoding': 'gzip, deflate, br', 'accept-language': 'en-US,en;q=0.9'}

        # initialize the session
        session = HTMLSession()

        try:
            # make the HTTP request and retrieve response
            response = session.get(url)
            if(response.status_code == 200):
                print('Data Fetched from this ' + url)
                try:
                    scrap(response.content, filename)
                except Exception as e:
                    print('Error - Data Scraping Failed ' +
                          url + ' {}'.format(e))
            else:
                print(url)
                print(url + ' is not responding or internet is not availabe {}'.format(
                    response.status_code))

        except Exception as e:
            print('Error - fetch Failed for ' + url + ' {}'.format(e))
    except Exception as e:
        print('Error - fetcher Failed for ' + url + ' {}'.format(e))


def scrap(soup, fileName):

    directory = "./temp_data"
    if(os.path.exists(directory) != True):
        os.mkdir(directory)

    if(fileName == 'cdsl-FII-DD'):
        rows = soup.findAll('tr')  # Extract and return first occurrence of tr

        td0 = rows[0].find_all('td')
        td1 = rows[1].find_all('td')
        td14 = rows[14].find_all('td')

        # for td in td0:
        #     print(td.get_text())
        # for td in td1:
        #     print(td.get_text())
        note = td14[0].get_text()
        note = note.replace("\n", "").replace("       ", "").replace(
            "\r", "").replace("               ", "").strip()

        data1 = {"Date": datetime.strptime(td1[0].get_text(), "%d-%b-%Y").strftime("%d-%b-%Y"),
                 "Equity": {"Stock Exchange":
                            {"Gross Purchases (Rs Crore)": td1[3].get_text(),
                             "Gross Sales (Rs Crore)": td1[4].get_text(),
                             "Net Investment (Rs Crore)": td1[5].get_text()}},
                 "Note": note}

        # print(data1)

        with open(directory+'/'+fileName+'-table1.json', 'w') as f:
            json.dump(data1, f)

        ##
        # Second table
        ##

        td0 = rows[15].find_all('td')
        td1 = rows[16].find_all('td')
        td2 = rows[17].find_all('td')
        td3 = rows[18].find_all('td')
        td4 = rows[19].find_all('td')
        td5 = rows[20].find_all('td')
        td6 = rows[21].find_all('td')
        td7 = rows[22].find_all('td')

        # print('------------------------------')
        # for td in td0:
        #     print(td.get_text())
        # print('------------------------------')
        # for td in td1:
        #     print(td.get_text())
        # print('------------------------------')
        # for td in td2:
        #     print(td.get_text())
        # print('------------------------------')
        # for td in td3:
        #     print(td.get_text())
        # print('------------------------------')
        # for td in td4:
        #     print(td.get_text())
        # print('------------------------------')
        # for td in td5:
        #     print(td.get_text())
        # print('------------------------------')
        # for td in td6:
        #     print(td.get_text())
        # print('------------------------------')
        # for td in td7:
        #     print(td.get_text())
        # print('------------------------------')

        note = td7[0].get_text()
        note = note.replace("\n", "").replace(
            "\r", "").replace("               ", "").strip()

        data2 = {"Date": datetime.strptime(td2[0].get_text(), "%d-%b-%Y").strftime("%d-%b-%Y"),
                 "Buy": {"No of contract":
                         {"Index Futures": td2[2].get_text(),
                          "Index Options": td3[2].get_text(),
                          "Stock Futures": td4[2].get_text(),
                          "Stock Options": td5[2].get_text(),
                          "Interest Rate Futures": td6[2].get_text()
                          },
                         "Amount in Crore":
                         {"Index Futures": td2[3].get_text(),
                          "Index Options": td3[3].get_text(),
                          "Stock Futures": td4[3].get_text(),
                          "Stock Options": td5[3].get_text(),
                          "Interest Rate Futures": td6[3].get_text()
                          }},
                 "Sell": {"No of contract":
                          {"Index Futures": td2[4].get_text(),
                           "Index Options": td3[4].get_text(),
                           "Stock Futures": td4[4].get_text(),
                           "Stock Options": td5[4].get_text(),
                           "Interest Rate Futures": td6[4].get_text()
                           },
                          "Amount in Crore":
                          {"Index Futures": td2[5].get_text(),
                           "Index Options": td3[5].get_text(),
                           "Stock Futures": td4[5].get_text(),
                           "Stock Options": td5[5].get_text(),
                           "Interest Rate Futures": td6[5].get_text()
                           }},
                 "Open Interest at the end of the date": {"No of contract":
                                                          {"Index Futures": td2[6].get_text(),
                                                           "Index Options": td3[6].get_text(),
                                                           "Stock Futures": td4[6].get_text(),
                                                           "Stock Options": td5[6].get_text(),
                                                           "Interest Rate Futures": td6[6].get_text()
                                                           },
                                                          "Amount in Crore":
                                                          {"Index Futures": td2[7].get_text(),
                                                           "Index Options": td3[7].get_text(),
                                                           "Stock Futures": td4[7].get_text(),
                                                           "Stock Options": td5[7].get_text(),
                                                           "Interest Rate Futures": td6[7].get_text()
                                                           }},
                 "Note": note}
        # prin(data2)
        with open(directory+'/'+fileName+'-table2.json', 'w') as f:
            json.dump(data2, f)

    if(fileName == 'bse-equity-CW-TO'):
        # Extract and return first occurrence of tr
        trs = soup.find(id='ContentPlaceHolder1_offTblBdyDII').find_all('tr')
        tds1 = trs[1]
        tdArray1 = []
        for td in tds1:
            try:
                tdArray1.append(td.get_text())
            except:
                pass
        trs = soup.find(id='offTblBdy').find_all('tr')
        tds2 = trs[0]
        tdArray2 = []
        for td in tds2:
            try:
                tdArray2.append(td.get_text())
            except:
                pass
        data1 = {
            "Category": tdArray1[0],
            "Date": datetime.strptime(tdArray1[1], "%d/%m/%Y").strftime("%d-%b-%Y"),
            "Buy Value": tdArray1[2],
            "Sale Value": tdArray1[3],
            "Net Value": tdArray1[4],
        }
        with open(directory+'/'+fileName+'-table1.json', 'w') as f:
            json.dump(data1, f)

        data2 = {
            "Date": datetime.strptime(tdArray2[0], "%d/%m/%Y").strftime("%d-%b-%Y"),
            "Clients": {
                "Buy": tdArray2[1],
                "Sales": tdArray2[2],
                "Net": tdArray2[3]
            },
            "NRI": {
                "Buy": tdArray2[4],
                "Sales": tdArray2[5],
                "Net": tdArray2[6]
            },
            "Proprietary": {
                "Buy": tdArray2[7],
                "Sales": tdArray2[8],
                "Net": tdArray2[9]
            }
        }
        # print(data2)
        with open(directory+'/'+fileName+'-table2.json', 'w') as f:
            json.dump(data2, f)

    if(fileName == 'equity-CM-CW-turnover'):
        with open(directory+'/'+fileName+'.xls', 'wb') as f:
            f.write(soup)

        # To open Workbook
        wb = xlrd.open_workbook(directory+'/'+fileName+'.xls')
        sheet = wb.sheet_by_index(0)

        # For row 0 and colom 0
        # print(sheet.cell_value(0, 0))
        if(sheet.cell_value(3, 1) == "BNK"):
            row_bnk = 3
        if(sheet.cell_value(4, 1) == "DFI"):
            row_dfi = 4
        else:
            row_dfi = None
        if(sheet.cell_value(4, 1) == "PRO-TRADES"):
            row_pro = 4
        if(sheet.cell_value(5, 1) == "PRO-TRADES"):
            row_pro = 5
        if(sheet.cell_value(5, 1) == "OTHERS"):
            row_other = 5
        if(sheet.cell_value(6, 1) == "OTHERS"):
            row_other = 6

        data = {
            "Date": datetime.strptime(sheet.cell_value(3, 0), "%d-%b-%y").strftime("%d-%b-%Y"),
            "Pro Trades": {
                "Buy Value (Rs. in Crores)": sheet.cell_value(row_pro, 2),
                "Sell Value (Rs. in Crores)": sheet.cell_value(row_pro, 3),
            },
            "Others Trades": {
                "Buy Value (Rs. in Crores)": sheet.cell_value(row_other, 2),
                "Sell Value (Rs. in Crores)": sheet.cell_value(row_other, 3),
            },
            "BNK&DFI": {
                "Buy Value (Rs. in Crores)": sheet.cell_value(row_bnk, 2) + (sheet.cell_value(row_dfi, 2) if(row_dfi != None) else 0),
                "Sell Value (Rs. in Crores)": sheet.cell_value(row_bnk, 3) + (sheet.cell_value(row_dfi, 3) if(row_dfi != None) else 0),
            },
        }
        # print(data)

        with open(directory+'/'+fileName+'.json', 'w') as f:
            json.dump(data, f)
