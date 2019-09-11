import mechanize
import xlsxwriter
from bs4 import BeautifulSoup
from http.cookiejar import CookieJar
import pandas as pd
# These imports are specific to the google sheets function
import pygsheets


def littlefield_script():

    # authorization for google sheets account
    gc = pygsheets.authorize(service_file='creds.json')

 
    # assign the necessary variables
    cj = CookieJar()
    br = mechanize.Browser()
    br.set_cookiejar(cj)
    br.open("http://op.responsive.net/lt/pmba/entry.html")
    br.select_form(nr=0)
    br.form['id'] = "fake_user_name"
    br.form['password'] = "fake_password"
    br.submit()
    url_list = ["CASH", "JOBIN", "JOBQ", "S1Q", "S2Q", "S3Q", "S1UTIL", "S2UTIL", "S3UTIL"]
    url_list_3col = ["JOBT", "JOBREV", "JOBOUT"]
    LF_DATA = {}

    # get INVENTORY first
    inv_url = "http://op.responsive.net/Littlefield/Plot?data=INV&x=all"
    soup = BeautifulSoup(br.open(inv_url), "lxml")
    data = soup.find_all("script")[5].string
    data = data.split("\n")[4].split("'")[3].split()
    counter = 1
    for i in data:
        if counter % 2 == 1:
            counter += 1
            day = float(i)
            LF_DATA[day] = []
        elif counter % 2 == 0:
            row_data = [float(i)]
            LF_DATA[day].extend(row_data)
            counter += 1
            
    # list comprehension to delete values from inventory dictionary that are not "integers". This is essentially getting rid of the data where the system records inventory receipts.
    delete = [i for i in LF_DATA if i % int(i) != 0]
    for i in delete:
        del LF_DATA[i]

    # iterate through and scrape all two-column tables
    for url in url_list:
        lf_url = "http://op.responsive.net/Littlefield/Plot?data=%s&x=all" % url
        soup = BeautifulSoup(br.open(lf_url), "lxml")
        data = soup.find_all("script")[5].string
        data = data.split("\n")[4].split("'")[3].split()
        counter = 1
        for i in data:
            if counter % 2 == 0:
                day = counter / 2
                LF_DATA[day].append(float(i))
                counter += 1
            else:
                counter += 1

    # iterate through and scrape all three-column tables
    for url in url_list_3col:
        lf_url = "http://op.responsive.net/Littlefield/Plot?data=%s&x=all" % url
        soup = BeautifulSoup(br.open(lf_url), "lxml")
        data = soup.find_all("script")[5].string
        data0 = data.split("\n")[4].split("'")[5].split()
        data1 = data.split("\n")[5].split("'")[5].split()
        data2 = data.split("\n")[6].split("'")[5].split()

        counter = 1
        for i in data0:
            if counter % 2 == 0:
                day = counter / 2
                LF_DATA[day].append(float(i))
                counter += 1
            else:
                counter += 1
        counter = 1
        for i in data1:
            if counter % 2 == 0:
                day = counter / 2
                LF_DATA[day].append(float(i))
                counter += 1
            else:
                counter += 1
        counter = 1
        for i in data2:
            if counter % 2 == 0:
                day = counter / 2
                LF_DATA[day].append(float(i))
                counter += 1
            else:
                counter += 1


    # Add dummy data to fill out fractional day rows
    dummy_data = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for key, value in LF_DATA.items():
        if len(value) < 19:
            value.extend(dummy_data)


    # Prepare the dataframe to be written to the google sheet
    writer = pd.ExcelWriter('data.xlsx', engine = 'xlsxwriter')
    headers = ["inventory", "cash", "orders", "order \nqueue", "s1\nqueue", "s2\nqueue", "s3\nqueue", "s1\nutilization", "s2\nutilization", "s3\nutilization", \
               "c1\naverage\nleadtime", "c2\naverage\nleadtime", "c3\naverage\nleadtime", "c1\naverage\nrevenues", "c2\naverage\nrevenues", "c3\naverage\nrevenues", \
               "c1\njobs\ncompleted", "c2\njobs\ncompleted", "c3\njobs\ncompleted"]
    df = pd.DataFrame.from_dict(LF_DATA, orient="index")
    df.index = df.index.map(str)
    df.columns = headers


    # Fix issue with cash in $1,000's
    df.loc[:, 'cash'] *= 1000


    # Open the google sheets doc where we want to write the data, select the first sheet, and update the sheet
    sh = gc.open_by_key('10xRAaT9-MMddEBcvmiSMclbwUeh2VS5Z_0wduWh4Xyo')
    wks = sh[1]
    wks.set_dataframe(df,'B1',copy_index=True)


    # Code to format the sheet appropriately
    currency = wks.cell('D2').set_number_format(pygsheets.FormatType.CURRENCY, '$#,##0.00')
    percent = wks.cell('J2').set_number_format(pygsheets.FormatType.PERCENT, '#0%')
    number = wks.cell('G2').set_number_format(pygsheets.FormatType.NUMBER, '#,##0.000')
    digits = wks.cell('G2').set_number_format(pygsheets.FormatType.NUMBER, '#,##0')

    j_thru_l = wks.get_values('J2', 'L280', returnas='range')
    g_thru_i = wks.get_values('G2', 'I280', returnas='range')
    e_thru_f = wks.get_values('E2', 'F280', returnas='range')
    d_thru_d = wks.get_values('D2', 'D280', returnas='range')
    p_thru_r = wks.get_values('P2', 'R280', returnas='range')
    m_thru_o = wks.get_values('M2', 'O280', returnas='range')
    s_thru_u = wks.get_values('S2', 'U280', returnas='range')

    j_thru_l.apply_format(percent)
    g_thru_i.apply_format(digits)
    e_thru_f.apply_format(digits)
    d_thru_d.apply_format(currency)
    p_thru_r.apply_format(currency)
    m_thru_o.apply_format(number)
    s_thru_u.apply_format(digits)


littlefield_script()
print("The script has been successfully executed.")

