import requests
import json
import pandas as pd
import xlwings as xw
import time

sym = "NIFTY"
exp_date = "28-Mar-2024"

file = xw.Book(
    r"C:\Users\bipin\OneDrive\Desktop\bipinmsit\tradetools\option_chain.xlsx"
)
sh1 = file.sheets("Sheet1")
sh2 = file.sheets("Sheet2")


def oc(symbol, expiry_date):
    url = "https://www.nseindia.com/api/option-chain-indices?symbol=" + symbol
    headers = {
        "accept-encoding": "gzip, deflate, br, zstd",
        "accept-language": "en-US,en;q=0.9",
        "referer": "https://www.nseindia.com/option-chain",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    }
    res = requests.get(url, headers=headers).text

    data = json.loads(str(res))
    expiry_list = data["records"]["expiryDates"]

    ce = {}
    pe = {}

    n = 0
    m = 0

    for i in data["records"]["data"]:
        if i["expiryDate"] == expiry_date:
            try:
                ce[n] = i["CE"]
                n = n + 1
            except:
                pass
            try:
                pe[m] = i["PE"]
                m = m + 1
            except:
                pass

    ce_df = pd.DataFrame.from_dict(ce).transpose()
    ce_df.columns += "_CE"
    pe_df = pd.DataFrame.from_dict(pe).transpose()
    pe_df.columns += "_PE"

    df = pd.concat([ce_df, pe_df], axis=1)

    return expiry_list, df


while True:
    try:
        data = oc(sym, exp_date)
        sh1.clear()
        sh1.range("A1").value = data[1]
        sh2.clear()
        sh2.range("A1").options(transpose=True).value = data[0]

        time.sleep(1)
    except:
        print("Retrying")
        time.sleep(3)


# print(df.columns)
