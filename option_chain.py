import requests
import json
import pandas as pd
import xlwings as xw
import time

index_options = ["NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY"]

# Parameters
sym = "NIFTY"
exp_date = "28-Mar-2024"
interval = 3  # In seconds
next_row = 5  # PCR Column Values
call_oi_chg_column = "K2"
put_oi_chg_column = "K3"

file = xw.Book(
    r"C:\Users\bipin\OneDrive\Desktop\bipinmsit\tradetools\option_chain.xlsx"
)
sh1 = file.sheets("option_chain")
sh2 = file.sheets("expiry_date")

# Column range to clear
column_to_clear = "A"
column_range = sh1.range("A:H")


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

    # Selected DataFrame
    df_selected = df[
        [
            "lastPrice_CE",
            "openInterest_CE",
            "changeinOpenInterest_CE",
            "strikePrice_CE",
            "changeinOpenInterest_PE",
            "openInterest_PE",
            "lastPrice_PE",
        ]
    ]

    final_df = df_selected.rename(
        columns={
            "lastPrice_CE": "Call_LTP",
            "openInterest_CE": "Call_OI",
            "changeinOpenInterest_CE": "OI_Chg",
            "strikePrice_CE": "Strike",
            "changeinOpenInterest_PE": "OI_Chg",
            "openInterest_PE": "Put_OI",
            "lastPrice_PE": "Put_LTP",
        }
    )

    return expiry_list, final_df


while True:
    try:
        data = oc(sym, exp_date)
        # Clear sheet with every interval
        # sh1.clear()
        column_range.clear_contents()
        sh1.range("A1").value = data[1]

        # Clear sheet with every interval
        sh2.clear()
        sh2.range("A1").options(transpose=True).value = data[0]

        total_call_change = sh1.range(call_oi_chg_column).value
        total_put_change = sh1.range(put_oi_chg_column).value

        pcr = round(total_put_change / total_call_change, 3)
        print("PCR: ", pcr)

        curr_time = time.strftime("%H:%M:%S", time.localtime())
        sh1.range("J{}".format(next_row)).value = curr_time
        sh1.range("K{}".format(next_row)).value = pcr

        # Save the sheet
        file.save()
        # file.close()

        time.sleep(interval)

        next_row = next_row + 1
    except:
        print("Retrying")
        time.sleep(5)
