import requests
import json
import pandas as pd
import xlwings as xw
import time

index_options = ["NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY"]

# Parameters
sym = "NIFTY"
exp_date = "28-Mar-2024"
interval = 180  # In seconds
next_row = 5  # PCR Column Values
call_oi_chg_column = "K2"
put_oi_chg_column = "K3"

file = xw.Book(
    r"C:\Users\bipin\OneDrive\Desktop\bipinmsit\tradetools\option_chain.xlsx"
)
sh1 = file.sheets("option_chain")

# Column range to clear
column_to_clear = "A"
column_range = sh1.range("A:H")


def oc_api_res(symbol):
    url = "https://www.nseindia.com/api/option-chain-indices?symbol=" + symbol
    headers = {
        "accept-encoding": "gzip, deflate, br, zstd",
        "accept-language": "en-US,en;q=0.9",
        "referer": "https://www.nseindia.com/option-chain",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    }
    res = requests.get(url, headers=headers).text

    return res


def oc_expiry_list(symbol):
    expiry_list = []

    try:
        res = oc_api_res(symbol)
        strike_data = json.loads(str(res))

        expiry_list = strike_data["records"]["expiryDates"]

        return expiry_list
    except Exception as e:
        print("Something went wrong while getting expiry list. Please run again!", e)

        return expiry_list


def oc_spot(symbol):
    spot = ""

    try:
        res = oc_api_res(symbol)
        strike_data = json.loads(str(res))

        spot = strike_data["records"]["underlyingValue"]

        return spot
    except Exception as e:
        print("Something went wrong while getting spot. Please run again!", e)

        return spot


def oc(symbol, expiry_date):
    res = oc_api_res(symbol)
    strike_data = json.loads(str(res))

    ce = {}
    pe = {}

    n = 0
    m = 0

    for i in strike_data["records"]["data"]:
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

    return final_df


def create_dropdown_in_excel(sheet, options, dropdown_range):
    # Write the options to the worksheet
    sheet.range(dropdown_range).options(transpose=True).value = [options]

    # Add a dropdown list to the range
    sheet.range(
        dropdown_range
    ).api.Validation.Delete()  # Remove any existing validation
    sheet.range(dropdown_range).api.Validation.Add(
        Type=3,  # Type 3 represents List validation
        AlertStyle=1,  # AlertStyle 1 represents Stop alert
        Formula1='"' + ",".join(options) + '"',  # Set the range of options
    )


# Create dropdown of index and expiry date
create_dropdown_in_excel(sh1, index_options, "L2:L10")
create_dropdown_in_excel(sh1, oc_expiry_list("NIFTY"), "M2:M10")

# Delete the records from the specified range
last_row_expiry = sh1.range("M" + str(sh1.cells.last_cell.row)).end("up").row
unwanted_row_expiry_delete = "M3:M" + str(last_row_expiry)
sh1.range(unwanted_row_expiry_delete).api.Delete()

# Delete the records from the specified range
last_row_expiry = sh1.range("L" + str(sh1.cells.last_cell.row)).end("up").row
unwanted_row_expiry_delete = "L3:L" + str(last_row_expiry)
sh1.range(unwanted_row_expiry_delete).api.Delete()

# Delete the records from the specified range
last_row_pcr_change = sh1.range("K" + str(sh1.cells.last_cell.row)).end("up").row
unwanted_row_pcr_change_delete = "K5:K" + str(last_row_pcr_change)
sh1.range(unwanted_row_pcr_change_delete).api.Delete()

# Delete the records from the specified range
last_row_time = sh1.range("J" + str(sh1.cells.last_cell.row)).end("up").row
unwanted_row_time_delete = "j5:J" + str(last_row_time)
sh1.range(unwanted_row_time_delete).api.Delete()


while True:
    try:
        data = oc(sym, exp_date)
        # Clear sheet with every interval
        # sh1.clear()
        column_range.clear_contents()
        sh1.range("A1").value = data

        total_call_change = sh1.range(call_oi_chg_column).value
        total_put_change = sh1.range(put_oi_chg_column).value

        pcr = round(total_put_change / total_call_change, 3)
        print("PCR: ", pcr)

        curr_time = time.strftime("%H:%M:%S", time.localtime())
        sh1.range("J{}".format(next_row)).value = curr_time
        sh1.range("K{}".format(next_row)).value = pcr

        sh1.range("N2").value = oc_spot("NIFTY")

        # Save the sheet
        file.save()
        # file.close()

        time.sleep(interval)

        next_row = next_row + 1
    except:
        print("Retrying")
        time.sleep(5)
