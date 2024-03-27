import requests
import json
import pandas as pd
import xlwings as xw
import time
from get_parameter import read_parameters_from_excel
import threading


index_options = ["NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY", "SENSEX"]
number_of_strike = ["3", "5", "10", "15", "20"]

# Parameters
sym = "NIFTY"
exp_date = "28-Mar-2024"
number_of_strike_buffer = 5
interval = 120  # In seconds

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

# test = read_parameters_from_excel(file, "option_chain", "Spot")
# print("dddddddddd", test)


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


def nearest_multiple_of_50(number):
    return round(number / 50) * 50


def nearest_multiple_of_100(number):
    return round(number / 100) * 100


def nearest_multiple_of_25(number):
    return round(number / 25) * 25


def oc(symbol, expiry_date, no_of_strike):
    nearest_strike = ""
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

    spot = strike_data["records"]["underlyingValue"]
    if sym == "NIFTY" or sym == "FINNIFTY":
        nearest_strike = nearest_multiple_of_50(spot)
    elif sym == "BANKNIFTY" or sym == "SENSEX":
        nearest_strike = nearest_multiple_of_100(spot)
    else:
        nearest_strike = nearest_multiple_of_25(spot)

    buffer_strike = 50 * no_of_strike
    number_of_strike_above = nearest_strike + buffer_strike
    number_of_strike_below = nearest_strike - buffer_strike

    ce_df = pd.DataFrame.from_dict(ce).transpose()
    ce_df.columns += "_CE"

    ce_df = ce_df[ce_df["strikePrice_CE"] >= number_of_strike_below]
    ce_df = ce_df.rename(
        columns={
            "strikePrice_CE": "strikePrice",
        }
    )

    pe_df = pd.DataFrame.from_dict(pe).transpose()
    pe_df.columns += "_PE"

    pe_df = pe_df[pe_df["strikePrice_PE"] <= number_of_strike_above]
    pe_df = pe_df.rename(
        columns={
            "strikePrice_PE": "strikePrice",
        }
    )

    # Merge based on common column 'ID'
    df = pd.merge(ce_df, pe_df, on="strikePrice", how="inner")

    # Selected DataFrame
    df_selected = df[
        [
            "lastPrice_CE",
            "openInterest_CE",
            "changeinOpenInterest_CE",
            "strikePrice",
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
            "strikePrice": "Strike",
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
create_dropdown_in_excel(sh1, oc_expiry_list(sym), "M2:M10")
create_dropdown_in_excel(sh1, number_of_strike, "N2:N10")

# Delete the records from the specified range
last_row_no_strike = sh1.range("N" + str(sh1.cells.last_cell.row)).end("up").row
unwanted_row_no_strike_delete = "N3:N" + str(last_row_no_strike)
sh1.range(unwanted_row_no_strike_delete).api.Delete()

# Delete the records from the specified range
last_row_expiry = sh1.range("M" + str(sh1.cells.last_cell.row)).end("up").row
unwanted_row_expiry_delete = "M3:M" + str(last_row_expiry)
sh1.range(unwanted_row_expiry_delete).api.Delete()

# Delete the records from the specified range
last_row_options = sh1.range("L" + str(sh1.cells.last_cell.row)).end("up").row
unwanted_row_options_delete = "L3:L" + str(last_row_options)
sh1.range(unwanted_row_options_delete).api.Delete()

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
        data = oc(sym, exp_date, number_of_strike_buffer)
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

        sh1.range("O2").value = oc_spot(sym)

        # Save the sheet
        file.save()
        # file.close()

        time.sleep(interval)

        next_row = next_row + 1
    except:
        print("Retrying")
        time.sleep(5)
