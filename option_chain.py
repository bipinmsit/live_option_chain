import requests
import json
import pandas as pd
import xlwings as xw
import time
from get_parameter import read_parameters_from_excel

index_options = ["NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY"]
number_of_strike = ["3", "5", "10", "15", "20"]
excel_file_path = (
    r"C:\Users\bipin\OneDrive\Desktop\bipinmsit\tradetools\option_chain.xlsx"
)
parameters_output = (
    r"C:\Users\bipin\OneDrive\Desktop\bipinmsit\tradetools\parameters.txt"
)

# # Parameters
sym = read_parameters_from_excel(excel_file_path, "option_chain", "L2")
if sym == "NIFTY":
    exp_date = read_parameters_from_excel(excel_file_path, "option_chain", "M2")
elif sym == "BANKNIFTY":
    exp_date = read_parameters_from_excel(excel_file_path, "option_chain", "N2")
elif sym == "FINNIFTY":
    exp_date = read_parameters_from_excel(excel_file_path, "option_chain", "O2")
else:
    exp_date = read_parameters_from_excel(excel_file_path, "option_chain", "P2")
number_of_strike_buffer = int(
    read_parameters_from_excel(excel_file_path, "option_chain", "Q2")
)
interval = 1  # In seconds
print(
    "index: {}, expiry_date: {}, no_of_strike: {}, time_stamp: {}".format(
        sym, exp_date, number_of_strike_buffer, interval
    )
)
with open(parameters_output, "w") as f:
    print(
        "index: {}, expiry_date: {}, no_of_strike: {}, time_stamp: {}".format(
            sym, exp_date, number_of_strike_buffer, interval
        ),
        file=f,
    )


next_row = 5  # PCR Column Values
call_oi_chg_column = "K2"
put_oi_chg_column = "K3"

file = xw.Book(excel_file_path)
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


def nearest_multiple_of_50(number):
    return round(number / 50) * 50


def nearest_multiple_of_100(number):
    return round(number / 100) * 100


def nearest_multiple_of_25(number):
    return round(number / 25) * 25


def download_data(symbol, no_of_strike):
    nearest_strike = ""
    buffer_strike = ""
    res = oc_api_res(symbol)
    strike_data = json.loads(str(res))

    ce = {}
    pe = {}

    n = 0
    m = 0

    for i in strike_data["records"]["data"]:
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
        buffer_strike = 50 * no_of_strike
    elif sym == "BANKNIFTY" or sym == "SENSEX":
        nearest_strike = nearest_multiple_of_100(spot)
        buffer_strike = 100 * no_of_strike
    else:
        nearest_strike = nearest_multiple_of_25(spot)
        buffer_strike = 25 * no_of_strike

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


def oc(symbol, expiry_date, no_of_strike):
    nearest_strike = ""
    buffer_strike = ""
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
        buffer_strike = 50 * no_of_strike
    elif sym == "BANKNIFTY" or sym == "SENSEX":
        nearest_strike = nearest_multiple_of_100(spot)
        buffer_strike = 100 * no_of_strike
    else:
        nearest_strike = nearest_multiple_of_25(spot)
        buffer_strike = 25 * no_of_strike

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


nifty_expiry_list = oc_expiry_list("NIFTY")
banknifty_expiry_list = oc_expiry_list("BANKNIFTY")
finnifty_expiry_list = oc_expiry_list("FINNIFTY")
midcpnifty_expiry_list = oc_expiry_list("MIDCPNIFTY")

# Create dropdown of index and expiry date
create_dropdown_in_excel(sh1, index_options, "L2:L2")
create_dropdown_in_excel(sh1, nifty_expiry_list, "M2:M2")
create_dropdown_in_excel(sh1, banknifty_expiry_list, "N2:N2")
create_dropdown_in_excel(sh1, finnifty_expiry_list, "O2:O2")
create_dropdown_in_excel(sh1, midcpnifty_expiry_list, "P2:P2")
create_dropdown_in_excel(sh1, number_of_strike, "Q2:Q2")


sh1.range("Q3:Q20").clear()
sh1.range("P3:P20").clear()
sh1.range("O3:O20").clear()
sh1.range("N3:N20").clear()
sh1.range("M3:M20").clear()
sh1.range("L3:L20").clear()
sh1.range("K5:K1000").clear()
sh1.range("J5:J1000").clear()


old_pcr = None
old_spot = None
while True:
    try:
        data = oc(sym, exp_date, number_of_strike_buffer)
        # Clear sheet with every interval
        # sh1.clear()
        column_range.clear_contents()
        sh1.range("A1").value = data

        total_call_change = sh1.range(call_oi_chg_column).value
        total_put_change = sh1.range(put_oi_chg_column).value

        new_pcr = round(total_put_change / total_call_change, 3)
        print("PCR_CHANGE: ", new_pcr)

        # Check if new value is different from old value
        if new_pcr != old_pcr:
            curr_time = time.strftime("%H:%M:%S", time.localtime())
            sh1.range("J{}".format(next_row)).value = curr_time
            sh1.range("K{}".format(next_row)).value = new_pcr

            next_row = next_row + 1

        # Update old value
        old_pcr = new_pcr

        new_spot = oc_spot(sym)
        if new_spot != old_spot:
            sh1.range("R2").value = oc_spot(sym)

        old_spot = new_spot

        # Save the sheet
        file.save()
        # file.close()

        time.sleep(interval)

    except:
        print("Retrying")
        time.sleep(5)
