import io
import os
import sys
import pandas as pd

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from era_data_merger import parse_general_report, FINAL_COLUMNS

OLD_CSV = """Order Owner,Order Status,Local Marketing Order Ref,Stock Order Ref,Order Placed Date,Order Line Reference,Local Marketing Asset,Local Marketing Order Line Ref,Stock Item,Stock Order Line Ref,Part,Quantity,Date Approved,Location,Workflow Reference Number,Sell Price (DT),Sell Price (FT),Sell Price (STK),Sell Price (STA),Sell Price (RFP),Buy Price (DT),Buy Price (FT)
5467 Blackburn,In Progress,BZL13333,,2025-07-08 17:28:28,24543/2,AS89292,BZL13333/43034,,,40x30 Poster OT,2,2025-07-11 16:21:38,Blackburn,BUZ17550/125557,0,9.82,0,0,0,0,13.1
5611 Barnsley Pontefract Road,In Progress,BZL13313,,2025-06-30 20:06:11,24610/6,AS87957,BZL13313/42982,,,A7 2pp Voucher (L) OT,100,2025-07-07 13:12:49,Barnsley Pontefract Road,BUZ17503/125411,0,17.18,0,0,0,0,2
5611 Barnsley Pontefract Road,In Progress,BZL13313,,2025-06-30 20:06:11,24610/7,AS86292,BZL13313/42981,,,A5 2PP Leaflet OT,200,2025-07-07 13:12:49,Barnsley Pontefract Road,BUZ17503/125410,0,4.89,0,0,0,0,40
5878 Sheffield Wadsley,In Progress,BZL13314,,2025-07-01 13:39:53,24614/7,AS85197,BZL13314/42987,,,40x30 Poster OT,4,2025-07-07 13:20:31,Sheffield Wadsley,BUZ17504/125416,19.64,0,0,0,0,26.19,0
5823 Washington,In Progress,BZL13315,,2025-07-01 14:41:11,24615/2,AS85987,BZL13315/42997,,,60x40 Poster OT,2,2025-07-07 13:23:02,Washington,BUZ17505/125426,17.14,0,0,0,0,21.38,0
"""

NEW_CSV = """Order Owner,Order Status,Local Marketing Order Ref,Stock Order Ref,Order Placed Date,Order Line Reference,Local Marketing Asset,Local Marketing Order Line Ref,Stock Item,Stock Order Line Ref,Part,Quantity,Date Approved,Location,Workflow Reference Number,If tender pre 5.25%,Print,Collate & pack,Despatch,Total
5861 Poole,Part Delivered,BZL13289,B3260,2025-06-27 14:44:50,24594/3,,,BUZ15866/118699,B3260/3,A5 2pp Perforated leaflet,2,2025-06-30 14:11:57,Poole,,,0,4.7,9.3,14
5861 Poole,Part Delivered,BZL13289,B3260,2025-06-27 14:44:50,24594/5,AS87957,BZL13289/42898,,,A7 2pp Voucher (L) OT,350,2025-06-30 14:11:57,Poole,BUZ17445/125121,,23.01,2.35,5.33,30.690000000000005
5472 Clacton on Sea,In Progress,BZL13293,,2025-06-27 16:59:48,24596/2,AS89981,BZL13293/42915,,,40x30 Poster OT,1,2025-06-30 14:21:08,Clacton on Sea,BUZ17449/125142,4.66,4.91,2.35,5.33,12.59
5472 Clacton on Sea,In Progress,BZL13293,,2025-06-27 16:59:48,24596/4,AS85986,BZL13293/42909,,,40x30 Poster OT,1,2025-06-30 14:21:08,Clacton on Sea,BUZ17449/125136,4.66,4.91,2.35,5.33,12.59
5543 Hanley,Completed,,B3261,2025-06-28 14:52:48,24604/1,,,BUZ12427/114467,B3261/1,Prize Draw Envelopes,4,2025-06-30 14:13:56,Hanley,,,0,9.4,18.6,28
"""

def _read_raw(csv_text: str) -> pd.DataFrame:
    return pd.read_csv(io.StringIO("\n" + csv_text), header=None)


def test_old_format_parsing():
    raw = _read_raw(OLD_CSV)
    df = parse_general_report(raw)
    assert list(df.columns) == FINAL_COLUMNS
    assert (df["Total"] > 0).any()
    assert df["Print"].sum() == 0
    assert df["Quantity"].dtype == int


def test_new_format_parsing():
    raw = _read_raw(NEW_CSV)
    df = parse_general_report(raw)
    assert list(df.columns) == FINAL_COLUMNS
    assert (df["Total"] > 0).any()
    assert df["Print"].sum() > 0
    assert df["Quantity"].dtype == int
