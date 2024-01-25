def table_vector():
    import gspread as gs
    import pandas as pd

    gc = gs.service_account(filename='gs_vectortable.json')
    sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/1Qn-rGHE-mBaXHzVL9GHgGURV2OK0ZMRxcWerJ98qQhY/edit?usp=sharing')
    ws = sh.worksheet('Sheet1')
    df = pd.DataFrame(ws.get_all_records())
    return df
