import os
import datetime as dt
import pandas_datareader.data as web
import ticker
 

FILE_PATH = "202209_ticker_list.xlsx"
codes = ticker.get_topix_codes(FILE_PATH) 
error_list=[]
for code in codes:
    print(code)
    #銘柄コード入力(7177はGMO-APです。)
    ticker_symbol=code[0]
    stock_name=code[1]
    industry33=code[2]
    industry17=code[3]
    size=code[4]
    ticker_symbol_dr=ticker_symbol + ".T"
    
    #2022-01-01以降の株価取得
    start='2022-01-01'
    end = dt.date.today()
    
    #データ取得
    try:
        df = web.DataReader(ticker_symbol_dr, data_source='yahoo', start=start,end=end)
    except:
        print("error in " + ticker_symbol)
        error_list.append(ticker_symbol)

    #2列目に銘柄コード追加
    df.insert(0, "code", ticker_symbol, allow_duplicates=False)
    df.insert(1, "name", stock_name, allow_duplicates=False)
    df.insert(2, "industry33", industry33, allow_duplicates=False)
    df.insert(3, "industry17", industry17, allow_duplicates=False)
    df.insert(4, "size", size, allow_duplicates=False)
    print(df)
    #csv保存
    df.to_csv( os.path.dirname(__file__) + 'output/stock_data_'+ ticker_symbol + '.csv')   

print("there was error on ") 
print(error_list) 
    



