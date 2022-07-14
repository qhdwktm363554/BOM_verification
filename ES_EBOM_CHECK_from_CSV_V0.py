import pandas as pd
import numpy as np
import os, glob
from datetime import datetime
import win32com.client as win32

now = datetime.now()
current_time = now.strftime('%Y-%m-%d')

pd.set_option('display.width', 320)
pd.set_option('display.max_columns',10)

CURRENT_DIR = os.getcwd()
EXCEL_FILE_NAMES = glob.glob(os.path.join(CURRENT_DIR,'*.xls'))
TEXT_FILE_NAMES = glob.glob(os.path.join(CURRENT_DIR,'*.txt'))

for BONGS in EXCEL_FILE_NAMES:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(BONGS)
    wb.SaveAs(BONGS + "X", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
    wb.Close()  # FileFormat = 56 is for .xls extension
    excel.Application.Quit()
    df = pd.read_excel(BONGS + 'X', header=None)
    os.remove(BONGS + 'x')

    Just_excel_name = os.path.splitext(os.path.basename(BONGS))[0]                  #FILE_NAME 준비
    SECOND_SIDE = df.iloc[1,5]                                                      #2차면 이름
    if SECOND_SIDE.endswith('91'):                                                  #1차면 이름 얻기
        x =  df[3].str.match('[A-Za-z0-9]*90$')
        Nr_of_90 = len(df[x==True])
        if Nr_of_90 >0 :
            TRUE_OR_FALSE = df[3].str.endswith("90")
            FIRST_SIDE = df.iloc[df[TRUE_OR_FALSE == True].index[0], 3]
        else:
            FIRST_SIDE = "NO_FIRST_SIDE"
    elif SECOND_SIDE.endswith("S"):
        FIRST_SIDE = SECOND_SIDE + "1"
    else:
        FIRST_SIDE = "NO_FIRST_SIDE"

    df2 = df.iloc[8:, 1:7]
    df2 = df2.rename(columns={1: 'SMD_Nr', 2: 'Item', 3: 'Before',6: 'Ref#'})
    df2.loc[df2['SMD_Nr'] == 1, 'SMD_Nr'] = SECOND_SIDE
    df2.loc[df2['SMD_Nr'] == 2, 'SMD_Nr'] = FIRST_SIDE
    df2['SMD_Nr'].fillna(method='ffill', inplace=True)
    df2['Item'].fillna(method='ffill', inplace=True)

    regex = '^AT=E'                                                                  #AT=E로 시작하는 것들 지우기 위해
    df2.drop(df2.loc[df2['Before'].str.match(regex) == True].index, inplace=True)

    df2['Description'] = df2['Before'].shift(1)
    df2['BOM_part'] = df2['Before'].shift(2)

    component = '^[A-Za-z0-9]{10,13}$'

    df2.loc[~df2.BOM_part.str.contains(component, na=False), 'BOM_part'] = np.nan     #component를 포함하는게 아닌 놈들을 NaN으로 변경
    df2['Description'] = df2['Description'].str.replace(component, '')
    df2['Description'] = df2['Description'].replace(r'^\s*$', np.NaN, regex=True)     #component를 공백으로 만들어주고 NaN으로 변경

    df2.BOM_part = df2.BOM_part.ffill()
    df2.Description = df2.Description.ffill()

    df2 = df2.loc[df2["Ref#"].notnull(), ["SMD_Nr", 'Item', 'Ref#', "Description","BOM_part"]]
    df2.columns = ["SMD_Nr", 'Item', 'Ref#', "Description","BOM_part"]

    #'Ref#'에 필요없는 문자열 삭제
    BOM = df2.loc[df2['Ref#'] == 'BOM', 'Ref#'].index  #BOM에 Ref# 자리에 'BOM' string이 들어간 경우가 있었다. 그래서 이놈 삭제
    PCB = df2.loc[df2['Ref#'] == 'PCB', 'Ref#'].index  # BOM에 Ref# 자리에 'PCB' string이 들어간 경우가 있었다. 그래서 이놈 삭제
    PC_Board = df2.loc[df2['Ref#'] == 'PC-Board', 'Ref#'].index  # BOM에 Ref# 자리에 'PC-Board' string이 들어간 경우가 있었다. 그래서 이놈 삭제
    Main_PCB = df2.loc[df2['Ref#'] == 'Main PCB', 'Ref#'].index  # BOM에 Ref# 자리에 'MAIN PCB' string이 들어간 경우가 있었다. 그래서 이놈 삭제
    df2['Ref#'] = df2['Ref#'].str.replace(' ', '')  # BOM에 Ref# 자리에 whitespace가 string에 들어간 경우가 있었다. 그래서 이놈 삭제
    df2.drop(BOM, inplace=True)
    df2.drop(PCB, inplace=True)
    df2.drop(PC_Board, inplace=True)
    df2.drop(Main_PCB, inplace=True)

    df4 = df2.loc[:, ['SMD_Nr', 'Item', 'Ref#', 'Description', 'BOM_part']]
    df4['Ref#'] = df4['Ref#'].str.upper()  # Ref#가 소문자가 포함된게 있더라 그래서 모두 대문자 처리

    if FIRST_SIDE == "NO_FIRST_SIDE":
        df6 = pd.read_csv(SECOND_SIDE + ".csv", usecols=range(7), header=None)  # 2ndside txt file import
        df6 = df6[[0, 1, 2, 8]]
        # df6[0] = df6[0].str.extract("^[\w\.\s\\\]*\\\\(.*)")  # folder 명 IBU2.0에서 .을 못찾아서 추가해줌 / folder에 공백 못찾아서 \s 추가해줌
        df6[2] = df6[2].str.extract("^[\w\.\s\\\]*\\\\(.*)")
        df6 = df6.rename(columns={0: 'List_side', 1: 'Ref#', 2: 'List_part', 8: 'List_omitted?'})
        df6.drop(df6.loc[df6['List_omitted?'] == True].index, inplace=True)
        df7 = df6
    else:
        df5 = pd.read_csv(FIRST_SIDE + ".csv", usecols=range(7), header=None)  # 1stside txt file import
        df5 = df5[[0, 2, 1, 6]]
        # df5[0] = df5[0].str.extract("^[\w\.\s\\\]*\\\\(.*)")  # folder 명 IBU2.0에서 .을 못찾아서 추가해줌 /folder에 공백 못찾아서 \s 추가해줌
        # df5[2] = df5[2].str.extract("^[\w\.\s\\\]*\\\\(.*)")
        df5 = df5.rename(columns={0: 'List_side', 2: 'Ref#', 1: 'List_part', 6: 'List_omitted?'})
        df5.drop(df5.loc[df5['List_omitted?'] == True].index, inplace=True)

        df6 = pd.read_csv(SECOND_SIDE + ".csv",usecols=range(7), header=None)  # 2ndside txt file import
        df6 = df6[[0, 2, 1, 6]]
        # df6[0] = df6[0].str.extract("^[\w\.\s\\\]*\\\\(.*)")  # folder 명 IBU2.0에서 .을 못찾아서 추가해줌
        # df6[2] = df6[2].str.extract("^[\w\.\s\\\]*\\\\(.*)")
        df6 = df6.rename(columns={0: 'List_side', 2: 'Ref#', 1: 'List_part', 6: 'List_omitted?'})
        df6.drop(df6.loc[df6['List_omitted?'] == True].index, inplace=True)
        df7 = pd.concat([df5, df6])  # BOM과 합치기 위해 df5,6 merging

    dfmerged = pd.merge(df4, df7, on='Ref#', how='outer')  # BOM과 Placement list merging
    dfmerged['Comparison'] = np.where(dfmerged['BOM_part'] == dfmerged['List_part'], 'OK', 'NG')
    dfmerged.dropna()
    dfmerged = dfmerged.sort_values(by=['Comparison'], ascending=True,
                                    na_position='first')  # 'Comparison' column을 기준으로 오름차순. NaN이 가장 위로!

    Comparison_Nr = dfmerged['Comparison'].to_list()  # 'Comparison' result를 밖에 써주기 위해
    OK_Nr = Comparison_Nr.count('OK')
    NG_Nr = Comparison_Nr.count('NG')
    OK_Nr_name = "OK %d, NG %d" % (OK_Nr, NG_Nr)
    
    #마지막에 시각적으로 좋게 하기 위해 열 순서를 변경
    dfmerged = dfmerged[['Item','Description', 'SMD_Nr', 'BOM_part','Ref#','List_part','List_side','List_omitted?', 'Comparison']] 

    dfmerged.to_excel(current_time + '_' + SECOND_SIDE + '_Result_' + OK_Nr_name + '.xlsx')
    print(FIRST_SIDE, "_completed")
print("FINISHED")