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
TEXT_FILE_NAMES = glob.glob(os.path.join(CURRENT_DIR,'*.csv'))

for BONGS in EXCEL_FILE_NAMES:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(BONGS)
    wb.SaveAs(BONGS + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
    wb.Close()  # FileFormat = 56 is for .xls extension
    excel.Application.Quit()
    df = pd.read_excel(BONGS + 'x', header=None)
    os.remove(BONGS + 'x')

    Just_excel_name = os.path.splitext(os.path.basename(BONGS))[0]                  #FILE_NAME 준비
    SECOND_SIDE = df.iloc[1,4]                                                      #2차면 이름
    if SECOND_SIDE.endswith('91'):                                             #1차면 이름
        x =  df[3].str.match('[A-Za-z0-9]*90$')
        Nr_of_90 = len(df[x==True])
        if Nr_of_90 >0 :
            TRUE_OR_FALSE = df[3].str.endswith("90")
            FIRST_SIDE = df.iloc[df[TRUE_OR_FALSE == True].index[0], 3]
        else:
            FIRST_SIDE = "NO_FIRST_SIDE"
    elif SECOND_SIDE.endswith("S"):
        x = df[3].str.match('[A-Za-z0-9]*S1$')
        Nr_of_S1 = len(df[x == True])
        if Nr_of_S1 >0:
            TRUE_OR_FALSE = df[3].str.endswith("S1")
            FIRST_SIDE = df.iloc[df[TRUE_OR_FALSE == True].index[0], 3]
        else: FIRST_SIDE = "NO_FIRST_SIDE"

        # FIRST_SIDE = SECOND_SIDE + "1"
    else:
        FIRST_SIDE = "NO_FIRST_SIDE"

    df[3] = df[3].str.replace('<', '')
    df[3] = df[3].str.replace('>', '')
    df[3] = df[3].str.replace('(', '')
    df[3] = df[3].str.replace(')', '')
    df[3] = df[3].str.replace('.', ',')     #ES BOM에 '.'인 것들이 있어서 추가해줌
    df[3] = df[3].str.replace(FIRST_SIDE, '')

    df2 = df.iloc[8:,1:4]

    df2 = df2.rename(columns = {1:'SMD_Nr',2:'Item',3:'Ref#'})
    df2.loc[df2['SMD_Nr'] == 1,'SMD_Nr'] = SECOND_SIDE
    df2.loc[df2['SMD_Nr'] == 2, 'SMD_Nr'] = FIRST_SIDE

    df2['SMD_Nr'].fillna(method='ffill', inplace = True)
    df2['Item'].fillna(method='ffill', inplace = True)
    df2=df2.dropna()

    to_find = df2['Ref#'].str.match('^[A-Za-z0-9]{10,13}$')
    df3 = pd.DataFrame({'BOM_part': df2[to_find]['Ref#'].tolist(),       # part#를 찾음. 10~13길이의 문자와 숫자의 조합. 특수문자 제외
                        'Description': df2.loc[df2[to_find].index + 1]['Ref#'].tolist()       #part#바로 다음의 index를 list로 만듦
                        })

    df4 = pd.merge(df2,df3, left_on='Ref#', right_on='BOM_part', how= 'left')
    df4.drop(df4[df4['Ref#'].isin(df4['Description'])].index, inplace = True)
    df4=df4.drop_duplicates()
    df4.loc[df4['Ref#'] == df4['BOM_part'], 'Ref#'] = ''
    df4['BOM_part'].fillna(method='ffill', inplace = True)
    df4['Description'].fillna(method='ffill', inplace = True)

    m = []                                                         #공백이라서 dropna가 안먹혔다. --> Ref#가 공백이 부분 index를 찾아 list로 만든 후 해당 index의 row를 한번에 drop시킴
    for n in df4['Ref#'].index:
        if df4.loc[n,'Ref#']=='':
            m = m+[n]
    df4=df4.drop(m)

    df4 = df4.set_index(['SMD_Nr','Item','BOM_part', 'Description']).apply(lambda x: x.str.split(',').explode()).reset_index()
    df4 = df4.dropna()                                             #V0에서 추가: FIRST_NAME인 것을 공백으로 만들고 바로 다음행의 description이 계속 남아서..

    P = []                                                         # 공백이라서 dropna가 안먹혔다. --> Ref#가 공백이 부분 index를 찾아 list로 만든 후 해당 index의 row를 한번에 drop시킴
    for n in df4['Ref#'].index:
        if df4.loc[n, 'Ref#'] == '':
            P = P + [n]
    df4=df4.drop(P)

    # 'Ref#'에 필요없는 문자열 삭제
    BOM = df4.loc[df4['Ref#'] == 'BOM', 'Ref#'].index  #BOM에 Ref# 자리에 'BOM' string이 들어간 경우가 있었다. 그래서 이놈 삭제
    PCB = df4.loc[df4['Ref#'] == 'PCB', 'Ref#'].index  # BOM에 Ref# 자리에 'PCB' string이 들어간 경우가 있었다. 그래서 이놈 삭제
    PC_Board = df4.loc[df4['Ref#'] == 'PC-Board', 'Ref#'].index  # BOM에 Ref# 자리에 'PC-Board' string이 들어간 경우가 있었다. 그래서 이놈 삭제
    Main_PCB = df4.loc[df4['Ref#'] == 'Main PCB', 'Ref#'].index  # BOM에 Ref# 자리에 'MAIN PCB' string이 들어간 경우가 있었다. 그래서 이놈 삭제
    df4['Ref#'] = df4['Ref#'].str.replace(' ', '')  # BOM에 Ref# 자리에 whitespace가 string에 들어간 경우가 있었다. 그래서 이놈 삭제
    df4.drop(BOM, inplace=True)
    df4.drop(PCB, inplace=True)
    df4.drop(PC_Board, inplace=True)
    df4.drop(Main_PCB, inplace=True)
    # df4.drop(Whitespace, inplace=True)

    df4['Ref#'] = df4['Ref#'].str.upper()              #Ref#가 소문자가 포함된게 있더라 그래서 모두 대문자 처리

    if FIRST_SIDE == "NO_FIRST_SIDE":
        df6 = pd.read_csv(SECOND_SIDE + ".csv", usecols=range(7), header=None)  # 2ndside txt file import
        df6 = df6[[0, 2, 1, 6]]
        # df6[0] = df6[0].str.extract("^[\w\.\s\\\]*\\\\(.*)")  # folder 명 IBU2.0에서 .을 못찾아서 추가해줌 / folder에 공백 못찾아서 \s 추가해줌
        # df6[2] = df6[2].str.extract("^[\w\.\s\\\]*\\\\(.*)")
        df6 = df6.rename(columns={0: 'List_side', 2: 'Ref#', 1: 'List_part', 6: 'List_omitted?'})
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

    dfmerged = pd.merge(df4, df7, on='Ref#', how = 'outer')       #BOM과 Placement list merging
    dfmerged['Comparison'] = np.where(dfmerged['BOM_part']==dfmerged['List_part'], 'OK', 'NG')
    dfmerged.dropna()
    dfmerged = dfmerged.sort_values(by=['Comparison'], ascending=True, na_position='first')     #'Comparison' column을 기준으로 오름차순. NaN이 가장 위로!

    Comparison_Nr = dfmerged['Comparison'].to_list()                                            #'Comparison' result를 밖에 써주기 위해
    OK_Nr = Comparison_Nr.count('OK')
    NG_Nr = Comparison_Nr.count('NG')
    OK_Nr_name = "OK %d, NG %d" % (OK_Nr,NG_Nr)

    #마지막에 시각적으로 좋게 하기 위해 열 순서를 변경
    dfmerged = dfmerged[['Item','Description', 'SMD_Nr', 'BOM_part','Ref#','List_part','List_side','List_omitted?', 'Comparison']]


    dfmerged.to_excel(current_time+ '_' +SECOND_SIDE+ '_Result_' +OK_Nr_name +'.xlsx')
    print(SECOND_SIDE, "_completed")
print("FINISHED")