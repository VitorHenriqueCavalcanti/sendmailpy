import gspread
import pandas as pd
import numpy as np
import win32com.client as win32
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from datetime import timedelta

SCOPE = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

CLIENT = './credential_gspread.json'
KEY_ENTER = '14XvYkYaUhs0wBMnguWgy3D1F2WKn0UFnm0kTmRPcHVw'


# Credenciando e autorizando acesso ao GSheets
gc = gspread.service_account(filename=CLIENT)
credentials = ServiceAccountCredentials.from_json_keyfile_name(CLIENT,scopes=SCOPE)
gc = gspread.authorize(credentials)

# Verificando a data atual e criando string para realizar busca na planilha
todaytime = datetime.today()
day_ago = todaytime - timedelta(days=9)
data_req = day_ago.strftime('%d/%m 18:00')
data_mail = day_ago.strftime('%d/%m')

# Abrindo Worksheet
gsheet = gc.open_by_key(KEY_ENTER)
wks = gsheet.get_worksheet(0)


df = pd.DataFrame(data=wks.get_all_values())
result = df.loc[df[2] == data_req]




op = str([result[14].values[0]]).replace('[','').replace(']','').replace("'",'')
init = str([result[2].values[0]]).replace('[','').replace(']','')
fim = str([result[3].values[0]]).replace('[','').replace(']','')
ta_md = str([result[4].values[0]]).replace('[','').replace(']','')
ta_me = str([result[5].values[0]]).replace('[','').replace(']','')
insp_ta = str([result[6].values[0]]).replace('[','').replace(']','')
insp_elf = str([result[7].values[0]]).replace('[','').replace(']','')
insp_ele_qcm = str([result[8].values[0]]).replace('[','').replace(']','')
insp_ele_gmg = str([result[9].values[0]]).replace('[','').replace(']','')
insp_est_ta = str([result[10].values[0]]).replace('[','').replace(']','')
obs_ronda = str([result[12].values[0]]).replace('[','').replace(']','')
footage_ronda = str([result[13].values[0]]).replace('[','').replace(']','')
footage_ronda_anomaly = str([result[17].values[0]]).replace('[','').replace(']','')


print(str(op))
print(str(init))
print(str(fim))
print(str(ta_md))
print(str(ta_me))
print(str(insp_ta))
print(str(insp_elf))
print(str(insp_ele_qcm))
print(str(insp_ele_gmg))
print(str(insp_est_ta))
print(str(obs_ronda))
print(str(footage_ronda))
print(str(footage_ronda_anomaly))