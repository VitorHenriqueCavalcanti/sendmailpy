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

# criando a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

# Credenciando e autorizando acesso ao GSheets
gc = gspread.service_account(filename=CLIENT)
credentials = ServiceAccountCredentials.from_json_keyfile_name(CLIENT,scopes=SCOPE)
gc = gspread.authorize(credentials)

# Verificando a data atual e criando string para realizar busca nas células da planilha
todaytime = datetime.today()
day_ago = todaytime - timedelta(days=1)
data_req = day_ago.strftime('%d/%m 18:00')

# Data título do email
data_mail = day_ago.strftime('%d/%m')

# Abrindo Worksheet
gsheet = gc.open_by_key(KEY_ENTER)
wks = gsheet.get_worksheet(0)

# Pegando valores da Planilha
df = pd.DataFrame(data=wks.get_all_values())
result = df.loc[df[2] == data_req]
df.to_html('data.html')

# Declarando variáveis para inserir nos campos do e-mail
op = str([result[14].values[0]]).replace('[','').replace(']','').replace("'",'')
init = str([result[2].values[0]]).replace('[','').replace(']','').replace("'",'')
fim = str([result[3].values[0]]).replace('[','').replace(']','').replace("'",'')
ta_md = str([result[4].values[0]]).replace('[','').replace(']','').replace("'",'')
ta_me = str([result[5].values[0]]).replace('[','').replace(']','').replace("'",'')
insp_ta = str([result[6].values[0]]).replace('[','').replace(']','').replace("'",'')
insp_elf = str([result[7].values[0]]).replace('[','').replace(']','').replace("'",'')
insp_ele_qcm = str([result[8].values[0]]).replace('[','').replace(']','').replace("'",'')
insp_ele_gmg = str([result[9].values[0]]).replace('[','').replace(']','').replace("'",'')
insp_est_ta = str([result[10].values[0]]).replace('[','').replace(']','').replace("'",'')
obs_ronda = str([result[12].values[0]]).replace('[','').replace(']','').replace("'",'')
footage_ronda = str([result[13].values[0]]).replace('[','').replace(']','').replace("'",'')
footage_ronda_anomaly = str([result[17].values[0]]).replace('[','').replace(']','').replace("'",'')

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

# Criando E-mail
email = outlook.CreateItem(0)

# Configurando as informações

email.To = 'vhenriquecavalcanti@gmail.com'
email.Subject = str('Relatório de Ronda Noturna ' + data_mail)
email.HTMLBody = f"""
<html>
  <head>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@100;300;400&display=swap" rel="stylesheet">
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  </head>
  <body style=" padding: 10px !important; background-color: #ffffff; font-family: sans-serif; -webkit-font-smoothing: antialiased; font-size: 14px; line-height: 1.4; margin: 0; padding: 0; -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100%; width: 100%;">
    <table style="background-color: rgb(255, 255, 255);border-collapse: separate; width: 100%;"cellpadding="0" cellspacing="0">
      <tr style="display: block;margin-left: auto;margin-right: auto;">
        <td style="font-family: sans-serif; font-size: 14px; vertical-align: top;">&nbsp;</td>
        <td style="display: block; margin: 0 auto !important; max-width: 100%; width: 100%;  height: 100%; font-family: sans-serif; font-size: 14px;" class="container">
          <div style="box-sizing: border-box; display: block; margin: 0 auto; max-width: 580px; height: 800px; padding: 10px; padding: 0 !important; " class="content">

            <!-- START CENTERED WHITE CONTAINER -->
            <table  style="border-collapse: separate; width: 100%;" role="presentation" class="main">

              <!-- START MAIN CONTENT AREA -->
              <tr style="font-family: Montserrat;">
                <td style=" font-size: 14px; vertical-align: top;" class="wrapper">
                  <table style="border-collapse: separate; width: 100%;" role="presentation" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td style=" font-size: 14px; vertical-align: top;">
                        <p style=" font-family: 'Poppins', sans-serif; font-size: 16px; text-align: center;font-weight: 400;">Relatório de ronda noturna -  </p>
                        <table style="border-collapse: separate; width: 100%;" role="presentation" border="0" cellpadding="0" cellspacing="0" class="btn btn-primary">
                          <tbody>
                            <tr>
                              <td style="font-family: sans-serif; font-size: 14px; vertical-align: top;" align="left">
                                <table style="border-collapse: separate; width: 100%;" role="presentation" border="0" cellpadding="0" cellspacing="0">
                                    <tbody>
                                        <br>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">Operador: {op}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">Data de início: {init}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">Data de término: {fim}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">TA's Margem Direita(MD): {ta_md}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">TA's Margem Esquerda(ME): {ta_me}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">Inspeção visual dos painéis dos TA's: {insp_ta}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">Inspeção visual dos painéis das ELF's: {insp_elf}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">Inspeção de todos elementos nas salas QCM's: {insp_ele_qcm}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">Inspeção de todos elementos nas salas GMG's: {insp_ele_gmg}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">Inspeção de todos elementos nas estruturas dos TA's (caixa de bombas, extravasores, refletores, portões, telas...) estão presentes e em perfeito funcionamento: {insp_est_ta}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">Relatou alguma anormalidade na ronda: {obs_ronda}</p>
                                        <p style="padding: 5px; font-family: 'Poppins', sans-serif; font-weight: 400;">Fotos {footage_ronda, footage_ronda_anomaly}</p>
                                    </tbody>
                                </table>
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>

            <!-- END MAIN CONTENT AREA -->
            </table>
            <!-- END CENTERED WHITE CONTAINER -->

            <!-- START FOOTER -->
            <div class="footer">
              <table style="border-collapse: separate; width: 100%;" role="presentation" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td style="text-align: center; padding-bottom: 10px; padding-top: 10px; font-size: 14px;" class="content-block">
                    <span style="font-family: 'Poppins', sans-serif;font-size: 11px;color: #383838;"class="apple-link">Este é um e-mail automático. Favor não respondê-lo</span>
                  </td>
                </tr>
                <tr style="text-align: center;">
                  <td style="padding-bottom: 10px; padding-top: 10px; font-size: 14px;">
                      <a style="font-family: 'Poppins', sans-serif; text-decoration: none;color: #383838;font-size: 12px;" href="http://cemservices.com.br">C&M Serviços Especializados</a>.
                  </td>
                </tr>
                <td style=" font-size: 14px; vertical-align: top;">
                    <br>
                    <br>
                    <img style="display: block; margin-left: auto; margin-right: auto;border: none; width: 20%;"src="https://cemservices.com.br/wp-content/uploads/2018/11/logo-cem-300x120.png">
                </td>
              </table>
            </div>
            <!-- END FOOTER -->

          </div>
        </td>
      </tr>
    </table>
  </body>
</html>
"""

email.Send()
print('Email Enviado!')
