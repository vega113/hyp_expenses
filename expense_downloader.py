import os
import requests
import pandas as pd

def download_expense(link, file_name, folder, cookies=None):
    response = requests.get(link, cookies=cookies)

    if response.status_code == 200:
        os.makedirs(folder, exist_ok=True)
        with open(os.path.join(folder, file_name), 'wb') as f:
            f.write(response.content)
    else:
        print(f"Failed to download file from {link}")

def parse_and_download_expenses(file_path, cookies=None):
    df = pd.read_excel(file_path)

    for _, row in df.iterrows():
        try:
            document_date = pd.to_datetime(row.iloc[4])
        except:
            document_date = pd.to_datetime(row.iloc[3])

        expense_description = row.iloc[9].replace('.', '_')
        total_amount = str(row.iloc[6]).replace('.', '_')
        provider_name = str(row.iloc[7]).replace('.', '_')
        link = row.iloc[18]
        file_extension = link.split('.')[-1]

        file_name = f"{document_date.strftime('%Y-%m-%d')}_{provider_name}_{expense_description}_{total_amount}.{file_extension}"
        folder = document_date.strftime('Expenses 2022')

        download_expense(link, file_name, folder, cookies)
        print(f"Downloaded: {file_name} to folder {folder}")

if __name__ == "__main__":
    file_path = "/Users/yuri.zelikov/devroot/hyp_expenses/Expenses.xlsx"
    cookies = {'cookie': 'session=d7e713f03be4ef73fadc5e5c156cc0c8a499225b; _gcl_au=1.1.647062654.1686075059; _hjSessionUser_90857=eyJpZCI6IjViMzZiNGYxLTIzOGYtNTU5ZC1hNzVlLTYwYTUxYTA1ZDk4NiIsImNyZWF0ZWQiOjE2ODYwNzUwNTk2MTEsImV4aXN0aW5nIjpmYWxzZX0=; _hjFirstSeen=1; _hjIncludedInSessionSample_90857=0; _hjSession_90857=eyJpZCI6ImEyMjIyYTU3LTIzODMtNGU4YS04NTE1LTc5NWM0ZDhhZTg5YSIsImNyZWF0ZWQiOjE2ODYwNzUwNTk2MTUsImluU2FtcGxlIjpmYWxzZX0=; _hjAbsoluteSessionInProgress=0; srcNew=346538393661333531626636303365323533316532313036636431353734316239323134383839623861663635316430336663623463613936343665336233383861313064353961343830373134663433343439343738333037386132653839623464336162633361396261663064393132313665653162366632653764396475347a455a507a4e5a5a6c68364a68794b6d3663366245596858786c635366754c312f4a7a7a314a3452595250745a514d6c6f6a3353544665734179486b4548574e4158635a46725667486633614979534c524b324c6d6937702b57505a5441624d69463141735a30575865695648306c4642364a527a78727a35583157532b514549567a44597873426f63764a6c697a30613071773d3d; _gid=GA1.3.590697679.1686075060; _ga=GA1.1.868641626.1686075060; _ga_H6E4KBL9TY=GS1.1.1686075059.1.1.1686075099.20.0.0; kdm_2fa_validation=b01a91bbe5dbb3eebc1da5122ae11105f3ccfbd6c789fec8fec575a240075087; kdm_a_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczpcL1wvZmlsZXMuZXpjb3VudC5jby5pbFwvIiwiYXVkIjoiaHR0cHM6XC9cL2ZpbGVzLmV6Y291bnQuY28uaWxcLyIsImlhdCI6MTY4NjA3NTEwMCwibmJmIjoxNjg2MDc1MTAwLCJwYXlsb2FkIjoiYTM1MDFiZjg4NTViMzEzOThlZjY0ZjEyNGRjMmY0OTdlZTQ2OTIyMmZmMDgyNWY0MjZiZDBhYWJhMjE0ZTI2MzFmM2RkOWNkZTBjMTM2MWZhYjA4ZTE4MGY1ZjU1YmVhYTJlNTQyNjMzMjYxZmM4N2RlNmExMDVmYjkyZGE3MjVuVVdIUW9aTmh4M0xQelBOTFFuc0xxOVlQRTVlb0l6NUUzenprREtMU2loZG9JZ0Nma1wva2JObHVXbm5TSXZGdHJpS3RcL09PQ2I5VHN2RDFLV0dOOGY3NFwvenpNazhtYjBsXC9BbDFXRVZXSUtIVEJYWGNLMFlVbnI5endoMFVtdjN1dkVyb0V6NUgxeDJtcWd6bVBkcVF4dUZVc2tYZE15cVB1ODF3T3g2T29oRkJtaWxjeFVtdnA5bEpYWkZrTW5GQWR0MDNKeUVZaFl2Y1VRZFdRdDJ3TTJlRDZxT2JUT0gwNG1XQVwvK2JmRXJNTGkwaVFQMFY0MkhDY3lXWnZVQ3dHOEUwandmR1B6ZzlUeG82THZsZHVRPT0iLCJhbnQiOiI0YzAwN2MzZDJjNGFmNWRkM2I5ZTAwNzU5MTExZWU0ZTEzZjdjODc2In0.ZDjw9oPCYjG3Wb5emEICFExvd7PxY5Zzzd4z4dRZwvs'} # replace <token> with the actual authorization token
    parse_and_download_expenses(file_path, cookies)
