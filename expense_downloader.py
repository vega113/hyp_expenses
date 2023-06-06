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
    cookies = {'cookie': ''} # replace <token> with the actual authorization token
    parse_and_download_expenses(file_path, cookies)
