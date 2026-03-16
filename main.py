from loguru import logger
import re
import pandas as pd

data = []

logger.info("читать файл   RPABank")

with open('RpaBank_report.txt', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for line in lines:
    parts = line.split()

    number = parts[0]
    index = parts[1]

    '''localDateTime + Transaction ID'''
    datetime_id = parts[2]
    local_datetime = datetime_id[:8]
    transaction_id = datetime_id[8:]

    '''Amount + Currency + Card Number'''
    amount_currency_card = parts[3]

    amount = re.findall(r'\d+\.\d+', amount_currency_card)[0]
    currency = re.findall(r"[A-Z]{3}", amount_currency_card)[0]
    card_number = re.findall(r'\d{6,}', amount_currency_card)[0]

    terminal_id = parts[0]

    data.append([
        number,
        index,
        local_datetime,
        transaction_id,
        amount,
        currency,
        card_number,
        terminal_id
    ])


columns = [
    "Number",
    "Index",
    "Local DateTime",
    "Transaction ID",
    "Transaction Amount",
    "Currency",
    "Card Number",
    "Terminal ID"
]

df_rpa = pd.DataFrame(data, columns=columns)
df_rpa.to_excel('RPABank_report.xlsx', index=False)





'''2. Pindodo_report'''

second_data = []
transaction = {}

logger.info("читать файл Pindodo_report")

with open('Pindodo_report.txt', 'r', encoding="utf-8") as f:
    pin_lines = f.readlines()

for li in pin_lines:
    l = li.strip()

    if not l or "---" in l:
        continue

    p = l.split()

    # Transaction Date
    if p[:2] == ['Transaction', 'Date']:
        if transaction:
            second_data.append(transaction)
            transaction = {}

        transaction["Transaction Date"] = p[-1]

    # Transaction Amount
    elif p[:2] == ["Transaction", "Amount"]:
        transaction['Transaction Amount'] = p[-1]


    # Transaction Currency
    elif p[:2] == ['Transaction', 'Currency']:
        transaction['Transaction Currency'] = p[-1]


    # - Retrieval Reference Number
    elif p[:3] == ["Retrieval", "Reference", "Number"]:
        transaction['Retrieval Referance Number'] = p[-1]


    # - Card Acceptor Terminal ID
    elif p[:4] == ['Card', 'Acceptor', 'Terminal', 'ID']:
        transaction['Card Acceptor Terminal ID'] = p[-1]

if transaction:
    second_data.append(transaction)

df_pindodo = pd.DataFrame(second_data)

print(df_pindodo)

df_pindodo.to_excel('Pindodo_report.xlsx', index=False)




''' Сверка'''


merged = pd.merge(df_rpa, df_pindodo, left_on=['Transaction Amount', 'Currency'],
                   right_on=['Transaction Amount', 'Transaction Currency'],
                   how='inner', indicator=True)

success = merged[merged['_merge'] == 'both']
rpa_fail = merged[merged['_merge'] == 'left_only']
pindodo_fail = merged[merged['_merge'] == 'right_only']

print("success:", len(success))
print("rpa_fail:", len(rpa_fail))
print("pindodo_fail:", len(pindodo_fail))


with pd.ExcelWriter('result.xlsx') as writer:
    success.to_excel(writer, sheet_name='Успешные', index=False)
    rpa_fail.to_excel(writer, sheet_name='RPABank_неуспешные', index=False)
    pindodo_fail.to_excel(writer, sheet_name='Pindodo_неумпешные', index=False)







