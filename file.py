import email
import os
import pandas as pd

# replace "path/to/email.msg" with the actual path to your MSG file
with open("path/to/email.msg", 'rb') as f:
    msg = email.message_from_binary_file(f)

for part in msg.walk():
    if part.get_content_type() == 'application/vnd.ms-excel':
        filename = part.get_filename()
        if filename.endswith('.xls') or filename.endswith('.xlsx'):
            content = part.get_payload(decode=True)
            df = pd.read_excel(content, engine='xlrd')
            print(df.head())
