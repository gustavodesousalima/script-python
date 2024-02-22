import pandas as pd
import os

def divide_emails(file_path, group_size=98):
    # Carrega o arquivo Excel
    df = pd.read_excel(file_path)
    
    # Extrai os e-mails da tabela
    emails = df.iloc[:, 0].dropna().tolist()  # Supondo que os e-mails estão na primeira coluna

    # Divide os e-mails em grupos de tamanho group_size
    divided_emails = [emails[i:i+group_size] for i in range(0, len(emails), group_size)]

    return divided_emails

def save_emails_to_files(divided_emails, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    for i, group in enumerate(divided_emails, 1):
        output_file = os.path.join(output_dir, f'emails-collection-group-{i}.xlsx')
        pd.DataFrame({'Emails': group}).to_excel(output_file, index=False)

# Exemplo de uso
file_path = 'C:/Users/TI003/Downloads/email-collection.xlsx'
output_dir = 'C:/Users/TI003/OneDrive - BONAMAISON COMERCIO E SERVICOS DE MOVEIS E ACESSO/Documentos/script-python'

divided_emails = divide_emails(file_path)

save_emails_to_files(divided_emails, output_dir)
