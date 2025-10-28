import os
import win32com.client

download_folder = r"C:\Users\pedro_moraes\Downloads"

outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # Caixa de Entrada

target_keyword = "ENC: Teste Impressora"

def search_folder(folder):
    try:
        messages = folder.Items
        for message in messages:
            if message.Class == 43:  # MailItem
                if target_keyword in message.Subject:
                    for i in range(1, message.Attachments.Count + 1):
                        attachment = message.Attachments.Item(i)
                        if attachment.FileName.lower().endswith(".msg"): # .msg = Item do outlook
                            save_path = os.path.join(
                                download_folder, attachment.FileName
                            )
                            attachment.SaveAsFile(save_path)
                            print(
                                f"Arquivo baixado de '{folder.Name}': {attachment.FileName}"
                            )
    except Exception as e:
        print(f"Erro ao baixar arquivo: '{folder.Name}': {str(e)}")

search_folder(inbox)
