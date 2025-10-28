import os
import win32com.client

download_folder = r"C:\Users\pedro_moraes\Downloads"
target_keyword = "ENC: Teste Impressora" # Busca por esse assunto

outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6: Caixa de Entrada

def salvar_pdfs(message, folder_name):
    for i in range(1, message.Attachments.Count + 1):
        att = message.Attachments.Item(i)
        filename = att.FileName.lower()

        if filename.endswith(".msg"): # .msg = Item do Outlook
            temp_path = os.path.join(download_folder, att.FileName)
            att.SaveAsFile(temp_path)
            try:
                msg_item = outlook.Application.CreateItemFromTemplate(temp_path) # Acessa o item do outlook
                for j in range(1, msg_item.Attachments.Count + 1):
                    inner_att = msg_item.Attachments.Item(j)
                    if inner_att.FileName.lower().endswith(".pdf"):
                        inner_save = os.path.join(download_folder, inner_att.FileName)
                        inner_att.SaveAsFile(inner_save)
                        print(f"PDF baixado de '{folder_name}': {inner_att.FileName}")
            except Exception as e:
                print(f"Erro ao abrir .MSG '{att.FileName}': {e}")
            finally:
                try:
                    os.remove(temp_path)
                except:
                    pass

def percorrer_pastas(folder):
    try:
        for message in folder.Items:
            if message.Class == 43 and target_keyword in (message.Subject or ""):
                salvar_pdfs(message, folder.Name)
    except Exception as e:
        print(f"Erro na pasta '{folder.Name}': {e}")

percorrer_pastas(inbox)
