# popup.py
import tkinter as tk
from tkinter import messagebox

root = tk.Tk()
root.withdraw()  # メインウィンドウを非表示にする
messagebox.showinfo("メッセージ", "kaneko Hello")
root.destroy()


import win32com.client  # pythonを使うためimport

# 検索したい送信元メールアドレス
target_address = "o365mc@microsoft.com"

# Outlookアプリケーションに接続
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

messages = inbox.Items
found = False

for message in messages:
    try:
        if message.SenderEmailAddress.lower() == target_address.lower():
            print(f"メール件名: {message.Subject}")
            found = True
    except AttributeError:
        continue  # メール以外のアイテム（会議通知など）はスキップ

messagebox.showinfo("メッセージ", "kaneko Hello")

if not found:
    print("指定アドレスからのメールは見つかりませんでした。")



