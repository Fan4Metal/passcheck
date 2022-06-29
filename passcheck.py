import os, sys
from datetime import datetime
from subprocess import check_output, STDOUT
from alive_progress import alive_bar
import tkinter as tk
from tkinter import filedialog


def get_logins_list():
    result = []
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(initialdir=".", title="Выберите файл с логинами")
    with open(file_path, 'r') as f:
        logins = f.readlines()
    for line in logins:
        result.append(line.strip())
    result = list(filter(None, result))
    return result


def get_passdate(login:str):
    result = []
    try:
        out = check_output(["net", "user", "/domain", login], stderr=STDOUT).decode("866")
    except:
        return result
    text = out.split(sep="\n")
    if text[2] == 'Не найдено имя пользователя.':
        return result
    result.append(text[3][39:].strip())
    result.append(text[10][39:].strip())
    result.append(text[11][39:].strip())
    return result


print("Проверка паролей:")
print("")
logins = get_logins_list()
with alive_bar(len(logins), spinner="classic", enrich_print=False, stats=False, elapsed=False, bar="classic2", length=50) as bar:
    for login in logins:
        if not login:
            bar()
            continue           
        bar.text(login)
        dates = get_passdate(login)
        if not dates:
            print(f'{login:40} - неверный логин')
            bar()
            continue
        if dates[2] == "Никогда":
            bar()
            continue
        date_pass_set = datetime.strptime(dates[1], '%d.%m.%Y %H:%M:%S')
        date_pass_exp = datetime.strptime(dates[2], '%d.%m.%Y %H:%M:%S')
        now = datetime.today()
        days_left = date_pass_exp - now
        if days_left.days < 5 and days_left.days >0 :
            print(f'{dates[0]:40} - {days_left.days} дн. до просрочки')
        elif days_left.days < 0:
            print(f'{dates[0]:40} - пароль просрочен на {-days_left.days} дн.')
        bar()
print("")
os.system("pause")
sys.exit()
