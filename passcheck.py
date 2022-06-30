import os, sys
from datetime import datetime
from subprocess import check_output, STDOUT
from tkinter import filedialog

from openpyxl import load_workbook
from alive_progress import alive_bar


def get_logins_list():
    result = []
    filetypes = (('Text, Excel files', '*.txt *.xlsx *.xlsm'),
                 ('Text files', '*.txt'),
                 ('Excel files', '*.xlsx *.xlsm'),
                 ('All files', '*.*')
    )
    file_path = filedialog.askopenfilename(initialdir = os.path.expanduser("~/Desktop"),
                                           title="Выберите файл с логинами",
                                           filetypes=filetypes)
    if not file_path:
        sys.exit()
    ext = os.path.splitext(file_path)[1]
    if ext == '.txt':
        with open(file_path, 'r') as f:
            logins = f.readlines()
            for login in logins:
                if login:
                    result.append(login.strip())
            result = list(filter(None, result))
            return result
    elif ext == '.xlsx' or ext == '.xlsm':
        wb = load_workbook(filename = file_path)
        try:
            sheet = wb['ВИП']
            col = sheet['E']
        except:
            print("Не удалость найти спиcок логинов в файле Excel")
            return result
        for i in range(1, len(col)):
            if col[i].value:
                login = str(col[i].value)
                login = login.strip()
                result.append(login)
        return result
    else:
        sys.exit()


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


def main():
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
            # date_pass_set = datetime.strptime(dates[1], '%d.%m.%Y %H:%M:%S')
            date_pass_exp = datetime.strptime(dates[2], '%d.%m.%Y %H:%M:%S')
            now = datetime.today()
            days_left = date_pass_exp - now
            if days_left.days < 5 and days_left.days > 0:
                print(f'{dates[0]:40} - {days_left.days} дн. до просрочки')
            elif days_left.days < 0:
                print(f'{dates[0]:40} - пароль просрочен на {-days_left.days} дн.')
            bar()
    print("")
    os.system("pause")
    sys.exit()

if __name__ == '__main__':
    main()