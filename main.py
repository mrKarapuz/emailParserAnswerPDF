import ssl
from MySQLdb._exceptions import OperationalError
from aiogram import Bot
import imaplib, smtplib, os, pdfplumber, MySQLdb, requests, shutil
from pdfminer.pdfparser import PDFSyntaxError
from time import sleep, strftime
from zipfile import ZipFile
from email import message_from_bytes
from email.header import decode_header, make_header
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

print('Обработчик запущен')
print('------------------')

SYMVOL = ('<', '>', ':', '"', '/', '\\', '|', '?', '*', ',')

# Функция отправки сообщения в группу в телеграме
def telegram_bot_send():
    token = '1901007184:AAHZ4DCUsqD9MRCzAdLYCgrHscz6erCy4cU'
    bot = Bot(token=token)
    chat_id = '-1001572207430'
    text = win_code
    url = f'https://api.telegram.org/bot{token}/sendMessage?chat_id={chat_id}&parse_mode=Markdown&text={text}'
    requests.get(url)

# Функция, которая срабатывает если не подключились к почте или не подключились к базе данных
def error_and_exit():
    print('---------------------')
    print(f'Выход через 25 сек...')
    sleep(10)
    print(f'Выход через 15 сек...')
    sleep(10)
    print(f'Выход через 5 сек...')
    sleep(5)
    exit()

# Функция загрузки файлов
def get_attachments(msg):
    sleep(1)
    global myzip
    cur.execute(f"SELECT win_code FROM `basa` WHERE win_code='{win_code}'")
    true_wincode = cur.fetchone()
    try:
        if true_wincode[0] == win_code:
            return 0
    except:
        print('Win cod нету в базе данных, сохраняем архив')
        
    with ZipFile(zip_dir +'.zip', 'w') as myzip:
        for part in msg.walk():
            if part.get_content_maintype()=='multipart':
                continue
            sleep(0.1)
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()
            if bool(fileName):
                sleep(0.1)
                filePath = os.path.join(str(make_header(decode_header(fileName))))
                with open(filePath,'wb') as f:
                    f.write(part.get_payload(decode=True))
                    myzip.write(filePath)
                os.remove(filePath)
    if len(myzip.infolist()) == 0:
        print('Архив пуст')
        os.remove(zip_dir + '.zip')
    else: 
        # Отправляем сообщение в группу телеграм
        telegram_bot_send()
        # Работа с базой данных
        Date = strftime('%Y-%m-%d')
        Email = from_user[from_user.find('<') + 1 :from_user.find('>')]
        status_saved = 'saved_archive'
        cur.execute(f'INSERT INTO basa (date, email, win_code, status) VALUES(%s, %s, %s, %s)', (Date, Email, win_code, status_saved))
        db.commit()

# Функция поиска Win кода на сервере
def serach_win_code_in_file(_win_code):
    global filename
    try:
        for elem in os.listdir(DIR_OF_UNTREATED_FILES):
            if len(elem) > 25 and str(elem[-4:]) == '.pdf':
                filename = DIR_OF_UNTREATED_FILES + '/' + elem
                text = ''
                pages = []
                with pdfplumber.open(filename) as pdf:
                    for i, page in enumerate(pdf.pages):
                        text = text+str(page.extract_text())
                        pages.append(page)
                win_code_in_file = text[(text.find('номер кузова:')) + 14 :(text.find('номер кузова:')) + 31]
                if _win_code == win_code_in_file : return True
            else: continue
    except FileNotFoundError:
        print(f'''Не удается найти путь к серверу: 
{DIR_OF_UNTREATED_FILES}  
Проверьте данные и перезапустите программу.''')
        error_and_exit()

# Функция отправки сообщения пользователю
def send_message_to_user(email_user):
    message = MIMEMultipart()
    message['From'] = LOGIN
    message['To'] = email_user
    message['Subject'] = 'Re:' + subject_of_mail_to_re
    if TEXT_MESSAGE:
        body = TEXT_MESSAGE
    else:
        body = '''Отримайте Вашу довідку'''
    name_to_file_on_mail = filename[filename.rfind('/') + 1 :]
    message.attach(MIMEText(body, 'plain'))
    pdfname = filename
    binary_pdf = open(pdfname, 'rb')
    payload = MIMEBase('application', 'octate-stream', Name=name_to_file_on_mail)
    payload.set_payload((binary_pdf).read())
    encoders.encode_base64(payload)
    payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
    message.attach(payload)
    TYPE_OF_EMAIL = 'smtp.' + LOGIN[LOGIN.find('@')+1:]
    session = smtplib.SMTP(TYPE_OF_EMAIL, 587)
    sleep(1)
    session.starttls()
    session.login(LOGIN, PASSWORD)
    text = message.as_string()
    session.sendmail(LOGIN, email_user, text)
    session.quit()
    try:
        cur.execute(f"UPDATE `basa` SET `status`= 'message_send' WHERE win_code = '{win_code}'")
    except:
        pass
    try:
        db.commit()
    except:
        try:
            db.commit()
        except:
            pass

def connection_mail():
    global con
    # Подключение к почте
    try:
        TYPE_OF_EMAIL = 'imap.' + LOGIN[LOGIN.find('@')+1:]
        con = imaplib.IMAP4_SSL(TYPE_OF_EMAIL)
        con.login(LOGIN, PASSWORD)
    except:
        print(f'''Не удается подключится к почтовому ящику, проверьте плавильность введенных вами данных в файле "data.xlsx", лист "settings":  
        Логин: {LOGIN}
        Пароль: {PASSWORD}
    Также проверьте что в настройках вашей почты включена функция доступа с помощью IMAP и разрешен доступ "небезопасных приложений" (для почты Gmail.com)''')
        error_and_exit()
    print('Начинаю обработку почты')
    print(strftime("%H:%M:%S"))

# Логин электронной почты
LOGIN = 'spravkieuro@gmail.com'
# Пароль электронной почты
PASSWORD = 'hj32l12j480i'
# Имя папки, из которой будут считываться письма, по умолчанию "Входящие"
NAME_OF_CATALOG_ON_MAIL = 'inbox'
# Папка сервера
DIR_OF_SERVER_from_base = '//192.168.88.254/Outsource/Маланюк Евро'
# Папка необработанных файлов
DIR_OF_UNTREATED_FILES = 'C:/Users/admin/Google Диск/Evro/Необработанные справки'
DIR_OF_SERVER = DIR_OF_SERVER_from_base + '/' + strftime('%Y') + '/' + strftime('%m.%Y') + '/' + strftime('%d.%m') + '/'

# Путь к папке сохранения архивов
ATTACHMENT_DIR = DIR_OF_SERVER
# Время перезагрузки сервера
TIME_TO_RELOAD = 120
# Текст сообщения, которое будет отправлено пользователюs
TEXT_MESSAGE = 'Довідка'

print(f'''
Путь к серверу: {DIR_OF_SERVER}
Интервал автоматической перезагрузки: {TIME_TO_RELOAD} сек.
Время: {strftime('%H:%M:%S')}
''')

# Список разрешенных Email адресов, от которых можно принимать сообщения
list_of_accesed_from_mail = ['goronik12@gmail.com', 'autobrokteam@gmail.com', 'alternativa1ocenka@gmail.com', 'odessabroker777@gmail.com', 'madl.llc.manager2@gmail.com', 'timoshovazhanna8891@gmail.com', 'universalautogroup2019@gmail.com', 'kleverbrok@ukr.net', 'info@mklogistic.od.ua', 'euroterminalodessa@gmail.com', 'dbrok@rastamojka.od.ua']

# Главная функция
def main_function():

    global db, cur, from_user, win_code, zip_dir, subject_of_mail, filename, subject_of_mail_to_re, DIR_OF_SERVER_DONE
    count = 0
    # Подключение к серверу
    try:
        os.listdir(DIR_OF_SERVER)
        # Папка отработанных данных
        DIR_OF_SERVER_DONE = DIR_OF_SERVER + '01 done' 
        try:
            os.listdir(DIR_OF_SERVER_DONE)
        except FileNotFoundError:
            os.mkdir(DIR_OF_SERVER_DONE)
    except:
        print('''----------------------------
    СЕРВЕР НЕДОСТУПЕН
----------------------------''')
        sleep(60)
        return 0
    

    try:
        # Все пдф из сервера перемещаем в папку "Не обработанные файлы"
        for elem in os.listdir(DIR_OF_SERVER):
            if elem[-4:] == '.pdf':
                # Убираем в конце названия отправляемого файла имя и фамилию
                os.rename(DIR_OF_SERVER + elem, DIR_OF_SERVER + elem[:elem.rfind('_')] + '.pdf')
  
        for elem in os.listdir(DIR_OF_SERVER):
            if elem[-4:] == '.pdf':    
                shutil.copy2(DIR_OF_SERVER + elem, DIR_OF_SERVER_DONE + '/' + elem)
                shutil.move(DIR_OF_SERVER + elem, DIR_OF_UNTREATED_FILES + '/' + elem)
        # Все архивы из необработанных файлов перемещаем на сервер
        for elem in os.listdir(DIR_OF_UNTREATED_FILES):
            if elem [-4:] == '.zip' or elem[-4:] == '.rar':
                shutil.move(DIR_OF_UNTREATED_FILES + '/' + elem, DIR_OF_SERVER + elem)
    except: 
        print('''----------------------------
СЕРВЕР НЕДОСТУПЕН
----------------------------''')
        sleep(60)
        return 0
    # Подключение к базе данных
    db = MySQLdb.connect(host="176.111.49.48",    
                        user="zkdqsgeo_euro",        
                        passwd="M8s4J5j2",     
                        db="zkdqsgeo_euro")
    cur = db.cursor()
    # Количество необработанных сообщений на почте
    connection_mail()
    try:
        count_of_messages = int(con.select(NAME_OF_CATALOG_ON_MAIL)[1][0].decode('utf-8'))
        sleep(1)
    except:
        print('Ошибка 1')
        return 0


    # Загружаем каждое сообщение по очереди
    for i in range(1, count_of_messages + 1):
        if i % 10 == 0:
            print('Обработано 9 сообщений')
            con.logout()
            sleep(TIME_TO_RELOAD)
            connection_mail()
            try:
                con.select(NAME_OF_CATALOG_ON_MAIL)[1][0].decode('utf-8')
                sleep(1)
            except:
                print('Ошибка 1')
                sleep(15)
                connection_mail()
                try: 
                    con.select(NAME_OF_CATALOG_ON_MAIL)[1][0].decode('utf-8')
                except:
                    continue
                
        print('Начинаю обрабатывать сообщение')
        try:
            result, data = con.fetch(f'{i}', '(RFC822)')
        except:
            print('Ошибка 2')
            sleep(15)
            connection_mail()
            try: 
                con.select(NAME_OF_CATALOG_ON_MAIL)[1][0].decode('utf-8')
                result, data = con.fetch(f'{i}', '(RFC822)')
            except:
                continue
        try:
            raw = message_from_bytes(data[0][1])
        except:
            print('Ошибка 3')
            sleep(15)
            connection_mail()
            try: 
                con.select(NAME_OF_CATALOG_ON_MAIL)[1][0].decode('utf-8')
                result, data = con.fetch(f'{i}', '(RFC822)')
                raw = message_from_bytes(data[0][1])
            except:
                continue
        # Имя и адрес отправителя
        from_user = str(make_header(decode_header(raw['FROM'])))
        accessed_mail = from_user[from_user.find('<') + 1 :from_user.find('>')]
        # Проверяем есть ли адрес в списке разрешенных
        if accessed_mail not in list_of_accesed_from_mail:
            print('Данного отправителя нет в списке разрешенных')
            try:
                con.store(f'{i}', '+FLAGS', '\\Deleted')
                count += 1
                sleep(1)
                
            except OperationalError:
                print('Не удалось удалить сообщение от неразрешенного пользователя')
                continue
        # Тема письма
        subject_of_mail = str(make_header(decode_header(raw['Subject'])))
        # Тема письма для формирования темы ответа
        subject_of_mail_to_re = subject_of_mail
        # Проверяем тему письма на запрещенные для сохранения символы
        for elem in SYMVOL:
            if elem in subject_of_mail:
                subject_of_mail = subject_of_mail.replace(elem, ' ')

        # Находим WIN-code машины в теме письма
        win_code = False
        for elem in subject_of_mail.split(' '):
            if len(elem) == 17 and (not 'i' 'o' 'q' in elem.lower()) and elem.isalnum():
                win_code = elem.upper()
                # Объявляем название архива
                zip_dir = os.path.join(ATTACHMENT_DIR, win_code)
                try:
                    os.listdir(DIR_OF_SERVER)
                except:
                    print('''----------------------------
СЕРВЕР НЕДОСТУПЕН
----------------------------''')   
                    sleep(60)
                    return 0

                    # Сохраняем архив
                try:
                    if get_attachments(raw) == 0:
                        pass
                except:
                    pass               
        
        # Если win code находится в одном из файлов
        if serach_win_code_in_file(win_code):
            filename_for_copy = filename

            # Отправляем сообщение пользователю
            send_message_to_user(accessed_mail)
            try:
                sleep(1)
                con.store(f'{i}', '+FLAGS', '\\Deleted')
                count += 1                
                sleep(1)
            except:
                print('Не удалось присвоить флаг удаления, повторная попытка через 15 сек..')
                con.logout()
                sleep(15)
                connection_mail()
                try:
                    con.store(f'{i}', '+FLAGS', '\\Deleted')
                    count += 1
                except:
                    print('Не успешная попытка')
                    pass
            print(f'Сообщение успешно обработано от пользователя: {from_user}', strftime('%b, %A, %H:%M:%S'))
            # Перемещаем файл в папку server_done
            try:
                shutil.copy2(filename, DIR_OF_SERVER_DONE + '/' + filename_for_copy[filename_for_copy.rfind('/') + 1:])
                os.replace(filename, DIR_OF_UNTREATED_FILES + '/Обработанные справки/' + filename[filename.rfind('/') + 1:])
            except FileNotFoundError:
                print(f'''Не удается удалить файл
{DIR_OF_UNTREATED_FILES} / {filename}
Проверьте данные и перезапустите программу.''')
                continue
    try:
        con.expunge()
        con.logout()
        db.close()
    except:
        print('Ошибка 6')
    print(f'Необработанных сообщений на почте: {count_of_messages - count}')
    sleep(TIME_TO_RELOAD)

# Автоматическое сканирование почты
while True:
    # Сервер спит с 20.00 до 07.00
    if int(strftime('%H')) >= 20:
        print('20.00, сервер запустится завтра в 07.00')
        sleep(39600)
        DIR_OF_SERVER = DIR_OF_SERVER_from_base + '/' + strftime('%Y') + '/' + strftime('%m.%Y') + '/' + strftime('%d.%m') + '/'
        ATTACHMENT_DIR = DIR_OF_SERVER
    # Сервер спит в воскресенье
    if strftime('%A') == 'Sunday':
        print('воскресенье, сервер запустится завтра в 07.00')
        sleep(86400)
        DIR_OF_SERVER = DIR_OF_SERVER_from_base + '/' + strftime('%Y') + '/' + strftime('%m.%Y') + '/' + strftime('%d.%m') + '/'
        ATTACHMENT_DIR = DIR_OF_SERVER
    main_function()