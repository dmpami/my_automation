# Распаковка архивов из писем, загрузка по url: v 0.1
# -*- encoding: utf-8 -*-
import email
import imaplib
import re
import urllib.request, urllib.parse, urllib.error
import uuid
import zipfile
import os


MyPath = os.path.abspath('')
MyEmail = ''
MyPass = ''
MySMTPserv = 'imap.yandex.ru'
MyFrom = ''

def normal_url(url_f): # Переводим из поколеченого url в нормальный 
    result = re.sub(r'=\r\n','',url_f)
    return result

def unpack_zip (directory,filename): # функция распоковки архива
    fantasy_zip = zipfile.ZipFile(os.path.join(directory, filename))
    fantasy_zip.extractall(directory)
    fantasy_zip.close()


mail = imaplib.IMAP4_SSL(MySMTPserv)
mail.login(MyEmail, MyPass)
 
mail.list() 
mail.select("sbbol") # Папка откуда берем письма

result, data = mail.search(None, "ALL") # Получаем массив со списком найденных почтовых сообщений
 
ids = data[0] # Cохраняем в переменную ids строку с номерами писем
id_list = ids.split() # Получаем массив номеров писем

id_list.reverse() # Переворачиваем что бы загружать сначало свежие 


for letter in id_list: 
    result, data = mail.fetch(letter, "(RFC822)")
    raw_email = data[0][1]
    raw_email_string = raw_email.decode('utf-8')

    email_message = email.message_from_string(raw_email_string)
    From = (email.utils.parseaddr(email_message['From']))
    if From[1] == MyFrom: 
       print ("Письмо от сбера!")
       raw_email_string
       textlookfor = r"https://bf.sberbank.ru:9443/sbns-app/download/[^\"]+" # Поиск нужного url 
       allresults = re.findall(textlookfor,raw_email_string)  # Поиск нужного url из тела письма
       allresults_str= ''.join(allresults) # Эт список переводи в строку 
       print(normal_url(allresults_str))
       try:
           url = normal_url(allresults_str)
           filename = str(uuid.uuid4()) # Генерирую временное имя архива
           filenameZip = ( filename + '.zip')
           PathFilenameZip = (MyPath + filenameZip)
           urllib.request.urlretrieve(url, PathFilenameZip) # Загрузка архива по url
           print ("-  :) Загрузил архив " + filenameZip)
           unpack_zip(MyPath , filenameZip) # Распоковка архива
           print ('--  Распоковал архив ' + filenameZip)
           os.remove(os.path.join(MyPath, filenameZip))  # Удаление архива
           print ('--- Удалил архив ' + filenameZip + '\n')
       except:
           print ("Не смог загрузить :( , возможно ссылка устарела\n")

    else:
        print ("Письмо нет от сбера\n")