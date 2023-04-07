from pickle import FALSE, TRUE
from ldap3 import Connection, Server, ALL, Tls, MODIFY_REPLACE, ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES
import ssl
import configparser
import logging
from ldap3.utils.log import set_library_log_activation_level, set_library_log_detail_level, get_detail_level_name, set_library_log_hide_sensitive_data, OFF, BASIC, NETWORK, EXTENDED, PROTOCOL
import random
import string
import xlsxwriter
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import COMMASPACE
from email import encoders


__author__ = "Muhammed Kayhan"
__version__ = "0.0.1"
__email__ = "muhammed.kayhan@turkcell.com.tr"
__maintainer__ = "Muhammed Kayhan"
__credits__ = ["Muhammed Kayhan"]
__status__ = "Production"

#loglama için parametreleri belirliyoruz

logging.basicConfig(filename='ldap_application.log', format="%(asctime)s - %(levelname)s - %(message)s", level=logging.DEBUG)
set_library_log_activation_level(logging.DEBUG)
set_library_log_detail_level(PROTOCOL)
set_library_log_hide_sensitive_data(True)

#config.ini dosyasındaki değerleri okuyoruz.

config = configparser.ConfigParser()
config.read('config.ini')

#random şifre üreten fonksiyon

def get_pass(stringLength=8):
    lettersAndDigits = string.ascii_letters + string.digits
    return ''.join((random.choice(lettersAndDigits) for i in range(stringLength)))

#Değerleri config.ini dosyasında set edebileceğimiz ldapa bağlantı açan fonksiyon

def openConnection(ldap_environment):
    ldap_ssl_connection = True
    tls_configuration = Tls(validate=ssl.CERT_REQUIRED, version=ssl.PROTOCOL_TLSv1_2)
    server = Server(host=config.get(ldap_environment, "host"), port=int(config.get(ldap_environment, "port")), use_ssl=ldap_ssl_connection, tls=tls_configuration)
    conn = Connection(server, user=config.get(ldap_environment, "user_dn"), password=config.get(ldap_environment, "password"), return_empty_attributes=TRUE)
    return conn

#LDAP'ta verilen parametrelere göre arama yapan search fonksiyonu

def ldapSearch(ldap_environment, search_base, search_filter, attributes):
    conn = openConnection(ldap_environment)
    conn.bind()
    conn.search(search_base=search_base, search_filter=search_filter, attributes=attributes, search_scope=2)
    return conn

#Şifre reset fonksiyonu

def ldapPasswordReset(ldap_environment):
    username = input("Please enter a username: ")
    print("Please choose one: \n" 
        "1- Generate new random password \n"
        "2- Set new password mannually")
    choose_password=input()
    if choose_password == '1':
        new_password=get_pass()
    elif choose_password == '2':
        new_password = input("Enter new password: ")
    else:
        print("Invalid option. Please try again")
    print(new_password)
    conn = ldapSearch(ldap_environment, search_base=config.get("password_reset", "base"), search_filter=f'(uid={username})')
    entries = conn.entries
    if len(entries) == 1:
        conn.modify(entries[0].entry_dn, {'userPassword' : [(MODIFY_REPLACE, [new_password])]})
        print("Successfully reset password!")
    elif len(entries) < 1:
        print("No entry found")
    else:
        print("There is more than 1 entry to given usearname string")
    return conn

#Servis Hesabı raporu için excel üreten fonksiyon.

def serviceUserPasswordChangePeriod(ldap_environment, ldap_task):
    search_base = config.get(ldap_task, "base")
    account_filter = config.get(ldap_task, "account_filter")
    user_filter = config.get(ldap_task, "user_filter")
    account_attributes = ['uid', 'mail', 'manager', 'pwdchangedtime', 'modifytimestamp']
    user_attributes = ['uid', 'cn', 'unitcode', 'functionalgroupname', 'divisionname', 'unitname', 'positionname']
    conn = ldapSearch(ldap_environment, search_base, account_filter, account_attributes)
    account_entries = conn.entries
    workbook = xlsxwriter.Workbook('Service_Users.xlsx')
    worksheet = workbook.add_worksheet('SERVICE USERS')
    integer_format = workbook.add_format({'num_format': '0'})
    
    row = 0
    col = 0

    column_titles = ["SERVIS HESABI", "LDAP KULLANICI ADI", "AD-SOYAD", "UNIT CODE", "FONKSİYONEL GRUP", "DİREKTÖRLÜK", "BİRİM", "EKİP", "SON ŞİFRE DEĞİŞİKLİĞİ", "ŞİFRE DEĞİŞİM PERİYODU", "KALAN GÜN SAYISI"]
    for item in column_titles:
        worksheet.write(row, col, item)
        col += 1 
    row = 1
    for i in range(len(account_entries)):
        worksheet.write(row, 0, str(account_entries[i].uid))
        if len(str(account_entries[i].manager)) > 2:
            conn2 = ldapSearch(ldap_environment, str(account_entries[i].manager), user_filter, user_attributes)
            if conn2.result['result'] == 0:
                user_entry = conn2.entries[0]
                worksheet.write(row, 1, str(user_entry.uid))
                worksheet.write(row, 2, str(user_entry.cn))
                worksheet.write(row, 3, str(user_entry.unitcode))
                worksheet.write(row, 4, str(user_entry.functionalgroupname))
                worksheet.write(row, 5, str(user_entry.divisionname))
                worksheet.write(row, 6, str(user_entry.unitname))
                worksheet.write(row, 7, str(user_entry.positionname))
        time_format = '%Y-%m-%d'
        str_now = str(datetime.now()).split(" ")[0]
        str_pwd_change_date = str(account_entries[i].pwdChangedTime).split(" ")[0]
        current_date = datetime.strptime(str_now, time_format)
        str_last_modify_date = str(account_entries[i].modifyTimeStamp).split(" ")[0]
        if len(str_pwd_change_date) == 10:
            worksheet.write(row, 8, str_pwd_change_date)
            pwd_change_date = datetime.strptime(str_pwd_change_date, time_format)
            one_year_later = pwd_change_date.replace(year=pwd_change_date.year + 1)
            worksheet.write(row, 9, str(one_year_later).split(" ")[0])
            time_diff =  one_year_later - current_date
            days_left = str(time_diff).split(",")[0]
            worksheet.write(row, 10, int(days_left.split(" ")[0]), integer_format)
        else:
            worksheet.write(row, 8, str_last_modify_date)
            last_modify_date = datetime.strptime(str_last_modify_date, time_format)
            one_year_later = last_modify_date.replace(year=last_modify_date.year + 1)
            worksheet.write(row, 9, str(one_year_later).split(" ")[0])
            time_diff = one_year_later - current_date
            days_left = str(time_diff).split(",")[0]
            
            worksheet.write(row, 10, int(days_left.split(" ")[0]), integer_format)
        row += 1
    workbook.close()

#Oluşturulan excel dosyasıynı mail ile gönderen fonksiyon

def sendMail(mail_config):
    msg = MIMEMultipart()
    msg['From'] = config.get(mail_config, "sender")
    msg['To'] = config.get(mail_config, "reciever")
    msg['Subject'] = config.get(mail_config, "subject")
    msg['Cc'] = config.get(mail_config, "cc")

    body = "Merhaba,\n\nEkte servis hesaplari raporlari bulunmaktadir."
    msg.attach(MIMEText(body, 'plain'))

    with open(config.get(mail_config, "path"), 'rb') as file:
        attach = MIMEApplication(file.read(),_subtype='xlsx')
        attach.add_header('Content-Disposition', 'attachment', filename='Servis Hesaplari Raporu.xlsx')
        msg.attach(attach)

    smtp_server = smtplib.SMTP(config.get(mail_config, "server"), int(config.get(mail_config, "port")))
    smtp_server.sendmail(config.get(mail_config, "sender"), config.get(mail_config, "reciever"), msg.as_string())
    smtp_server.quit()
