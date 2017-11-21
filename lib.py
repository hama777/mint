import requests
import codecs
from bs4 import BeautifulSoup

login_url = "https://www.lib.city.kobe.jp/opac/opacs/login?user[login]=%s" +\
         "&user[passwd]=%s"+\
         "&act_login=%E3%83%AD%E3%82%B0%E3%82%A4%E3%83%B3&nextAction=mypage_display&prevAction=find_request"

f = open("user.txt","r")
user = readline(f)
passwd = readline(f)

ckurl = login_url % (user,passwd)
ses = requests.session()
#r = requests.get(ckurl)
r = ses.post(ckurl)

f = codecs.open("www.htm","w","utf-8")
s = r.text
f.write(s)
f.close()

ckurl = "https://www.lib.city.kobe.jp/opac/opacs/reservation_display"
r = ses.get(ckurl)

f = codecs.open("book.htm","w","utf-8")
s = r.text
f.write(s)
f.close()
print("end")

def AnalizeReserveList() :
    topurl = "https://www.lib.city.kobe.jp/opac/opacs/reservation_cancel_confirmation?reservation_order_confirmation=%e9%a0%86%e4%bd%8d%e7%a2%ba%e8%aa%8d"

    html = open("book.htm","r").read()
    sp = BeautifulSoup(html)
    table = sp.findall("table")
    for row in rows:
        cell =row.findAll("td")[2]   # 予約番号は3カラム目
        print(cell)
