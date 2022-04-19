import pyodbc
import matplotlib.pyplot as plt
from datetime import datetime

templist = []
temp2final = []
final = []
final_giris = []
final_cikis = []
final_tum = []
takoz = [" - "," ÇIKIŞ YOK"]
takoz2 = [" - "," GİRİŞ YOK"]

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\kaan\Desktop\ATT2000.mdb;')
cursor = conn.cursor()
baslangic = input("Başlangıç Tarihi Giriniz: ")+ " 00:00:00"
bitis = input("Bitiş Tarihi Giriniz: ")+" 23:59:59"

baslangic = datetime.strptime(baslangic, '%d/%m/%Y %H:%M:%S')
bitis = datetime.strptime(bitis, '%d/%m/%Y %H:%M:%S')

for row in cursor.execute("select USERID,CHECKTIME,CHECKTYPE from CHECKINOUT"): # burada pydobc kütüphanesi ile database içerisinden satır satır verileri alıyoruz #
    if row.CHECKTIME <= bitis and row.CHECKTIME >= baslangic: # çekilecek verilerin istenen tarih aralığında olup olmadığını kontrol ediyoruz #

        templist.append(row.USERID)

        templist.append(row.CHECKTIME)
        if row.CHECKTYPE == "I":
            templist.append("Giriş")
        else:
            templist.append("Çıkış")
        temp2final.append(templist)
        templist = []

#print(temp2final[0][0])

for i in range(len(temp2final)): # burada id lere göre giriş ve çıkışları sıralıyoruz #
    final.append(temp2final[i])

    for j in range(len(temp2final)):
        if temp2final[i][0] == temp2final[j][0] and temp2final[i][2] == "Giriş" and temp2final[j][2] == "Çıkış":
            final.append(temp2final[j])
            break



for i in range(len(final)):
    if final[i][2] == "Giriş":
        final_giris.append(final[i])


for i in range(len(final)):
    if final[i][2] == "Çıkış":
        final_cikis.append(final[i])

for i in range(len(final_giris)):
    final_tum.append(final_giris[i])
    bos_kontrol = False
    for j in range(len(final_cikis)):
        if final_giris[i][0] == final_cikis[j][0] and len(final_giris[i]) == len(final_cikis[j]):
            final_tum[i].extend(final_cikis[j][1:3])
            bos_kontrol = True
            break
    if(bos_kontrol == False):
        final_tum[i].extend(takoz)

for row in range(len(final_tum)):
    if final[row][2] == "Çıkış" and final[row][0] != final_tum[row][0]:
        final_tum.insert(row,[final[row][0],takoz2[0],takoz2[1],final[row][1],final[row][2]])


for row in cursor.execute("select USERID,Name from USERINFO"): # burada pydobc kütüphanesi ile database içerisinden satır satır kişi isimlerini ID verileri ile alıyoruz #
    for i in range(len(final_tum)):
        if row.USERID == final_tum[i][0]:
            final_tum[i].insert(1,row.Name)



print(final_tum)

fig, ax = plt.subplots(1,1,figsize=(10, 10))
ax.axis("off")
ax.table(cellText=final_tum, colLabels=["ID","İsim","Tarih","İşlem","Tarih","İşlem"], loc='center', cellLoc='center')
plt.show()