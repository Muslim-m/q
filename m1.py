import pandas as pd
import datetime

# ========================= YO'LLAR =========================
katalog_path = 'data/12-инвест каталог sheet1.xlsx'
basa_path    = 'data/table_data - 2025-11-28T171733.084.xlsx'

# ========================= KATALOGNI O'QISH =========================
xl = pd.ExcelFile(katalog_path)
sheet = xl.sheet_names[0]
print(f"✔ Katalog fayldan o'qilayotgan varaq: '{sheet}'")
katalog_excel = pd.read_excel(katalog_path, sheet_name=sheet, dtype=str)

df_baza = pd.read_excel(basa_path, dtype=str)

# Tozalash
katalog_excel['okpo'] = katalog_excel['okpo'].astype(str).str.strip()
katalog_excel['SOATO-4'] = katalog_excel['SOATO-4'].astype(str).str.strip()
katalog_excel['SOATO-7'] = katalog_excel['SOATO-7'].astype(str).str.strip()

df_baza['KTUT'] = df_baza['KTUT'].astype(str).str.strip()

# ========================= VILOYAT VA TUMANLAR (hammasi str) =========================
cities = pd.DataFrame({
    'city_id': ['1735','1703','1706','1708','1710','1712','1714','1718','1722','1724','1727','1730','1733','1726'],
    'city_name': ['Qoraqalpog‘iston Respublikasi','Andijon','Buxoro','Jizzax','Qashqadaryo','Navoiy',
                  'Namangan','Samarqand','Surxondaryo','Sirdaryo','Toshkent','Farg‘ona','Xorazm','Toshkent shahri']
})

tumanlar = pd.DataFrame({
    'tuman_id': ['1735401','1735204','1735207','1735209','1735211','1735212','1735215','1735218','1735222','1735225','1735228','1735230','1735233','1735236','1735240','1735243','1735250',
                 '1703401','1703408','1703202','1703203','1703206','1703209','1703210','1703211','1703214','1703217','1703220','1703224','1703227','1703230','1703232','1703236',
                 '1706401','1706403','1706204','1706207','1706212','1706215','1706219','1706230','1706232','1706240','1706242','1706246','1706258',
                 '1708401','1708201','1708204','1708209','1708212','1708215','1708218','1708220','1708223','1708225','1708228','1708235','1708237',
                 '1710401','1710405','1710207','1710212','1710220','1710224','1710229','1710232','1710233','1710234','1710235','1710237','1710240','1710242','1710245','1710250',
                 '1712401','1712408','1712412','1712211','1712216','1712230','1712234','1712238','1712244','1712248','1712251',
                 '1714401','1714204','1714207','1714212','1714216','1714219','1714224','1714229','1714234','1714236','1714237','1714242',
                 '1718401','1718406','1718203','1718206','1718209','1718212','1718215','1718216','1718218','1718224','1718227','1718230','1718233','1718235','1718236','1718238',
                 '1722401','1722201','1722202','1722203','1722204','1722207','1722210','1722212','1722214','1722215','1722217','1722220','1722221','1722223','1722226',
                 '1724401','1724410','1724413','1724206','1724212','1724216','1724220','1724226','1724228','1724231','1724235',
                 '1727401','1727404','1727407','1727413','1727415','1727419','1727424','1727206','1727212','1727220','1727224','1727228','1727233','1727237','1727239','1727248','1727249','1727250','1727253','1727256','1727259','1727265',
                 '1730401','1730405','1730408','1730412','1730203','1730206','1730209','1730212','1730215','1730218','1730221','1730224','1730226','1730227','1730230','1730233','1730236','1730238','1730242',
                 '1733401','1733406','1733204','1733208','1733212','1733217','1733220','1733221','1733223','1733226','1733230','1733233','1733236',
                 '1726262','1726264','1726266','1726269','1726273','1726277','1726280','1726283','1726287','1726290','1726292','1726294'],
    'tuman_name': ['Nukus shahri','Amudaryo tumani','Beruniy tumani','Bo‘zatov tumani','Qorao‘zak tumani','Kegeyli tumani','Qo‘ng‘irot tumani','Qanliko‘l tumani','Mo‘ynoq tumani','Nukus tumani','Taxiatash tumani','Taxtako‘pir tumani','To‘rtko‘l tumani','Xo‘jayli tumani','Chimboy tumani','Shumanay tumani','Ellikqal‘a tumani',
                   'Andijon shahri','Xonobod shahri','Oltinko‘l tumani','Andijon tumani','Baliqchi tumani','Bo‘ston tumani','Buloqboshi tumani','Jalaquduq tumani','Izboskan tumani','Ulug‘nor tumani','Qo‘rg‘ontepa tumani','Asaka tumani','Marhamat tumani','Shahrixon tumani','Paxtaobod tumani','Xo‘jaobod tumani',
                   'Buxoro shahri','Kogon shahri','Olot tumani','Buxoro tumani','Vobkent tumani','G‘ijduvon tumani','Kogon tumani','Qorako‘l tumani','Qorovulbozor tumani','Peshku tumani','Romitan tumani','Jondor tumani','Shofirkon tumani',
                   'Jizzax shahri','Arnasoy tumani','Baxmal tumani','G‘allaorol tumani','Sh.Rashidov tumani','Do‘stlik tumani','Zomin tumani','Zarbdor tumani','Mirzacho‘l tumani','Zafarobod tumani','Paxtakor tumani','Forish tumani','Yangiobod tumani',
                   'Qarshi shahri','Shahrisabz shahri','G‘uzor tumani','Dehqonobod tumani','Qamashi tumani','Qarshi tumani','Koson tumani','Kitob tumani','Mirishkor tumani','Muborak tumani','Nishon tumani','Kasbi tumani','Ko‘kdala tumani','Chiroqchi tumani','Shahrisabz tumani','Yakkabog‘ tumani',
                   'Navoiy shahri','Zarafshon shahri','G‘ozg‘on shahri','Konimex tumani','Qiziltepa tumani','Navbahor tumani','Karmana tumani','Nurota tumani','Tomdi tumani','Uchquduq tumani','Xatirchi tumani',
                   'Namangan shahri','Mingbuloq tumani','Kosonsoy tumani','Namangan tumani','Norin tumani','Pop tumani','To‘raqo‘rg‘on tumani','Uychi tumani','Uchqo‘rg‘on tumani','Chortoq tumani','Chust tumani','Yangiqo‘rg‘on tumani',
                   'Samarqand shahri','Kattaqo‘rg‘on shahri','Oqdaryo tumani','Bulung‘ur tumani','Jomboy tumani','Ishtixon tumani','Kattaqo‘rg‘on tumani','Qo‘shrabot tumani','Narpay tumani','Payariq tumani','Pastdarg‘om tumani','Paxtachi tumani','Samarqand tumani','Nurobod tumani','Urgut tumani','Toyloq tumani',
                   'Termiz shahri','Oltinsoy tumani','Angor tumani','Bandixon tumani','Boysun tumani','Muzrabot tumani','Denov tumani','Jarqo‘rg‘on tumani','Qumqo‘rg‘on tumani','Qiziriq tumani','Sariosiyo tumani','Termiz tumani','Uzun tumani','Sherobod tumani','Sho‘rchi tumani',
                   'Guliston shahri','Shirin shahri','Yangier shahri','Oqoltin tumani','Boyovut tumani','Sayxunobod tumani','Guliston tumani','Sardoba tumani','Mirzaobod tumani','Sirdaryo tumani','Xovos tumani',
                   'Nurafshon shahri','Olmaliq shahri','Angren shahri','Bekobod shahri','Ohangaron shahri','Chirchiq shahri','Yangiyo‘l shahri','Oqqo‘rg‘on tumani','Ohangaron tumani','Bekobod tumani','Bo‘stonliq tumani','Bo‘ka tumani','Quyi Chirchiq tumani','Zangiota tumani','Yuqori Chirchiq tumani','Qibray tumani','Parkent tumani','Piskent tumani','O‘rta Chirchiq tumani','Chinoz tumani','Yangiyo‘l tumani','Toshkent tumani',
                   'Farg‘ona shahri','Qo‘qon shahri','Quvasoy shahri','Marg‘ilon shahri','Oltiariq tumani','Qo‘shtepa tumani','Bag‘dod tumani','Buvayda tumani','Beshariq tumani','Quva tumani','Uchkо‘prik tumani','Rishton tumani','So‘x tumani','Toshloq tumani','O‘zbekiston tumani','Farg‘ona tumani','Dang‘ara tumani','Furqat tumani','Yozyovon tumani',
                   'Urganch shahri','Xiva shahri','Bog‘ot tumani','Gurlan tumani','Qo‘shko‘pir tumani','Urganch tumani','Hazorasp tumani','To‘roqqal‘a tumani','Xonqa tumani','Xiva tumani','Shovot tumani','Yangiariq tumani','Yangibozor tumani',
                   'Uchtepa tumani','Bektemir tumani','Yunusobod tumani','Mirzo Ulug‘bek tumani','Mirobod tumani','Shayxontohur tumani','Olmazor tumani','Sergeli tumani','Yakkasaroy tumani','Yashnobod tumani','YangihaYot tumani','Chilonzor tumani']
})

# ========================= RESPONDENTLAR SONI =========================
count_4 = katalog_excel.groupby('SOATO-4').size().reset_index(name='count')
count_7 = katalog_excel.groupby('SOATO-7').size().reset_index(name='count')

cities_cnt = pd.merge(cities, count_4, how='left', left_on='city_id', right_on='SOATO-4', suffixes=('', '_4'))
tuman_cnt  = pd.merge(tumanlar, count_7, how='left', left_on='tuman_id', right_on='SOATO-7', suffixes=('', '_7'))

cities_cnt['count'] = cities_cnt['count'].fillna(0).astype(int)
tuman_cnt['count']  = tuman_cnt['count'].fillna(0).astype(int)

regions = pd.concat([
    cities_cnt[['city_id', 'city_name', 'count']].rename(columns={'city_id':'SOATO', 'city_name':'Name', 'count':'Total'}),
    tuman_cnt[['tuman_id', 'tuman_name', 'count']].rename(columns={'tuman_id':'SOATO', 'tuman_name':'Name', 'count':'Total'})
], ignore_index=True)

# ========================= MERGE =========================
merged = pd.merge(df_baza, katalog_excel, left_on='KTUT', right_on='okpo', how='right')

merged['Qabul'] = (merged['Joriy holati'] == 'Qabul qilingan').astype(int)
merged['Rad']   = (merged['Joriy holati'] == 'Rad etilgan').astype(int)
merged['Korib'] = ((merged['Joriy holati'] == "Ko'rib chiqish jarayonida") | (merged['Joriy holati'] == "Jo'natildi")).astype(int)

def sum_by_both(col):
    s4 = merged.groupby('SOATO-4')[col].sum()
    s7 = merged.groupby('SOATO-7')[col].sum()
    return pd.concat([s4, s7]).groupby(level=0).sum()

qabul_all = sum_by_both('Qabul').reset_index().rename(columns={'Qabul':'Qabul_qilingan'})
rad_all   = sum_by_both('Rad').reset_index().rename(columns={'Rad':'Rad_etilgan'})
korib_all = sum_by_both('Korib').reset_index().rename(columns={'Korib':"Ko'rib_chiqish_jarayonida"})

# ========================= FINAL =========================
final = regions.merge(qabul_all, left_on='SOATO', right_on='SOATO-4', how='left').fillna(0)
final = final.merge(rad_all,   left_on='SOATO', right_on='SOATO-4', how='left').fillna(0)
final = final.merge(korib_all, left_on='SOATO', right_on='SOATO-4', how='left').fillna(0)

final['Qabul_qilingan'] = final['Qabul_qilingan'].astype(int)
final['Rad_etilgan']    = final['Rad_etilgan'].astype(int)
final["Ko'rib_chiqish_jarayonida"] = final["Ko'rib_chiqish_jarayonida"].astype(int)

final['Taqdim_etganlar'] = final['Qabul_qilingan'] + final['Rad_etilgan'] + final["Ko'rib_chiqish_jarayonida"]
final['Taqdim_etmaganlar'] = final['Total'] - final['Taqdim_etganlar']

final['Taqdim_etganlar_%'] = (final['Taqdim_etganlar'] * 100 / final['Total'].replace(0,1)).round(1)
final['Qabul_%']           = (final['Qabul_qilingan'] * 100 / final['Total'].replace(0,1)).round(1)
final['Korib_%']           = (final["Ko'rib_chiqish_jarayonida"] * 100 / final['Total'].replace(0,1)).round(1)
final['Taqdim_etmagan_%']  = (final['Taqdim_etmaganlar'] * 100 / final['Total'].replace(0,1)).round(1)

final = final[['SOATO', 'Name', 'Total', 'Taqdim_etganlar', 'Taqdim_etganlar_%',
               'Qabul_qilingan', 'Qabul_%', 'Rad_etilgan', "Ko'rib_chiqish_jarayonida",
               'Taqdim_etmaganlar', 'Taqdim_etmagan_%']]

# O‘zbekiston jami
uzb = pd.DataFrame([{
    'SOATO': '1700',
    'Name': 'O‘zbekiston Respublikasi',
    'Total': final['Total'].sum(),
    'Taqdim_etganlar': final['Taqdim_etganlar'].sum(),
    'Qabul_qilingan': final['Qabul_qilingan'].sum(),
    'Rad_etilgan': final['Rad_etilgan'].sum(),
    "Ko'rib_chiqish_jarayonida": final["Ko'rib_chiqish_jarayonida"].sum(),
    'Taqdim_etmaganlar': final['Taqdim_etmaganlar'].sum(),
}])

uzb['Taqdim_etganlar_%'] = round(uzb['Taqdim_etganlar'] * 100 / uzb['Total'], 1) if uzb['Total'].iloc[0] > 0 else 0
uzb['Qabul_%']           = round(uzb['Qabul_qilingan'] * 100 / uzb['Total'], 1) if uzb['Total'].iloc[0] > 0 else 0
uzb['Korib_%']           = round(uzb["Ko'rib_chiqish_jarayonida"] * 100 / uzb['Total'], 1) if uzb['Total'].iloc[0] > 0 else 0
uzb['Taqdim_etmagan_%']  = round(uzb['Taqdim_etmaganlar'] * 100 / uzb['Total'], 1) if uzb['Total'].iloc[0] > 0 else 0

final_result = pd.concat([uzb, final], ignore_index=True)

# ========================= EXCEL =========================
with pd.ExcelWriter('Corrected Report.xlsx', engine='xlsxwriter') as writer:
    final_result.to_excel(writer, sheet_name='Sheet1', startrow=7, header=False, index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    title_fmt = workbook.add_format({'bold':True, 'font_size':14, 'align':'center', 'valign':'vcenter', 'text_wrap':True, 'font_name':'Times New Roman'})
    head_fmt  = workbook.add_format({'bold':True, 'align':'center', 'valign':'vcenter', 'text_wrap':True, 'border':1, 'font_name':'Times New Roman'})
    cell_fmt  = workbook.add_format({'align':'center', 'valign':'vcenter', 'border':1, 'font_name':'Times New Roman'})
    red_fmt   = workbook.add_format({'bg_color':'#FF0000', 'font_color':'#FFFFFF', 'align':'center', 'border':1})
    green_fmt = workbook.add_format({'bg_color':'#00FF00', 'align':'center', 'border':1})

    worksheet.merge_range('A1:K2', "2025-yilning yanvar-oktabr oylarida tijorat korxonalari tomonidan taqdim etilgan 12-invest kuzatuvi qamrovini ta'minlanishi to'g'risida.\nMA'LUMOT", title_fmt)
    worksheet.merge_range('A3:E4', "06.11.2025 y.", cell_fmt)
    worksheet.merge_range('F3:K4', datetime.datetime.now().strftime('%d.%m.%Y %H:%M'), cell_fmt)

    headers = ['№','Hudud nomi','Respondentlar soni','Hisobot taqdim etganlar','Ulushi, %','Qabul qilingan','Ulushi, %','Rad etilgan',"Ko'rib chiqish jarayonida",'Hisobot taqdim etmaganlar','Ulushi, %']
    for c, h in enumerate(headers):
        if c in [2,3,5,8,9]:
            worksheet.merge_range(4, c, 5, c, h, head_fmt)
        else:
            worksheet.write(4, c, '', head_fmt)
        worksheet.write(5, c, '', head_fmt)
    for c, n in enumerate(['1','2','3','4','5','6','7','8','9','10','11']):
        worksheet.write(6, c, n, head_fmt)

    uzb_percent = final_result.iloc[0, 4]
    for r in range(len(final_result)):
        row = 7 + r
        for c in range(11):
            val = final_result.iat[r, c]
            if c == 4 and r > 0 and r <= 14:
                try:
                    v = float(val)
                    if v == 100:
                        worksheet.write(row, c, val, green_fmt)
                    elif v < float(uzb_percent):
                        worksheet.write(row, c, val, red_fmt)
                    else:
                        worksheet.write(row, c, val, cell_fmt)
                except:
                    worksheet.write(row, c, val, cell_fmt)
            else:
                worksheet.write(row, c, val, cell_fmt)

    worksheet.set_column('A:A', 8)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:K', 14)

print("Corrected Report.xlsx muvaffaqiyatli yaratildi!")