import pandas as pd
import xlsxwriter
import datetime


# 1. Katalog va e-stat Excel fayllarini o'qish
katalog_excel = pd.read_excel('data/Katalog.xlsx', index_col=False)
df = pd.read_excel('data/basa.xlsx', index_col=False)

# 2. SOATO-4 va SOATO-7 bo'yicha yozuvlar sonini hisoblash (count)
count_soato_4 = katalog_excel.groupby('SOATO-4').size().reset_index(name='count')
count_soato_7 = katalog_excel.groupby('SOATO-7').size().reset_index(name='count')

# 3. Viloyatlar DataFrame
cities = pd.DataFrame({
    'city_id': [1735, 1703, 1706, 1708, 1710, 1712, 1714, 1718,
                1722, 1724, 1727, 1730, 1733, 1726],
    'city_name': ['Qoraqalpog‘iston Respublikasi', 'Andijon', 'Buxoro', 'Jizzax',
                  'Qashqadaryo', 'Navoiy', 'Namangan', 'Samarqand',
                  'Surxondaryo', 'Sirdaryo', 'Toshkent', 'Farg‘ona',
                  'Xorazm', 'Toshkent shahri']
})

# 4. Tumanlar DataFrame (sizning to‘liq ro‘yxatingiz bilan)
tumanlar = pd.DataFrame({
    'tuman_id': [1735401 ,1735204, 1735207, 1735209, 1735211, 1735212, 1735215, 1735218, 1735222, 1735225, 1735228,
                 1735230, 1735233, 1735236, 1735240, 1735243, 1735250, 1703401, 1703408, 1703202, 1703203, 1703206,
                 1703209, 1703210, 1703211, 1703214, 1703217, 1703220, 1703224, 1703227, 1703230, 1703232, 1703236,
                 1706401, 1706403, 1706204, 1706207, 1706212, 1706215, 1706219, 1706230, 1706232, 1706240, 1706242,
                 1706246, 1706258, 1708401, 1708201, 1708204, 1708209, 1708212, 1708215, 1708218, 1708220, 1708223,
                 1708225, 1708228, 1708235, 1708237, 1710401, 1710405, 1710207, 1710212, 1710220, 1710224, 1710229,
                 1710232, 1710233, 1710234, 1710235, 1710237, 1710240, 1710242, 1710245, 1710250, 1712401, 1712408,
                 1712412, 1712211, 1712216, 1712230, 1712234, 1712238, 1712244, 1712248, 1712251, 1714401, 1714204,
                 1714207, 1714212, 1714216, 1714219, 1714224, 1714229, 1714234, 1714236, 1714237, 1714242, 1718401,
                 1718406, 1718203, 1718206, 1718209, 1718212, 1718215, 1718216, 1718218, 1718224, 1718227, 1718230,
                 1718233, 1718235, 1718236, 1718238, 1722401, 1722201, 1722202, 1722203, 1722204, 1722207, 1722210,
                 1722212, 1722214, 1722215, 1722217, 1722220, 1722221, 1722223, 1722226, 1724401, 1724410, 1724413,
                 1724206, 1724212, 1724216, 1724220, 1724226, 1724228, 1724231, 1724235, 1727401, 1727404, 1727407,
                 1727413, 1727415, 1727419, 1727424, 1727206, 1727212, 1727220, 1727224, 1727228, 1727233, 1727237,
                 1727239, 1727248, 1727249, 1727250, 1727253, 1727256, 1727259, 1727265, 1730401, 1730405, 1730408,
                 1730412, 1730203, 1730206, 1730209, 1730212, 1730215, 1730218, 1730221, 1730224, 1730226, 1730227,
                 1730230, 1730233, 1730236, 1730238, 1730242, 1733401, 1733406, 1733204, 1733208, 1733212, 1733217,
                 1733220, 1733221, 1733223, 1733226, 1733230, 1733233, 1733236, 1726262, 1726264, 1726266, 1726269,
                 1726273, 1726277, 1726280, 1726283, 1726287, 1726290, 1726292, 1726294],
    'tuman_name': ['Nukus shahri', 'Amudaryo tumani', 'Beruniy tumani', 'Bo‘zatov tumani', 'Qorao‘zak tumani',
                   'Kegeyli tumani', 'Qo‘ng‘irot tumani', 'Qanliko‘l tumani', 'Mo‘ynoq tumani', 'Nukus tumani',
                   'Taxiatash tumani', 'Taxtako‘pir tumani', 'To‘rtko‘l tumani', 'Xo‘jayli tumani', 'Chimboy tumani',
                   'Shumanay tumani', 'Ellikqal‘a tumani', 'Andijon shahri', 'Xonobod shahri', 'Oltinko‘l tumani',
                   'Andijon tumani', 'Baliqchi tumani', 'Bo‘ston tumani', 'Buloqboshi tumani', 'Jalaquduq tumani',
                   'Izboskan tumani', 'Ulug‘nor tumani', 'Qo‘rg‘ontepa tumani', 'Asaka tumani', 'Marhamat tumani',
                   'Shahrixon tumani', 'Paxtaobod tumani', 'Xo‘jaobod tumani', 'Buxoro shahri', 'Kogon shahri',
                   'Olot tumani', 'Buxoro tumani', 'Vobkent tumani', 'G‘ijduvon tumani', 'Kogon tumani',
                   'Qorako‘l tumani', 'Qorovulbozor tumani', 'Peshku tumani', 'Romitan tumani', 'Jondor tumani',
                   'Shofirkon tumani', 'Jizzax shahri', 'Arnasoy tumani', 'Baxmal tumani', 'G‘allaorol tumani',
                   'Sh.Rashidov tumani', 'Do‘stlik tumani', 'Zomin tumani', 'Zarbdor tumani', 'Mirzacho‘l tumani',
                   'Zafarobod tumani', 'Paxtakor tumani', 'Forish tumani', 'Yangiobod tumani', 'Qarshi shahri',
                   'Shahrisabz shahri', 'G‘uzor tumani', 'Dehqonobod tumani', 'Qamashi tumani', 'Qarshi tumani',
                   'Koson tumani', 'Kitob tumani', 'Mirishkor tumani', 'Muborak tumani', 'Nishon tumani',
                   'Kasbi tumani', 'Ko‘kdala tumani', 'Chiroqchi tumani', 'Shahrisabz tumani', 'Yakkabog‘ tumani',
                   'Navoiy shahri', 'Zarafshon shahri', 'G‘ozg‘on shahri', 'Konimex tumani', 'Qiziltepa tumani',
                   'Navbahor tumani', 'Karmana tumani', 'Nurota tumani', 'Tomdi tumani', 'Uchquduq tumani',
                   'Xatirchi tumani', 'Namangan shahri', 'Mingbuloq tumani', 'Kosonsoy tumani', 'Namangan tumani',
                   'Norin tumani', 'Pop tumani', 'To‘raqo‘rg‘on tumani', 'Uychi tumani', 'Uchqo‘rg‘on tumani',
                   'Chortoq tumani', 'Chust tumani', 'Yangiqo‘rg‘on tumani', 'Samarqand shahri',
                   'Kattaqo‘rg‘on shahri', 'Oqdaryo tumani', 'Bulung‘ur tumani', 'Jomboy tumani', 'Ishtixon tumani',
                   'Kattaqo‘rg‘on tumani', 'Qo‘shrabot tumani', 'Narpay tumani', 'Payariq tumani', 'Pastdarg‘om tumani',
                   'Paxtachi tumani', 'Samarqand tumani', 'Nurobod tumani', 'Urgut tumani', 'Toyloq tumani',
                   'Termiz shahri', 'Oltinsoy tumani', 'Angor tumani', 'Bandixon tumani', 'Boysun tumani',
                   'Muzrabot tumani', 'Denov tumani', 'Jarqo‘rg‘on tumani', 'Qumqo‘rg‘on tumani', 'Qiziriq tumani',
                   'Sariosiyo tumani', 'Termiz tumani', 'Uzun tumani', 'Sherobod tumani', 'Sho‘rchi tumani',
                   'Guliston shahri', 'Shirin shahri', 'Yangier shahri', 'Oqoltin tumani', 'Boyovut tumani',
                   'Sayxunobod tumani', 'Guliston tumani', 'Sardoba tumani', 'Mirzaobod tumani', 'Sirdaryo tumani',
                   'Xovos tumani', 'Nurafshon shahri', 'Olmaliq shahri', 'Angren shahri', 'Bekobod shahri',
                   'Ohangaron shahri', 'Chirchiq shahri', 'Yangiyo‘l shahri', 'Oqqo‘rg‘on tumani', 'Ohangaron tumani',
                   'Bekobod tumani', 'Bo‘stonliq tumani', 'Bo‘ka tumani', 'Quyi Chirchiq tumani', 'Zangiota tumani',
                   'Yuqori Chirchiq tumani', 'Qibray tumani', 'Parkent tumani', 'Piskent tumani', 'O‘rta Chirchiq tumani',
                   'Chinoz tumani', 'Yangiyo‘l tumani', 'Toshkent tumani', 'Farg‘ona shahri', 'Qo‘qon shahri',
                   'Quvasoy shahri', 'Marg‘ilon shahri', 'Oltiariq tumani', 'Qo‘shtepa tumani', 'Bag‘dod tumani',
                   'Buvayda tumani', 'Beshariq tumani', 'Quva tumani', 'Uchkо‘prik tumani', 'Rishton tumani',
                   'So‘x tumani', 'Toshloq tumani', 'O‘zbekiston tumani', 'Farg‘ona tumani', 'Dang‘ara tumani',
                   'Furqat tumani', 'Yozyovon tumani', 'Urganch shahri', 'Xiva shahri', 'Bog‘ot tumani',
                   'Gurlan tumani', 'Qo‘shko‘pir tumani', 'Urganch tumani', 'Hazorasp tumani', 'To‘roqqal‘a tumani',
                   'Xonqa tumani', 'Xiva tumani', 'Shovot tumani', 'Yangiariq tumani', 'Yangibozor tumani',
                   'Uchtepa tumani', 'Bektemir tumani', 'Yunusobod tumani', 'Mirzo Ulug‘bek tumani', 'Mirobod tumani',
                   'Shayxontohur tumani', 'Olmazor tumani', 'Sergeli tumani', 'Yakkasaroy tumani', 'Yashnobod tumani',
                   'YangihaYot tumani', 'Chilonzor tumani']
})

# 5. Count larni viloyatlar va tumanlar bilan bog‘lash
cities_with_count = pd.merge(cities, count_soato_4, how='left', left_on='city_id', right_on='SOATO-4')
tumanlar_with_count = pd.merge(tumanlar, count_soato_7, how='left', left_on='tuman_id', right_on='SOATO-7')

# 6. NaN ni 0 ga almashtirish va int ga o'tkazish
cities_with_count['count'] = cities_with_count['count'].fillna(0).astype(int)
tumanlar_with_count['count'] = tumanlar_with_count['count'].fillna(0).astype(int)

# 7. Ustun nomlarini o'zgartirish (ID='1', nom='2', count='3')
cities_with_count.rename(columns={'city_id': '1', 'city_name': '2', 'count': '3'}, inplace=True)
tumanlar_with_count.rename(columns={'tuman_id': '1', 'tuman_name': '2', 'count': '3'}, inplace=True)

# 8. Bitta DataFrame ga birlashtirish va 'SOATO-4' deb nomlash
regions = pd.concat([cities_with_count[['1', '2', '3']], tumanlar_with_count[['1', '2', '3']]], ignore_index=True)
regions.rename(columns={'1': 'SOATO-4'}, inplace=True)

# 9. e-stat va katalogni birlashtirish
merged = pd.merge(df, katalog_excel, left_on='KTUT', right_on='OKPO', how='right')

# 10. Qabul qilinganlarni hisoblash
merged['Qabul_qilingan_soni'] = (merged['Joriy holati'] == 'Qabul qilingan').astype(int)
qabul_4 = merged.groupby('SOATO-4', as_index=False)['Qabul_qilingan_soni'].sum()
qabul_7 = merged.groupby('SOATO-7', as_index=False)['Qabul_qilingan_soni'].sum()
qabul_7.rename(columns={'SOATO-7': 'SOATO-4'}, inplace=True)
qabul_all = pd.concat([qabul_4, qabul_7], ignore_index=True)

# 11. Rad etilganlarni hisoblash
merged['Rad etilgan'] = (merged['Joriy holati'] == 'Rad etilgan').astype(int)
rad_4 = merged.groupby('SOATO-4', as_index=False)['Rad etilgan'].sum()
rad_7 = merged.groupby('SOATO-7', as_index=False)['Rad etilgan'].sum()
rad_7.rename(columns={'SOATO-7': 'SOATO-4'}, inplace=True)
rad_all = pd.concat([rad_4, rad_7], ignore_index=True)

# 12. Ko‘rib chiqish jarayonidagilarni hisoblash
merged["Ko'rib chiqish jarayonida"] = (
    (merged['Joriy holati'] == "Ko'rib chiqish jarayonida") |
    (merged['Joriy holati'] == "Jo'natildi")
).astype(int)
korib_4 = merged.groupby('SOATO-4', as_index=False)["Ko'rib chiqish jarayonida"].sum()
korib_7 = merged.groupby('SOATO-7', as_index=False)["Ko'rib chiqish jarayonida"].sum()
korib_7.rename(columns={'SOATO-7': 'SOATO-4'}, inplace=True)
korib_all = pd.concat([korib_4, korib_7], ignore_index=True)

# 13. Statistikalarni regions bilan birlashtirish
final = pd.merge(regions, qabul_all, on='SOATO-4', how='left')
final = pd.merge(final, rad_all, on='SOATO-4', how='left')
final = pd.merge(final, korib_all, on='SOATO-4', how='left')

# 14. NaN ni 0 ga almashtirish va int ga o'tkazish
final[['Qabul_qilingan_soni', 'Rad etilgan', "Ko'rib chiqish jarayonida"]] = \
    final[['Qabul_qilingan_soni', 'Rad etilgan', "Ko'rib chiqish jarayonida"]].fillna(0).astype(int)

# 15. Hisobot taqdim etganlar va taqdim etmaganlarni hisoblash
final['Hisobot taqdim etganlar'] = final['Qabul_qilingan_soni'] + final['Rad etilgan'] + final["Ko'rib chiqish jarayonida"]
final['Hisobot taqdim etmaganlar'] = final['3'] - final['Hisobot taqdim etganlar']

# 16. Ulushlar (foizlarda, 1 xonali)
final['Hisobot taqdim etganlar ulushi %'] = (final['Hisobot taqdim etganlar'] * 100.0 / final['3']).round(1)
final['Qabul qilinganlar ulushi %'] = (final['Qabul_qilingan_soni'] * 100.0 / final['3']).round(1)
final['Ko‘rib chiqilayotganlar ulushi %'] = (final["Ko'rib chiqish jarayonida"] * 100.0 / final['3']).round(1)
final['Hisobot taqdim etmaganlar ulushi %'] = (final["Hisobot taqdim etmaganlar"] * 100.0 / final['3']).round(1)

# 17. Ustun tartibi
final = final[['SOATO-4', '2', '3',
               'Hisobot taqdim etganlar', 'Hisobot taqdim etganlar ulushi %',
               'Qabul_qilingan_soni', 'Qabul qilinganlar ulushi %',
               'Rad etilgan', "Ko'rib chiqish jarayonida",
               'Hisobot taqdim etmaganlar', 'Hisobot taqdim etmaganlar ulushi %']]

# 18. O‘zbekiston Respublikasi uchun umumiy yig‘indi
uzb_sum = cities_with_count['3'].sum()
uzb_qabul_sum = final['Qabul_qilingan_soni'].sum() / 2  # Nima uchun 2 ga bo‘lyapti, asl kodda shunday edi
uzb_rad_sum = final['Rad etilgan'].sum() / 2
uzb_korib_sum = final["Ko'rib chiqish jarayonida"].sum() / 2

uzb_row = pd.DataFrame({
    'SOATO-4': [1700],
    '2': ['O‘zbekiston Respublikasi'],
    '3': [uzb_sum],
    'Hisobot taqdim etganlar': [uzb_qabul_sum + uzb_rad_sum + uzb_korib_sum],
    'Hisobot taqdim etganlar ulushi %': [round((uzb_qabul_sum + uzb_rad_sum + uzb_korib_sum) * 100.0 / uzb_sum, 1)],
    'Qabul_qilingan_soni': [uzb_qabul_sum],
    'Qabul qilinganlar ulushi %': [round(uzb_qabul_sum * 100.0 / uzb_sum, 1)],
    'Rad etilgan': [uzb_rad_sum],
    "Ko'rib chiqish jarayonida": [uzb_korib_sum],
    'Hisobot taqdim etmaganlar': [uzb_sum - (uzb_qabul_sum + uzb_rad_sum + uzb_korib_sum)],
    'Hisobot taqdim etmaganlar ulushi %': [round((uzb_sum - (uzb_qabul_sum + uzb_rad_sum + uzb_korib_sum)) * 100.0 / uzb_sum, 1)]
})

# 19. Yakuniy jadvalga qo‘shish
final_result = pd.concat([uzb_row, final], ignore_index=True)

# --- Faylga yozish: chiroyli header bilan ---
# Create the Excel file with proper headersimport pandas as pd
with pd.ExcelWriter('Corrected Report.xlsx', engine='xlsxwriter') as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet('Sheet1')
    writer.sheets['Sheet1'] = worksheet

    # Define formats
    merge_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1,
        'font_name': 'Times New Roman'
    })

    bold_center = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1,
        'font_name': 'Times New Roman'
    })

    normal_center = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'font_name': 'Times New Roman'
    })

    # Format for red background (when value < E8 reference)
    red_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'font_name': 'Times New Roman',
        'bg_color': '#FF0000',  # Red background
        'font_color': '#FFFFFF'  # White text for better visibility
    })

    # Format for green background (when value = 100)
    green_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'font_name': 'Times New Roman',
        'bg_color': '#00FF00',  # Green background
        'font_color': '#000000'  # Black text for better visibility
    })

    # Row 1-2: Title (merged A1:K2)
    worksheet.merge_range('A1:K2',
                          "2025-yilning yanvar-oktabr oylarida tijorat korxonalari tomonidan taqdim etilgan 12-invest kuzatuvi qamrovini ta'minlanishi to'g'risida.\nMA'LUMOT",
                          merge_format)

    # Row 3-4: First date (merged A3:E4) - 2 qator, 5 ustun
    worksheet.merge_range('A3:E4', "06.11.2025 y.", normal_center)

    # Row 3-4: Second date (merged F3:K4) - 2 qator, 6 ustun
    worksheet.merge_range('F3:K4', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), normal_center)

    # Row 5-6: Main headers - leave first two columns blank
    # Columns A and B are blank (empty)

    # Column C (index 2)
    worksheet.merge_range('C5:C6', 'Respondentlar\nsoni', bold_center)

    # Column D (index 3)
    worksheet.merge_range('D5:D6', 'Hisobot\ntaqdim\netganlar', bold_center)

    # Column E (index 4)
    worksheet.merge_range('E5:E6', 'Ulushi, %', bold_center)

    # Column F (index 5)
    worksheet.merge_range('F5:F6', 'Qabul qilingan', bold_center)

    # Column G (index 6)
    worksheet.merge_range('G5:G6', 'Ulushi, %', bold_center)

    # Column H (index 7)
    worksheet.merge_range('H5:H6', 'Rad etilgan', bold_center)

    # Column I (index 8)
    worksheet.merge_range('I5:I6', "Ko'rib chiqish\njarayonida", bold_center)

    # Column J (index 9)
    worksheet.merge_range('J5:J6', 'Hisobot\ntaqdim\netmaganlar', bold_center)

    # Column K (index 10)
    worksheet.merge_range('K5:K6', 'Ulushi, %', bold_center)

    # Row 7: Column numbers
    row7_numbers = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11']
    for col_num, value in enumerate(row7_numbers):
        worksheet.write(6, col_num, value, bold_center)

    # Write data starting from row 8 (index 7) with center alignment
    start_row = 7

    # Get reference value from E8 (first row of data, column E which is index 4)
    e8_reference_value = None
    if len(final_result) > 0:
        try:
            e8_reference_value = float(final_result.iloc[0, 4])  # First row, column E (index 4)
        except (ValueError, TypeError):
            e8_reference_value = None

    for row_idx, row_data in final_result.iterrows():
        for col_idx, value in enumerate(row_data):
            # Check if we're in column E (index 4) and rows 9-21 (not row 8)
            # E9 corresponds to row_idx=1, E21 corresponds to row_idx=13
            if col_idx == 4 and 1 <= row_idx <= 13 and e8_reference_value is not None:
                # Try to convert value to float and apply formatting
                try:
                    if pd.notna(value):
                        float_value = float(value)
                        # Check if value equals 100 (green background)
                        if float_value == 100:
                            worksheet.write(start_row + row_idx, col_idx, value, green_format)
                        # Check if value is below E8 reference (red background)
                        elif float_value < e8_reference_value:
                            worksheet.write(start_row + row_idx, col_idx, value, red_format)
                        else:
                            worksheet.write(start_row + row_idx, col_idx, value, normal_center)
                    else:
                        worksheet.write(start_row + row_idx, col_idx, value, normal_center)
                except (ValueError, TypeError):
                    worksheet.write(start_row + row_idx, col_idx, value, normal_center)
            else:
                worksheet.write(start_row + row_idx, col_idx, value, normal_center)

    # Set column widths
    worksheet.set_column('A:B', 15)
    worksheet.set_column('C:K', 12)

print("✅ Corrected Report.xlsx fayli yaratildi.")