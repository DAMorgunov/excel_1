# создаю фал excel
import openpyxl

book = openpyxl.open('isk_rodina.xlsx')
sheet1 = book['Счет']
sheet2 = book['СЧФ']
sheet3 = book['Акт']

doc_num = input('Введите номер счёта: ')
doc_data = input('Введите дату в формате "ЧИСЛО месяц": ')
doc_perinod = input('Введите период в формате "месяц": ')


def summa(month):
    a = '67 270,00'
    b = '65 100,00'
    c = '60 760,00'
    days31 = ['январь', 'март', 'май', 'июль', 'август', 'октябрь', 'декабрь']
    days30 = ['апрель', 'июнь', 'сентябрь', 'ноябрь']

    if month in days31:
        return a

    if month in days30:
        return b

    else:
        return c


def summa_text(n):
    a = 'Шестьдесят семь тысяч двести семьдесят'
    b = 'Шестьдесят пять тысяч сто'
    c = 'Шестьдесят тысяч семьсот шестьдесят'

    if summa(doc_perinod) == '67 270,00':
        return a

    if summa(doc_perinod) == '65 100,00':
        return b

    else:
        return c


def nds(n):
    nds_67 = '11 211,67'
    nds_65 = '10 850,00'
    nds_60 = '10 126,67'

    if summa(doc_perinod) == '67 270,00':
        return nds_67

    if summa(doc_perinod) == '65 100,00':
        return nds_65

    else:
        return nds_60


nds_rub = nds(summa(doc_perinod))[:-3]
nds_kop = nds(summa(doc_perinod))[-2:]

num_doc = sheet1['A9'].value  # считываю строку для указания номера документа
num_doc = num_doc.replace('number', doc_num)  # указываю номер документа
sheet1['A9'] = num_doc

data_doc = sheet1['A9'].value  # считываю строку для указания даты документа
data_doc = data_doc.replace('data', doc_data)  # указываю дату документа
sheet1['A9'] = data_doc

period_doc = sheet1['B16'].value  # считываю строку для указания периода оплаты
period_doc = period_doc.replace('period', doc_perinod.capitalize())  # указываю период оплаты
sheet1['B16'] = period_doc

sheet1['L16'] = sheet1['M16'] = sheet1['M17'] = sheet2['Z17'] = sheet2['Z18'] = sheet3['G8'] = sheet3['H8'] = sheet3[
    'H10'] = summa(doc_perinod)  # указываю сумму на листах Счета СЧФ и Акты

sheet1['M18'] = sheet2['W17'] = sheet2['W18'] = sheet3['H11'] = nds(summa(doc_perinod))  # указываю ндс на листах Счета СЧФ и Акты

itog1_doc = sheet1['A19'].value  # считываю строку для указания итоговой суммы оплаты
itog1_doc: object = itog1_doc.replace('summa', summa(doc_perinod))  # указываю итоговую сумму
sheet1['A19'] = itog1_doc

itog2_doc = sheet1['A20'].value  # считываю строку для указания итоговой суммы текстом
itog2_doc = itog2_doc.replace('summa_text', summa_text(summa(doc_perinod)))
sheet1['A20'] = itog2_doc

nds_text = sheet1['A20'].value  # считываю строку для указания ндс рублей текстом
nds_text = nds_text.replace('nds_rub', nds_rub)
sheet1['A20'] = nds_text

nds_text2 = sheet1['A20'].value  # считываю строку для указания ндс копеек текстом
nds_text2 = nds_text2.replace('nds_kop', nds_kop)
sheet1['A20'] = nds_text2

# работа с листом СЧФ

num_shf = input('Введите номер счёта-фактуры: ')  # запрос номера СЧФ
sheet2['C1'] = num_shf

data_doc = sheet2['K1'].value  # считываю строку для указания даты документа
data_doc = data_doc.replace('data', doc_data)  # указываю дату документа
sheet2['K1'] = data_doc


def summa_bez_nds(n):
    a = '56 038,33'
    b = '54 250,00'
    c = '50 633,33'

    if summa(doc_perinod) == '67 270,00':
        return a

    if summa(doc_perinod) == '65 100,00':
        return b

    else:
        return c


sheet2['M17'] = sheet2['O17'] = sheet2['O18'] = summa_bez_nds(summa(doc_perinod))

# работа с листом Акты

num_akt = input('Введите номер Акта: ')  # запрос номера Акта выполненных работ

akt_num = sheet3['A1'].value
akt_num = akt_num.replace('number', num_akt)
sheet3['A1'] = akt_num  # записываю номер акта в ячейку

akt_data = sheet3['A1'].value
akt_data = akt_data.replace('data', data_doc)
sheet3['A1'] = akt_data  # записываю дату акта в  ячейку

akt_period = sheet3['B8'].value
akt_period = akt_period.replace('period', doc_perinod.capitalize())
sheet3['B8'] = akt_period  # записываю периода акта

akt_itog = sheet3['A13'].value  # считываю строку для указания итоговой суммы оплаты
akt_itog: object = akt_itog.replace('summa', summa(doc_perinod))  # указываю итоговую сумму
sheet3['A13'] = akt_itog

sheet3['A14'] = itog2_doc = nds_text = nds_text2  # сумму прописью переношу с первого листа

book.save(f'документы за {doc_perinod}.xlsx')  # создаю книгу с названием периода
