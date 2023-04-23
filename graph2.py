from main import *


def Graphics2():
    lymf = float(name_LYMF.get())
    neu = float(name_NEU.get())
    cd3 = float(name_CD3.get())
    cd4 = float(name_CD4.get())
    cd8 = float(name_CD8.get())
    cd19 = float(name_CD19.get())
    date = str(name_dateAnal.get())

    wb2 = Workbook()
    ws2 = wb2.active
    curr_wb2 = load_workbook('ГрафикиСезоны.xlsx')
    ws2 = curr_wb2['гипотеза']

    ws2['C3'] = date
    ws2['C4'] = cd8
    ws2['C7'] = cd4
    ws2['C10'] = neu / lymf
    ws2['C13'] = neu / cd3
    ws2['C16'] = neu / cd8
    ws2['C19'] = neu / cd4

    ws2['C25'] = date
    ws2['C26'] = cd8
    ws2['C29'] = cd4
    ws2['C32'] = neu / lymf
    ws2['C35'] = neu / cd19
    ws2['C38'] = neu / cd4
    ws2['C41'] = neu / cd8

    number = Count()
    curr_wb2.save(f'{number}grapic.xlsx')
    curr_wb2.close()
