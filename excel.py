import xlsxwriter

workbook = xlsxwriter.Workbook('excel.xlsx')
worksheet = workbook.add_worksheet('TDSheet')
x = 0
y = 0

worksheet.insert_image(1, 1, 'img.png', {'x_scale': 2, 'y_scale': 2})

worksheet.write('B7:D7', 'Клиент:')
worksheet.write('B8', 'Валюта:')

worksheet.write('R7', 'Контрагент:')
worksheet.write('R8', 'Склад:')
worksheet.write('R9', 'Желаемая дата отгрузки:')
worksheet.write('R10', 'Не отгружать частями:')
worksheet.write('B12', 'График оплаты:')


merge_format = workbook.add_format({
    'bold': True,
    'border': 2,
    'align': 'left',
})

worksheet.write('B14', '№', merge_format)
worksheet.merge_range('C14:Y14', 'Вариант оплаты', merge_format)
worksheet.merge_range('Z14:AC14', 'Дата', merge_format)
worksheet.merge_range('AD14:AE14', '%', merge_format)
worksheet.merge_range('AF14:AH14', 'Сумма', merge_format)

worksheet.write('B15', '', merge_format)
worksheet.merge_range('C15:Y15', '', merge_format)
worksheet.merge_range('Z15:AC15', '', merge_format)
worksheet.merge_range('AD15:AE15', '', merge_format)
worksheet.merge_range('AF15:AH15', '', merge_format)

worksheet.merge_range('B19:B20', '№', merge_format)
worksheet.write('B21', '', merge_format)

worksheet.merge_range('C19:N20', 'Товары (работы, услуги)', merge_format)
worksheet.merge_range('O19:P20', 'План', merge_format)
worksheet.merge_range('Q19:R20', 'Заказ', merge_format)
worksheet.merge_range('S19:T20', 'Ед.изм.', merge_format)
worksheet.merge_range('U19:W20', 'Цена', merge_format)
worksheet.merge_range('X19:Z20', 'Сумма', merge_format)
worksheet.merge_range('AA19:AH20', 'Комментарий', merge_format)
worksheet.merge_range('AI19:AV20', 'Товары (работы, услуги)', merge_format)
worksheet.merge_range('AW19:AY19', 'Скидка', merge_format)
worksheet.merge_range('AZ19:BA20', 'Сумма со скидкой', merge_format)
worksheet.write('AW20', '%', merge_format)
worksheet.merge_range('AX20:AY20', 'Сумма', merge_format)

worksheet.write('B17', 'Товары:')

queryset = [
    {
        "id": 1,
        'name': 'Product1',
        'count': 2,
        'price': 3000,
        'discount': 30,  # %
        'comment': 'good product'
    },
    {
        "id": 2,
        'name': 'Product2',
        'count': 3,
        'price': 4000,
        'discount': 40,  # %
        'comment': 'good product'
    },
    {
        "id": 3,
        'name': 'Product3',
        'count': 4,
        'price': 3000,
        'discount': 50,
        'comment': 'good product'
    },
    {
        "id": 4,
        'name': 'Product4',
        'count': 4,
        'price': 5000,
        'discount': 50,
        'comment': 'good product'
    },
]

worksheet.set_column(0, 50, 2)  # razmer yacheyki


for a in queryset:
    print(a)
    id = a['id']
    name = a['name']
    price = a['price']
    count = a['count']
    comment = a['comment']
    discount = int(a['discount'])
    v = count*price
    worksheet.write(x+20, y+1, id, merge_format)
    worksheet.merge_range(x+20, y+2, x+20, y+13, name, merge_format) #tovari i uslugi
    worksheet.merge_range(x+20, y+14, x+20, y+15, '', merge_format) #plan
    worksheet.merge_range(x+20, y+16, x+20, y+17, id, merge_format) #order
    worksheet.merge_range(x+20, y+18, x+20, y+19, count, merge_format) #ed iz
    worksheet.merge_range(x+20, y+20, x+20, y+22, price, merge_format) # price
    worksheet.merge_range(x+20, y+23, x+20, y+25, count*price, merge_format) #all
    worksheet.merge_range(x+20, y+26, x+20, y+33, comment, merge_format) #comment
    worksheet.merge_range(x+20, y+34, x+20, y+47, name, merge_format) #tovari
    worksheet.write(x+20, y+48, discount, merge_format)
    worksheet.merge_range(x+20, y+49, x+20, y+50, price*count, merge_format) #summa
    worksheet.merge_range(x+20, y+51, x+20, y+52, v-(discount*price)/100, merge_format) #summa so sidkoy

    x += 1



worksheet.write(x+22, y+1, 'Итого заказано на сумму:')

worksheet.write(x+25, y+1, 'Скидка на заказ в целом:')
worksheet.write(x+29, y+1, 'Сумма-')
worksheet.write(x+25, y+8, '% -         ')

worksheet.write(x+27, y+1, 'Итого сумма заказа с учетом скидки:')

worksheet.write(x+29, y+1, 'Заказ принял:       ')
worksheet.write(x+29, y+16, 'Дата:       ')
worksheet.write(x+29, y+24, 'Подпись:       ')





workbook.close()
