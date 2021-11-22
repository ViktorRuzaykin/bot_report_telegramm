import openpyxl
import shutil
import telebot
import os

report_file = 'Сводка по ИС ППС11 СМУ 11.2 Окунайский.xlsx'
report_sheet = 'Сводка'
bot = telebot.TeleBot('1937067149:AAGnI-FHNXK_9BYYJBUYWqKWnBHXp__l97I')
line_report = 11  # количество строк для итерации(ключи для словаря book_report)

book_report = {
    'СМР не закончены': {},
    'Полка': {},
    'На подписи СТНГ': {},
    'Не подписывает ИТЦ': {},
    'Оформлена': {},
    'На согласовании': {},
    'На оформлении': {},
    'На подписи ГСП': {},
    'На подписи у ИТЦ': {},
    'Передано в ПТГ': {},
    'Итого': {}
}

dict_data = {
    'Траншея': None,
    'Подушка': None,
    'Геоматрица': None,
    'Укладка': None,
    'Присыпка': None
}


def read_report():
    if os.path.exists("C:\\bot_report\\Сводка по ИС ППС11 СМУ 11.2 Окунайский.xlsx"):
        os.remove("C:\\bot_report\\Сводка по ИС ППС11 СМУ 11.2 Окунайский.xlsx")
    shutil.copyfile("C:\\Облако\\ИД\\Сводка по ИС ППС11 СМУ 11.2 Окунайский.xlsx",
                    "C:\\bot_report\\Сводка по ИС ППС11 СМУ 11.2 Окунайский.xlsx")
    wb = openpyxl.reader.excel.load_workbook(filename=report_file, data_only=True)
    sheet = wb[report_sheet]
    list_key_book_report = book_report.keys()
    count = 3
    for key in list_key_book_report:
        book_report[key] = {
            'Траншея': format(sheet[f'B{count}'].value, '.2f'),
            'Подушка': format(sheet[f'C{count}'].value,'.2f'),
            'Геоматрица': format(sheet[f'D{count}'].value, '.2f'),
            'Укладка': format(sheet[f'E{count}'].value, '.2f'),
            'Присыпка': format(sheet[f'F{count}'].value, '.2f')
        }
        count += 1


# read_report()


def write_str(key):
    str_for_send = f'{key}:'
    work_dict = book_report[key]
    for str in work_dict:
        # print(str, work_dict[str])
        str_for_send += f'\n{str} - {work_dict[str]} м.'
    return str_for_send

# print(str(book_report['На оформлении']))

# print(write_str('На оформлении'))


"""@bot.message_handler(content_types=['text'])
def fr(message):
    if message.text == "На оформлении":
        message_str = write_str('На оформлении')
        bot.send_message(message.from_user.id, message_str)

    if message.text == "не подписывает ИТЦ":
        message_str = write_str('не подписывает ИТЦ')
        bot.send_message(message.from_user.id, message_str)
    else:
        bot.send_message(message.from_user.id, 'Попроуй еще раз')
"""


@bot.message_handler(commands=['сводка'])
def start_message(message):
    read_report()
    markup = telebot.types.InlineKeyboardMarkup()
    markup.add(telebot.types.InlineKeyboardButton(text='СМР не закончены', callback_data='СМР не закончены'))
    markup.add(telebot.types.InlineKeyboardButton(text='Полка', callback_data='Полка'))
    markup.add(telebot.types.InlineKeyboardButton(text='Не подписывает ИТЦ', callback_data='Не подписывает ИТЦ'))
    markup.add(telebot.types.InlineKeyboardButton(text='Оформлена', callback_data='Оформлена'))
    markup.add(telebot.types.InlineKeyboardButton(text='На согласовании', callback_data='На согласовании'))
    markup.add(telebot.types.InlineKeyboardButton(text='На оформлении', callback_data='На оформлении'))
    markup.add(telebot.types.InlineKeyboardButton(text='На подписи ГСП', callback_data='На подписи ГСП'))
    markup.add(telebot.types.InlineKeyboardButton(text='На подписи у ИТЦ', callback_data='На подписи у ИТЦ'))
    markup.add(telebot.types.InlineKeyboardButton(text='Передано в ПТГ', callback_data='Передано в ПТГ'))
    markup.add(telebot.types.InlineKeyboardButton(text='Итого', callback_data='Итого'))

    bot.send_message(message.chat.id, text="Какие данные интересуют?", reply_markup=markup)


@bot.callback_query_handler(func=lambda call: True)
def query_handler(call):
    # bot.answer_callback_query(callback_query_id=call.id, text='Спасибо за честный ответ!')
    answer = ''
    if call.data == 'СМР не закончены':
        # message_str = write_str('СМР не закончены')
        answer = write_str('СМР не закончены')
    elif call.data == 'Полка':
        answer = write_str('Полка')
    elif call.data == 'Не подписывает ИТЦ':
        answer = write_str('Не подписывает ИТЦ')
    elif call.data == 'Оформлена':
        answer = write_str('Оформлена')
    elif call.data == 'На согласовании':
        answer = write_str('На согласовании')
    elif call.data == 'На оформлении':
        answer = write_str('На оформлении')
    elif call.data == 'На подписи ГСП':
        answer = write_str('На подписи ГСП')
    elif call.data == 'На подписи у ИТЦ':
        answer = write_str('На подписи у ИТЦ')
    elif call.data == 'Передано в ПТГ':
        answer = write_str('Передано в ПТГ')
    elif call.data == 'Итого':
        answer = write_str('Итого')

    bot.send_message(call.message.chat.id, answer)


bot.polling(none_stop=True, interval=0)
