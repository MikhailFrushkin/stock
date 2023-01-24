import io
import json
import os
import sys
import time
from pathlib import Path

import httplib2
import pandas as pd
import requests
from PIL import Image
from loguru import logger

bu = None
rooms = []
rdiff_groups = []


def file_name() -> list:
    """нахождение файла с 6.1
    :return имена файлов"""
    file_exsel = sorted(Path('').glob('*.xlsx'))
    if len(file_exsel) == 0:
        logger.error('Нет файлов с 6.1')
        exit_error()
    elif len(file_exsel) > 1 and 'Результат сверки стока.xlsx' not in [str(i) for i in file_exsel] \
            and 'Положительные RDiff(0 на V_Sales).xlsx' not in [str(i) for i in file_exsel] \
            and 'Минусовые RDiff, которые нужно проверить.xlsx' \
            not in [str(i) for i in file_exsel]:
        logger.error('Удалите лишние файлы с расширением .xlsx\n'
                     'Вожможно файл открыт')
        exit_error()
    return file_exsel


def read_file(files: list):
    print('Чтение файла с остатками...')
    try:
        global bu
        global rooms
        global rdiff_groups
        df = pd.read_excel(files[0], skiprows=14)
        df.drop(["Поставщик", "Наименование"], axis=1, inplace=True)
        bu = list(df['БЮ'].unique())[0]
        rdiff_groups = sorted(list(df[(df.Склад == 'RDiff_{}'.format(bu))]['ТГ'].unique()))
        sklad_list = list(df['Склад'].unique())
        for i in sklad_list:
            if i.startswith(('A', 'a')):
                rooms.append(i)
    except Exception as ex:
        logger.error('Ошибка при чтении файла c 6.1 {}\n{}'.format(files[0], ex))
        exit_error()

    groups_tdd = [
        11, 12, 22, 23, 24, 25, 26, 27, 28, 29,
        '11', '12', '22', '23', '24', '25', '26', '27', '28', '29'
    ]
    sklad_tdd = df[((df.Склад == '011_{}'.format(bu)) |
                    (df.Склад == '012_{}'.format(bu)) &
                    ((df.Склад != '012_{}-OX'.format(bu))) &
                    (df.НГ != '112')) &
                   (df.Доступно > 0) &
                   (df.ТГ.isin(groups_tdd))
                   ]
    sklad_art_list_tdd = list(sklad_tdd['Код \nноменклатуры'].unique())
    df_tdd = df[((df.Склад == 'V_{}'.format(bu)) |
                 (df.Склад == 'S_{}'.format(bu))) &
                (df.Доступно > 0) &
                (df.ТГ.isin(groups_tdd))
                ]
    tdd_art_list = sorted(list(df_tdd['Код \nноменклатуры'].unique()))

    units_group_list = sorted(list(df['ТГ'].unique()))
    units_tdd_df = df[((df.Склад == 'V_{}'.format(bu)) |
                       (df.Склад == 'S_{}'.format(bu))) &
                      (df.Доступно == 1) &
                      (df.ТГ.isin(units_group_list))
                      ]

    reserved_tdd = df[((df.Склад == 'V_{}'.format(bu)) |
                       (df.Склад == 'S_{}'.format(bu))) &
                      ((df['Доступно'].isnull()) &
                       (df['Физические \nзапасы'] > 0)) &
                      (df.ТГ.isin(groups_tdd))
                      ]

    groups_mebel = [
        30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40,
        '30', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40'
    ]
    sklad_mebel = df[((df.Склад == '011_{}'.format(bu)) |
                      (df.Склад == '012_{}'.format(bu))) &
                     (df.Доступно > 0) &
                     (df.ТГ.isin(groups_mebel))
                     ]
    df_mebel = df[((df.Склад.isin(rooms)) |
                   (df.Склад == 'V_{}'.format(bu))) &
                  (df.Доступно > 0) &
                  (df.ТГ.isin(groups_mebel))
                  ]
    sklad_art_list_mebel = list(sklad_mebel['Код \nноменклатуры'].unique())
    mebel_art_list = list(df_mebel['Код \nноменклатуры'].unique())

    none_tdd = set(sklad_art_list_tdd).difference(set(tdd_art_list))

    none_mebel = set(sklad_art_list_mebel).difference(set(mebel_art_list))

    none_all = pd.DataFrame(
        {'': ['Тдд', 'Мебель', 'Зарезервированный товар под 0 с V_sales', 'Единичек'],
         'Количество': [len(none_tdd), len(none_mebel), len(reserved_tdd), len(units_tdd_df)]})

    print('\n--------------------------------------------'
          '\nНе выставленный товар:'
          '\nТдд: {}'
          '\nМебель: {}'
          '\n--------------------------------------------'
          '\nЗарезервированный товар под 0 с V_sales: {}'
          '\n--------------------------------------------'
          '\nЕдиничек: {}'
          '\n--------------------------------------------'.
          format(len(none_tdd), len(none_mebel), len(reserved_tdd), len(units_tdd_df)))

    # if len(reserved_tdd) != 0:
    #     for num, row in enumerate(reserved_tdd.values):
    #         try:
    #             art = row[3]
    #             data = parse(art)
    #             save_image(art, data['picture'])
    #         except Exception as ex:
    #             logger.error('Ошибка парсинга артикула {}\n{}'.format(row[3], ex))

    if none_tdd:
        new_df_tdd = set_to_df(none_tdd, sklad_tdd)
        # for num, row in enumerate(new_df_tdd.values):
        #     try:
        #         art = row[3]
        #         data = parse(art)
        #         save_image(art, data['picture'])
        #     except Exception as ex:
        #         logger.error('Ошибка парсинга артикула {}\n{}'.format(row[3], ex))
    else:
        new_df_tdd = pd.DataFrame()

    if none_mebel:
        new_df_mebel = set_to_df(none_mebel, sklad_mebel)
        # for num, row in enumerate(new_df_mebel.values):
        # try:
        #     art = row[3]
        #     data = parse(art)
        #     save_image(art, data['picture'])
        # except Exception as ex:
        #     logger.error('Ошибка парсинга артикула {}\n{}'.format(row[3], ex))
    else:
        new_df_mebel = pd.DataFrame()
    write_exsel(df, none_all, new_df_tdd, new_df_mebel, reserved_tdd, units_tdd_df)


def write_to_excel_rdiff(files):
    print('Проверка RDiff...')
    try:
        df = pd.read_excel(files[0], skiprows=14)
        df.drop(["Поставщик", "Наименование"], axis=1, inplace=True)
    except Exception as ex:
        logger.error('Ошибка при чтении файла c 6.1 {}\n{}'.format(files[0], ex))
        exit_error()

    writer = pd.ExcelWriter('Положительные RDiff(0 на V_Sales).xlsx', engine='xlsxwriter',
                            engine_kwargs={'options': {'strings_to_numbers': True}})
    workbook = writer.book
    cell_format = workbook.add_format({'align': 'left', 'valign': 'top', 'font_size': 12})

    for i in rdiff_groups:
        try:
            print('Сканируются плюсовые RDiff ТГ {}'.format(i))
            temp = df[
                (df.ТГ == i) &
                (df.Склад == 'RDiff_{}'.format(bu)) &
                (df.Доступно > 0)
                ]
            temp_art_list = sorted(list(temp['Код \nноменклатуры'].unique()))
            temp = pd.DataFrame()
            for art in temp_art_list:
                try:
                    rdiff_art_vls = df[(df.Склад == 'V_{}'.format(bu))
                                       & (df['Код \nноменклатуры'] == art)
                                       & (df.Доступно > 0)]
                    if len(rdiff_art_vls) == 0:
                        temp = pd.concat([temp, df.loc[(df['Код \nноменклатуры'] == art)]])
                except Exception as ex:
                    print(art, ex)
            new_df = temp

            if len(new_df) > 0:
                try:
                    new_df.sort_values(by='Код \nноменклатуры', ascending=False). \
                        to_excel(writer, sheet_name='ТГ {}'.format(i), index=False, na_rep='')
                    worksheet = writer.sheets['ТГ {}'.format(i)]
                    set_column(df, worksheet, cell_format=cell_format)
                except Exception as ex:
                    print(ex)
        except Exception as ex:
            print(ex)
    writer.close()


def write_to_excel_minus_rdiff(files):
    try:
        df = pd.read_excel(files[0], skiprows=14)
        df.drop(["Поставщик", "Наименование"], axis=1, inplace=True)
    except Exception as ex:
        logger.error('Ошибка при чтении файла c 6.1 {}\n{}'.format(files[0], ex))
        exit_error()

    writer = pd.ExcelWriter('Минусовые RDiff, которые нужно проверить.xlsx', engine='xlsxwriter',
                            engine_kwargs={'options': {'strings_to_numbers': True}})
    workbook = writer.book
    cell_format = workbook.add_format({'align': 'left', 'valign': 'top', 'font_size': 12})
    for i in rdiff_groups:
        print('Сканируются минусовые RDiff ТГ {}'.format(i))
        try:
            group_rdiff_art = df[
                (df.ТГ == i) &
                (df.Склад == 'RDiff_{}'.format(bu)) &
                (df.Доступно < 0)
                ]

            list_minus = list(group_rdiff_art['Код \nноменклатуры'].unique())
            temp = pd.DataFrame()
            for art in list_minus:
                try:
                    temp = pd.concat([temp, df.loc[df['Код \nноменклатуры'] == art]])
                except Exception as ex:
                    print(art, ex)
            new_df = temp
            if len(new_df) > 0:
                try:
                    new_df.sort_values(by='Код \nноменклатуры', ascending=False). \
                        to_excel(writer, sheet_name='ТГ {}'.format(i), index=False, na_rep='')
                    worksheet = writer.sheets['ТГ {}'.format(i)]
                    set_column(df, worksheet, cell_format=cell_format)
                except Exception as ex:
                    print(ex)
        except Exception as ex:
            print(ex)
    writer.close()


def set_to_df(art_set, df_sklad):
    new_df = pd.DataFrame()
    for art in art_set:
        try:
            new_df = pd.concat([new_df, df_sklad.loc[(df_sklad['Код \nноменклатуры'] == art)]])
        except Exception as ex:
            print(art, ex)
    return new_df


def write_exsel(df, none_all, new_df_tdd, new_df_mebel, reserved_tdd, units_tdd_df):
    try:
        writer = pd.ExcelWriter('Результат сверки стока.xlsx', engine='xlsxwriter',
                                engine_kwargs={'options': {'strings_to_numbers': True}})
        workbook = writer.book
        cell_format = workbook.add_format({'align': 'left', 'valign': 'top', 'font_size': 12})

        none_all.to_excel(writer, sheet_name='Общий', index=False, na_rep='')
        worksheet = writer.sheets['Общий']
        worksheet.autofit()

        if len(new_df_tdd) != 0:
            new_df_tdd.sort_values(by='Код \nноменклатуры'). \
                to_excel(writer, sheet_name='ТДД', index=False, na_rep='')
            worksheet2 = writer.sheets['ТДД']
            set_column(new_df_tdd, worksheet2, cell_format=cell_format)

            # list_art = new_df_tdd.sort_values(by='Код \nноменклатуры')["Код \nноменклатуры"].tolist()
            # for num, item in enumerate(list_art, start=1):
            #     insert_images(worksheet2, num, cell_format, item)

        if len(new_df_mebel) != 0:
            new_df_mebel.sort_values(by='Код \nноменклатуры'). \
                to_excel(writer, sheet_name='Мебель', index=False, na_rep='')
            worksheet3 = writer.sheets['Мебель']
            set_column(new_df_mebel, worksheet3, cell_format=cell_format)

            # list_art = new_df_mebel.sort_values(by='Код \nноменклатуры')["Код \nноменклатуры"].tolist()
            # for num, item in enumerate(list_art, start=1):
            #    insert_images(worksheet3, num, cell_format, item)

        if len(reserved_tdd) != 0:
            reserved_tdd.sort_values(by='Код \nноменклатуры'). \
                to_excel(writer, sheet_name='Резерв ТДД под 0', index=False, na_rep='')
            worksheet4 = writer.sheets['Резерв ТДД под 0']
            set_column(reserved_tdd, worksheet4, cell_format=cell_format)

            # list_art = reserved_tdd.sort_values(by='Код \nноменклатуры')["Код \nноменклатуры"].tolist()
            # for num, item in enumerate(list_art, start=1):
            #     insert_images(worksheet4, num, cell_format, item)

        groups_tdd = sorted(list(df['ТГ'].unique()))

        for i in groups_tdd:
            try:
                temp_df = units_tdd_df[(units_tdd_df.ТГ == i)]
                if len(temp_df) != 0:
                    temp_df.sort_values(by='Код \nноменклатуры'). \
                        to_excel(writer, sheet_name='Ед. ТГ {}'.format(i), index=False, na_rep='')
                    worksheet = writer.sheets['Ед. ТГ {}'.format(i)]
                    set_column(temp_df, worksheet, cell_format=cell_format)
            except Exception as ex:
                logger.error(ex)
        writer.close()
    except Exception as ex:
        logger.error('Ошибка записи результата {}'.format(ex))
        exit_error()


def set_column(df, worksheet, cell_format=None):
    (max_row, max_col) = df.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.set_column('A:A', 5, cell_format)
    worksheet.set_column('B:B', 8, cell_format)
    worksheet.set_column('C:E', 22, cell_format)
    worksheet.set_column('F:F', 65, cell_format)
    worksheet.set_column('G:G', 22, cell_format)
    worksheet.set_column('H:I', 6, cell_format)
    worksheet.set_column('J:J', 17, cell_format)
    worksheet.set_column('K:K', 14, cell_format)
    worksheet.set_column('L:L', 17, cell_format)
    worksheet.set_column('M:W', 14, cell_format)


def insert_images(worksheet, num, cell_format, image):
    """Вставка картинки в 1й столбец"""
    try:
        worksheet.set_row(num, 120, cell_format)
        url = parse(art=image)['url']
        image_buffer, image = resize('img/{}.jpg'.format(image), (512, 512), format='JPEG')
        data = {'x_scale': 180 / image.width, 'y_scale': 160 / image.height, 'object_position': 1}
        worksheet.insert_image(num, 0, 'img/{}.jpg'.format(image), {
            'url': url,
            'image_data': image_buffer, **data})
    except Exception as ex:
        logger.error('ошибка вставки картинки {}'.format(ex))


def exit_error():
    """выход из консоли"""
    time.sleep(5)
    exit()


def buffer_image(image: Image, format: str = 'JPEG'):
    """Сохранение картинки из буфера"""
    buffer = io.BytesIO()
    image.save(buffer, format=format)
    return buffer, image


def resize(path: str, size: tuple[int, int], format='JPEG'):
    """изменение размера изображения"""
    image = Image.open(path)
    image = image.resize(size)
    return buffer_image(image, format)


def save_image(name, url):
    """сохранение изображения"""
    if not os.path.exists('img/{}.jpg'.format(name)):
        h = httplib2.Http()
        response, content = h.request(url)
        out = open('img/{}.jpg'.format(name), 'wb')
        out.write(content)
        out.close()


def parse(art):
    """Парсер данных с сайта по артикулу"""
    data = {
        'articul': art,
        'url': '',
        'name': '',
        'picture': 'https://upload.wikimedia.org/wikipedia/commons/9/'
                   '9a/%D0%9D%D0%B5%D1%82_%D1%84%D0%BE%D1%82%D0%BE.png'
    }
    if os.path.exists('json/{}.json'.format(art)):
        with open('json/{}.json'.format(art), 'r', encoding='utf-8') as file:
            data = json.load(file)
            return data
    else:
        try:
            print('Сканирование ', art)
            params = {
                'articul': art,
            }
            response = requests.get('https://hoff.ru/vue/catalog/product/', params=params).json()
            with open("json.json", "w", encoding='utf-8') as write_file:
                json.dump(response, write_file, indent=4, ensure_ascii=False)
            data['articul'] = response.get('data').get('articul')
            data['url'] = response.get('data').get('meta').get('canonical')
            data['name'] = response.get('data').get('name')
            data['picture'] = response.get('data').get('slider').get('previews')[0]
        except Exception:
            print('Не удалось найти на сайте: {}'.format(art))

        with open('json/{}.json'.format(art), "w", encoding='utf-8') as write_file:
            json.dump(data, write_file, indent=4, ensure_ascii=False)
        return data


if __name__ == "__main__":
    logger.add(sys.stderr, format="{time} {level} {message}", filter="my_module")
    read_file(file_name())
    write_to_excel_rdiff(file_name())
    write_to_excel_minus_rdiff(file_name())
    print('Завершено!')
    exit_error()
