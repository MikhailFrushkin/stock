import os

import pandas as pd
from PyQt5.QtWidgets import QMessageBox


def check_stock(self, file_path, min_vitrina=False, plus=False, minus=False,
                name_file_min_vitrina=None):
    try:
        rooms = []
        df = pd.read_excel(file_path, skiprows=14)
        df.drop(["Поставщик", "Наименование"], axis=1, inplace=True)
        bu = list(df['БЮ'].unique())[0]
        rdiff_groups = sorted(list(df[(df.Склад == 'RDiff_{}'.format(bu))]['ТГ'].unique()))
        sklad_list = list(df['Склад'].unique())
        for i in sklad_list:
            if i.startswith(('A', 'a')):
                rooms.append(i)
    except Exception as ex:
        QMessageBox.critical(self, 'Ошибка!', 'Ошибка при чтении файла c 6.1 {}\n{}'.format(file_path, ex))
        self.restart1()
    groups_tdd = [
        11, 12, 22, 23, 24, 25, 26, 27, 28, 29,
        '11', '12', '22', '23', '24', '25', '26', '27', '28', '29'
    ]
    ignored_groups = ['112', '174', '175', '176', '177']

    sklad_tdd = df[((df.Склад == '011_{}'.format(bu)) |
                    (df.Склад == '012_{}'.format(bu))) &
                   (df.Склад != '012_{}-OX'.format(bu)) &
                   ~(df.НГ.isin(ignored_groups)) &
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

    if min_vitrina:
        try:
            df_min = pd.read_excel(name_file_min_vitrina).fillna(0)
            df_min['Артикул'] = df_min['Артикул'].astype('int64')

            daily_sales = sklad_tdd.groupby([pd.Grouper(key='Код \nноменклатуры')]).agg(
                Количество_склад=('Доступно', 'sum')).reset_index()

            df_tdd_min = df[(df.Склад == 'V_{}'.format(bu)) &
                            (df.Доступно > 0) &
                            (df.ТГ.isin(groups_tdd))
                            ]
            df_min_vitrina = pd.merge(df_tdd_min, df_min, left_on='Код \nноменклатуры', right_on='Артикул')
            df_min_vitrina = pd.merge(df_min_vitrina, daily_sales, left_on='Код \nноменклатуры',
                                      right_on='Код \nноменклатуры')
            df_min_vitrina = df_min_vitrina.assign(Разница=df_min_vitrina['Доступно'] - df_min_vitrina['Количество мин'])
            df_min_vitrina = df_min_vitrina[(df_min_vitrina['Разница'] < 0)]
            df_min_vitrina.drop(
                ["Reason code", "Физические \nзапасы", "Продано", "Зарезерви\nровано", 'Артикул'], axis=1,
                inplace=True)

            name_dir_min = '{}\{}'.format(self.current_dir, 'Мин.витрина')
            name_dir_pst = '{}\{}'.format(self.current_dir, 'Файлы для импорта')

            if not os.path.exists(name_dir_pst) or not os.path.isdir(name_dir_pst):
                os.mkdir(name_dir_pst)
            try:
                for f in os.listdir(name_dir_pst):
                    os.remove(os.path.join(name_dir_pst, f))
            except:
                pass

            if not os.path.exists(name_dir_min) or not os.path.isdir(name_dir_min):
                os.mkdir(name_dir_min)
            writer = pd.ExcelWriter(f'{name_dir_min}\Мин по группам.xlsx', engine='xlsxwriter')
            for group in groups_tdd:
                try:
                    temp_df = df_min_vitrina[(df_min_vitrina.ТГ == group)]
                    if len(temp_df) != 0:

                        data = {
                            'Описание товара, Удалить перед импортом': [],
                            'Номенклатура': [],
                            'Кол-во': [],
                            'Со склада': [],
                            'С ячейки': [],
                            'На БЮ': [],
                            'На склад': [],
                            'Дата отгрузки': [],
                            'Промо': [],
                            'С "reason code"': [],
                            'На "reason code"': [],
                            'С профиля учета': [],
                            'На профиль учета': [],
                            'В ячейку': [],
                            'С сайта': [],
                            'На сайт': [],
                            'С владельца': [],
                            'На владельца': [],
                            'Из партии': [],
                            'В партию': [],
                            'Из ГТД': [],
                            'В ГТД': [],
                            'С серийного номера': [],
                            'На серийный номер': []
                        }

                        temp_df.sort_values(by='Код \nноменклатуры'). \
                            to_excel(writer, sheet_name='Min. ТГ {}'.format(group), index=False, na_rep='')
                        worksheet = writer.sheets['Min. ТГ {}'.format(group)]
                        set_column_min(temp_df, worksheet)

                        temp_dict = (temp_df
                                     .groupby("Код \nноменклатуры")
                                     .apply(lambda x: x.drop(columns="Код \nноменклатуры").to_dict("records"))
                                     .to_dict())
                        for key, value in temp_dict.items():
                            art_df_sklad = sklad_tdd[(sklad_tdd['Код \nноменклатуры'] == int(key))]
                            for i, row in art_df_sklad.iterrows():
                                if value[0]['Разница'] < 0:
                                    data['Описание товара, Удалить перед импортом'].append(row['Описание товара'])

                                    data['Номенклатура'].append(row['Код \nноменклатуры'])
                                    data['Со склада'].append('012_825')
                                    data['С ячейки'].append(row['Местоположение'])
                                    data['На БЮ'].append('825')
                                    data['На склад'].append('V_825')
                                    data['Дата отгрузки'].append('')
                                    data['Промо'].append('')
                                    data['С "reason code"'].append('')
                                    data['На "reason code"'].append('')
                                    data['С профиля учета'].append('')
                                    data['На профиль учета'].append('')
                                    data['В ячейку'].append('V-sales_825')
                                    data['С сайта'].append('')
                                    data['На сайт'].append('')
                                    data['С владельца'].append('')
                                    data['На владельца'].append('')
                                    data['Из партии'].append('')
                                    data['В партию'].append('')
                                    data['Из ГТД'].append('')
                                    data['В ГТД'].append('')
                                    data['С серийного номера'].append('')
                                    data['На серийный номер'].append('')
                                    if row['Доступно'] < -(value[0]['Разница']):
                                        data['Кол-во'].append(row['Доступно'])
                                        value[0]['Разница'] += row['Доступно']
                                    else:
                                        data['Кол-во'].append(-(value[0]['Разница']))
                                        break
                        temp_df_pst = pd.DataFrame(data=data)
                        temp_df_pst['Со склада'] = temp_df_pst['Со склада'].astype('string')
                        temp_df_pst_n = temp_df_pst.convert_dtypes()

                        writer_pst = pd.ExcelWriter(f'Файлы для импорта/Пст ТГ.{group}.xlsx', engine='xlsxwriter')
                        temp_df_pst_n.to_excel(writer_pst, sheet_name='ТГ {}'.format(group), index=False, na_rep='')
                        worksheet = writer_pst.sheets['ТГ {}'.format(group)]
                        set_column_pst(temp_df, worksheet)
                        writer_pst.close()

                except Exception as ex:
                    QMessageBox.critical(self, 'Ошибка!', f'Формирования файлов пст\n{ex}')
                    self.restart1()
            writer.close()
        except Exception as ex:
            QMessageBox.critical(self, 'Ошибка!', 'Ошибка записи результата мин.витрина{}'.format(ex))
            self.restart1()

    if none_tdd:
        new_df_tdd = set_to_df(self, none_tdd, sklad_tdd)
    else:
        new_df_tdd = pd.DataFrame()

    if none_mebel:
        new_df_mebel = set_to_df(self, none_mebel, sklad_mebel)
    else:
        new_df_mebel = pd.DataFrame()
    write_exsel(self, df, none_all, new_df_tdd, new_df_mebel, reserved_tdd, units_tdd_df)

    if plus:
        writer = pd.ExcelWriter('Положительные RDiff(0 на V_Sales).xlsx', engine='xlsxwriter',
                                engine_kwargs={'options': {'strings_to_numbers': True}})
        workbook = writer.book
        cell_format = workbook.add_format({'align': 'left', 'valign': 'top', 'font_size': 12})

        for i in rdiff_groups:
            try:
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
                        QMessageBox.critical(self, 'Ошибка!', 'Ошибка записи плюсовых рдиффов{}'.format(ex))
                        self.restart1()

                new_df = temp

                if len(new_df) > 0:
                    try:
                        new_df.sort_values(by='Код \nноменклатуры', ascending=False). \
                            to_excel(writer, sheet_name='ТГ {}'.format(i), index=False, na_rep='')
                        worksheet = writer.sheets['ТГ {}'.format(i)]
                        set_column(df, worksheet, cell_format=cell_format)
                    except Exception as ex:
                        QMessageBox.critical(self, 'Ошибка!', f'{ex}')
                        self.restart1()

            except Exception as ex:
                QMessageBox.critical(self, 'Ошибка!', f'{ex}')
                self.restart1()

        writer.close()
    if minus:
        try:
            writer = pd.ExcelWriter('Минусовые RDiff, которые нужно проверить.xlsx', engine='xlsxwriter')
            workbook = writer.book
            cell_format = workbook.add_format({'align': 'left', 'valign': 'top', 'font_size': 12})
            for i in rdiff_groups:
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
                            QMessageBox.critical(self, 'Ошибка!', f'{art}{ex}')
                            self.restart1()
                    new_df = temp
                    if len(new_df) > 0:
                        try:
                            new_df.sort_values(by='Код \nноменклатуры', ascending=False). \
                                to_excel(writer, sheet_name='ТГ {}'.format(i), index=False, na_rep='')
                            worksheet = writer.sheets['ТГ {}'.format(i)]
                            set_column(df, worksheet, cell_format=cell_format)
                        except Exception as ex:
                            QMessageBox.critical(self, 'Ошибка!', f'{ex}')
                            self.restart1()

                except Exception as ex:
                    QMessageBox.critical(self, 'Ошибка!', f'{ex}')
                    self.restart1()

            writer.close()
        except Exception as ex:
            QMessageBox.critical(self, 'Ошибка!', 'Ошибка записи минусовых рдиффов{}'.format(ex))
            self.restart1()

    return '\nНе выставленный товар:\nТдд: {}\nМебель: {}\nЗарезервированный товар под 0 с V_sales: {}\nЕдиничек: {}'. \
        format(len(none_tdd), len(none_mebel), len(reserved_tdd), len(units_tdd_df))


def write_exsel(self, df, none_all, new_df_tdd, new_df_mebel, reserved_tdd, units_tdd_df):
    try:
        writer = pd.ExcelWriter('Результат сверки стока.xlsx', engine='xlsxwriter')
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

        if len(new_df_mebel) != 0:
            new_df_mebel.sort_values(by='Код \nноменклатуры'). \
                to_excel(writer, sheet_name='Мебель', index=False, na_rep='')
            worksheet3 = writer.sheets['Мебель']
            set_column(new_df_mebel, worksheet3, cell_format=cell_format)

        if len(reserved_tdd) != 0:
            reserved_tdd.sort_values(by='Код \nноменклатуры'). \
                to_excel(writer, sheet_name='Резерв ТДД под 0', index=False, na_rep='')
            worksheet4 = writer.sheets['Резерв ТДД под 0']
            set_column(reserved_tdd, worksheet4, cell_format=cell_format)

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
                QMessageBox.critical(self, 'Ошибка!', f'{ex}')
                self.restart1()
        writer.close()
    except Exception as ex:
        QMessageBox.critical(self, 'Ошибка!', 'Ошибка записи результата {}'.format(ex))
        self.restart1()


def set_to_df(self, art_set, df_sklad):
    new_df = pd.DataFrame()
    for art in art_set:
        try:
            new_df = pd.concat([new_df, df_sklad.loc[(df_sklad['Код \nноменклатуры'] == art)]])
        except Exception as ex:
            QMessageBox.critical(self, 'Ошибка!', f'{art}{ex}')
            self.restart1()

    return new_df


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


def set_column_min(df, worksheet, cell_format=None):
    (max_row, max_col) = df.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.set_column('A:A', 5, cell_format)
    worksheet.set_column('B:B', 10, cell_format)
    worksheet.set_column('C:D', 20, cell_format)
    worksheet.set_column('E:E', 23, cell_format)
    worksheet.set_column('F:F', 60, cell_format)
    worksheet.set_column('G:H', 10, cell_format)
    worksheet.set_column('I:I', 15, cell_format)
    worksheet.set_column('J:K', 20, cell_format)
    worksheet.set_column('L:L', 17, cell_format)


def set_column_pst(df, worksheet, cell_format=None):
    (max_row, max_col) = df.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.set_column('A:A', 65, cell_format)
    worksheet.set_column('B:B', 17, cell_format)
    worksheet.set_column('C:G', 17, cell_format)
    worksheet.set_column('H:X', 15, cell_format)
