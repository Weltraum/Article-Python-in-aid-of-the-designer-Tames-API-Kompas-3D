import os
import re
import subprocess
import pythoncom
from win32com.client import Dispatch, gencache
from tkinter import Tk
from tkinter.filedialog import askopenfilenames


# Подключение к API7 программы Компас 3D
def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const


# Функция проверки, запущена-ли программа КОМПАС 3D
def is_running():
    proc_list = \
    subprocess.Popen('tasklist /NH /FI "IMAGENAME eq KOMPAS*"', shell=False, stdout=subprocess.PIPE).communicate()[0]
    return True if proc_list else False


# Посчитаем количество листов каждого из формата
def amount_sheet(doc7):
    sheets = {"A0": 0, "A1": 0, "A2": 0, "A3": 0, "A4": 0, "A5": 0}
    for sheet in range(doc7.LayoutSheets.Count):
        format = doc7.LayoutSheets.Item(sheet).Format  # sheet - номер листа, отсчёт начинается от 0
        sheets["A" + str(format.Format)] += 1 * format.FormatMultiplicity
    return sheets


# Прочитаем основную надпись чертежа
def stamp(doc7):
    for sheet in range(doc7.LayoutSheets.Count):
        style_filename = os.path.basename(doc7.LayoutSheets.Item(sheet).LayoutLibraryFileName)
        style_number = int(doc7.LayoutSheets.Item(sheet).LayoutStyleNumber)

        if style_filename in ['graphic.lyt', 'Graphic.lyt'] and style_number == 1:
            stamp = doc7.LayoutSheets.Item(sheet).Stamp
            return {"Scale": re.findall(r"\d+:\d+", stamp.Text(6).Str)[0],
                    "Designer": stamp.Text(110).Str}

    return 'Неопределенный стиль оформления'


# Подсчет технических требований, в том случае, если включена автоматическая нумерация
def count_demand(doc7, module7):
    IDrawingDocument = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IDrawingDocument'], pythoncom.IID_IDispatch)
    drawing_doc = module7.IDrawingDocument(IDrawingDocument)
    text_demand = drawing_doc.TechnicalDemand.Text

    count = 0  # Количество пунктов технических требований
    for i in range(text_demand.Count):  # Прохоим по каждой строчке технических требований
        if text_demand.TextLines[i].Numbering == 1:  # и проверяем, есть ли у строки нумерация
            count += 1

    # Если нет нумерации, но есть текст
    if not count and text_demand.TextLines[0]:
        count += 1

    return count


# Подсчёт размеров на чертеже, для каждого вида по отдельности
def count_dimension(doc7, module7):
    IKompasDocument2D = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D'],
                                                     pythoncom.IID_IDispatch)
    doc2D = module7.IKompasDocument2D(IKompasDocument2D)
    views = doc2D.ViewsAndLayersManager.Views

    count_dim = 0
    for i in range(views.Count):
        ISymbols2DContainer = views.View(i)._oleobj_.QueryInterface(module7.NamesToIIDMap['ISymbols2DContainer'],
                                                                    pythoncom.IID_IDispatch)
        dimensions = module7.ISymbols2DContainer(ISymbols2DContainer)

        # Складываем все необходимые раpмеры
        count_dim += dimensions.AngleDimensions.Count + \
                     dimensions.ArcDimensions.Count + \
                     dimensions.Bases.Count + \
                     dimensions.BreakLineDimensions.Count + \
                     dimensions.BreakRadialDimensions.Count + \
                     dimensions.DiametralDimensions.Count + \
                     dimensions.Leaders.Count + \
                     dimensions.LineDimensions.Count + \
                     dimensions.RadialDimensions.Count + \
                     dimensions.RemoteElements.Count + \
                     dimensions.Roughs.Count + \
                     dimensions.Tolerances.Count

    return count_dim


def parse_design_documents(paths):
    is_run = is_running()  # True, если программа Компас уже запущена

    module7, api7, const7 = get_kompas_api7()  # Подключаемся к программе
    app7 = api7.Application  # Получаем основной интерфейс программы
    app7.Visible = True  # Показываем окно пользователю (если скрыто)
    app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы

    table = []  # Создаём таблицу парметров
    for path in paths:
        doc7 = app7.Documents.Open(PathName=path,
                                   Visible=True,
                                   ReadOnly=True)  # Откроем файл в видимом режиме без права его изменять

        row = amount_sheet(doc7)  # Посчитаем кол-во листов каждого формат
        row.update(stamp(doc7))  # Читаем основную надпись
        row.update({
            "Filename": doc7.Name,  # Имя файла
            "CountTD": count_demand(doc7, module7),  # Количество пунктов технических требований
            "CountDim": count_dimension(doc7, module7),  # Количество пунктов технических требований
        })
        table.append(row)  # Добавляем строку параметров в таблицу

        doc7.Close(const7.kdDoNotSaveChanges)  # Закроем файл без изменения

    if not is_run: app7.Quit()  # Закрываем программу при необходимости
    return table


def print_to_excel(result):
    excel = Dispatch("Excel.Application")  # Подключаемся к программе Excel
    excel.Visible = True  # Делаем окно видимым
    wb = excel.Workbooks.Add()  # Добавляем новую книгу
    sheet = wb.ActiveSheet  # Получаем ссылку на активный лист

    # Создаём заголовок таблицы
    sheet.Range("A1:J1").value = ["Имя файла", "Разработчик", "Кол-во размеров", "Кол-во пунктов ТТ",
                                  "А0", "А1", "А2", "А3", "А4", "Масштаб"]

    # Заполняем таблицу
    for i, row in enumerate(result):
        sheet.Cells(i + 2, 1).value = row['Filename']
        sheet.Cells(i + 2, 2).value = row['Designer']
        sheet.Cells(i + 2, 3).value = row['CountDim']
        sheet.Cells(i + 2, 4).value = row['CountTD']
        sheet.Cells(i + 2, 5).value = row['A0']
        sheet.Cells(i + 2, 6).value = row['A1']
        sheet.Cells(i + 2, 7).value = row['A2']
        sheet.Cells(i + 2, 8).value = row['A3']
        sheet.Cells(i + 2, 9).value = row['A4']
        sheet.Cells(i + 2, 10).value = "".join(('="', row['Scale'], '"'))


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Скрываем основное окно и сразу окно выбора файлов

    filenames = askopenfilenames(title="Выберети чертежи деталей", filetypes=[('Компас 3D', '*.cdw'), ])

    print_to_excel(parse_design_documents(filenames))

    root.destroy()  # Уничтожаем основное окно
    root.mainloop()
