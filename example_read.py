# -*- coding: utf-8 -*-
# |fragment_translate

import pythoncom
from win32com.client import Dispatch, gencache


def KompasAPI():  # подключение API компаса
    global KompasAPI7  # значение делаем глобальным
    global iApplication  # значение делаем глобальным

    KompasAPI7 = gencache.EnsureModule(
        '{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0)  # API7 КОМПАСа
    # интерфейс приложения КОМПАС-3D.
    iApplication = Dispatch('Kompas.Application.7')


def Read_text(iSelectedObject):  # прочитаем выделеный текст
    if iSelectedObject:  # если выделено
        if iSelectedObject.DrawingObjectType == 4:  # если это текст
            iDrawingText = iDrawingTexts.DrawingText(
                iSelectedObject.Reference)  # интерфейс текста на чертеже
            iText = KompasAPI7.IText(iDrawingText)  # интерфейс текста
            print('Текст: '+str(iText.Str))
            return iText.Str
    return None


def Read_all_texts():  # прочитаем все тексты в документе
    global iDrawingTexts

    print("=== ПОИСК ВСЕХ ТЕКСТОВ В ДОКУМЕНТЕ ===")

    # Получаем все текстовые объекты
    if iDrawingTexts and iDrawingTexts.Count > 0:
        print("Найдено текстовых объектов: {}".format(iDrawingTexts.Count))

        for i in range(iDrawingTexts.Count):
            try:
                iDrawingText = iDrawingTexts.DrawingText(i)
                iText = KompasAPI7.IText(iDrawingText)
                text_content = iText.Str
                if text_content and text_content.strip():
                    print("Текст [{}]: '{}'".format(i, text_content))
            except Exception as e:
                print("Ошибка при чтении текста [{}]: {}".format(i, e))
    else:
        print("Текстовые объекты не найдены")


def Read_tables():  # прочитаем все таблицы в документе
    print("\n=== ПОИСК ТАБЛИЦ В ДОКУМЕНТЕ ===")

    try:
        # Получаем контейнер символов 2D
        iSymbols2DContainer = KompasAPI7.ISymbols2DContainer(iView)
        if not iSymbols2DContainer:
            print("Контейнер символов 2D не найден")
            return

        # Получаем таблицы
        iDrawingTables = iSymbols2DContainer.DrawingTables
        if not iDrawingTables or iDrawingTables.Count == 0:
            print("Таблицы не найдены")
            return

        print("Найдено таблиц: {}".format(iDrawingTables.Count))

        # Обрабатываем каждую таблицу
        for table_idx in range(iDrawingTables.Count):
            print("\n--- Таблица {} ---".format(table_idx + 1))

            try:
                iDrawingTable = iDrawingTables.DrawingTable(table_idx)
                iTable = KompasAPI7.ITable(iDrawingTable)

                print("Строк: {}, Столбцов: {}".format(
                    iTable.RowsCount, iTable.ColumnsCount))

                # Способ 1: Через Range
                try:
                    table_range = iTable.Range(
                        0, 0, iTable.RowsCount, iTable.ColumnsCount)
                    print("Текст через Range:")
                    print(table_range.Texts)
                except Exception as e:
                    print("Ошибка при получении Range: {}".format(e))

                # Способ 2: Через ячейки по отдельности
                print("\nТекст по ячейкам:")
                for row in range(iTable.RowsCount):
                    for col in range(iTable.ColumnsCount):
                        try:
                            # Получаем ячейку по ID
                            cell_id = row * iTable.ColumnsCount + col + 1
                            iTableCell = iTable.CellById(cell_id)
                            iText = KompasAPI7.IText(iTableCell.Text)
                            text_cell = iText.Str
                            if text_cell and text_cell.strip():
                                print("[{},{}] ID={}: '{}'".format(
                                    row, col, cell_id, text_cell))
                        except Exception as e:
                            print("[{},{}] Ошибка: {}".format(row, col, e))

                # Способ 3: Через Text(row, col)
                print("\nТекст через Text(row, col):")
                for row in range(iTable.RowsCount):
                    for col in range(iTable.ColumnsCount):
                        try:
                            iText = iTable.Text(row, col)
                            if hasattr(iText, 'Str'):
                                text_cell = iText.Str
                                if text_cell and text_cell.strip():
                                    print("[{},{}]: '{}'".format(
                                        row, col, text_cell))
                        except Exception as e:
                            print("[{},{}] Ошибка: {}".format(row, col, e))

            except Exception as e:
                print("Ошибка при обработке таблицы {}: {}".format(table_idx, e))

    except Exception as e:
        print("Ошибка при поиске таблиц: {}".format(e))


# Основной код
KompasAPI()  # подключение API компаса

# получить текущий активный документ
iKompasDocument = iApplication.ActiveDocument

iKompasDocument2D = KompasAPI7.IKompasDocument2D(
    iKompasDocument)  # базовый класс графических документов КОМПАС

# менеджер слоев и видов графического документа
iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager
# интерфейс коллекции видов графического документа
iViews = iViewsAndLayersManager.Views
iView = iViews.ActiveView  # интерфейс вида графического документа

# интерфейс контейнера объектов вида графического документа
iDrawingContainer = KompasAPI7.IDrawingContainer(iView)
# указатель на интерфейс коллекции текстов на чертеже
iDrawingTexts = iDrawingContainer.DrawingTexts

iKompasDocument2D1 = KompasAPI7.IKompasDocument2D1(
    iKompasDocument)  # дополнительный интерфейс IKompasDocument2D

# Читаем все тексты автоматически
Read_all_texts()

# Читаем все таблицы автоматически
Read_tables()

print("\n=== ВЫГРУЗКА ЗАВЕРШЕНА ===")
