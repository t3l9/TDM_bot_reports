"""
lk_prefect.py — адаптирован для TDM (убраны зависимости от Telegram).
Парсинг и обработка данных ЛК Префекта (gorod.mos.ru).

Изменения в этой версии:
- Две сводные таблицы на одном листе:
    1) Регламентный срок у сообщения (Портал) × Район
    2) Дата отображения (Монитор) × Район   (даты в виде ДД.ММ.ГГГГ)
- Улучшенное оформление: стиль PivotStyleMedium9, тепловая карта по значениям,
  заголовок отчёта и подзаголовки над таблицами.
- Исправлены баги: PDF теперь сохраняется с расширением .pdf (а не .xlsx),
  убран дублирующийся фильтр по району, пустой результат возвращает None.
"""

import os
import time
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.formatting.rule import CellIsRule

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

home_dir  = os.path.expanduser("~")
directory = os.path.join(home_dir, "Downloads")

load_dotenv()
login_NG    = os.getenv("login_NG")
password_NG = os.getenv("password_NG")


# ──────────────────────────────────────────────
# Парсинг (синхронный)
# ──────────────────────────────────────────────

def parcing_data_lk_prefekta_sync(on_error=None) -> bool:
    """
    Синхронный парсинг ЛК Префекта.
    on_error(text) — опциональный callback для сообщения об ошибке.
    """
    chrome_install    = ChromeDriverManager().install()
    folder            = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")
    driver            = webdriver.Chrome(service=ChromeService(chromedriver_path))

    try:
        driver.get("https://gorod.mos.ru/api/service/auth/auth")

        username = driver.find_element(By.XPATH, '//input[@placeholder="Логин *"]')
        password = driver.find_element(By.XPATH, '//input[@placeholder="Пароль*"]')
        username.send_keys(login_NG)
        password.send_keys(password_NG)

        login_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div/div/main/div/div/div/div[2]/form[1]/button"
        )
        login_button.click()

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,
            '//div[@class="dashboard__block-link"]'
            '//div[@class="button-big link"]'
            '//div[@class="dashboard-container__links-title" and contains(text(), "Аналитика")]'
        )))

        driver.get("https://gorod.mos.ru/admin/ker/olap/report/155")
        time.sleep(10)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,
            "/html/body/div[3]/div/div[2]/div/div/div/div/form/footer/button[3]/span[2]/span"
        )))
        driver.find_element(By.XPATH,
            "/html/body/div[3]/div/div[2]/div/div/div/div/form/footer/button[3]/span[2]/span"
        ).click()
        time.sleep(1)

        driver.find_element(By.XPATH,
            "//button[contains(@class, 'bg-primary')]//span[text()='Экспорт']"
        ).click()
        time.sleep(1)

        driver.get("https://gorod.mos.ru/admin/ker/olap/downloads")
        WebDriverWait(driver, 1500).until(EC.presence_of_element_located((By.XPATH,
            "/html/body/div[1]/div/div[2]/main/div/div[1]/div/div[2]"
            "/div[1]/table/tbody/tr[1]/td[5]/div/i"
        )))
        driver.find_element(By.XPATH,
            "/html/body/div[1]/div/div[2]/main/div/div[1]/div/div[2]"
            "/div[1]/table/tbody/tr[1]/td[5]/div/i"
        ).click()
        time.sleep(20)
        return True

    except Exception as e:
        msg = f"❌ Ошибка при выгрузке ЛК Префекта: {e}"
        print(msg)
        if on_error:
            on_error(msg)
        return False
    finally:
        driver.quit()


# Async-обёртка для обратной совместимости
async def parcing_data_lk_prefekta(context=None, chat_id=None):
    import asyncio
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, parcing_data_lk_prefekta_sync)


# ──────────────────────────────────────────────
# VBA-макрос (две сводные + оформление)
# ──────────────────────────────────────────────

VBA_MACRO = r"""
Sub CreateReport()
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt1 As PivotTable, pt2 As PivotTable
    Dim lastRow As Long, lastCol As Long, startRow2 As Long

    Set wsData = ThisWorkbook.Sheets("Sheet1")

    ' Формат дат в исходных данных
    wsData.Columns("B").NumberFormat = "DD.MM.YYYY"  ' Дата отображения (Монитор)
    wsData.Columns("C").NumberFormat = "DD.MM.YYYY"  ' Регламентный срок (Портал)

    ' Пересоздаём лист сводных
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Сводная таблица").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "Сводная таблица"

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    ' Один кэш на обе сводные
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Cells(1, 1).Resize(lastRow, lastCol))

    ' Заголовок отчёта
    With wsPivot.Range("A1")
        .Value = "Отчёт ЛК Префекта ЮВАО"
        .Font.Name = "Times New Roman": .Font.Size = 14: .Font.Bold = True
    End With
    With wsPivot.Range("A2")
        .Value = "Сформировано: " & Format(Now, "DD.MM.YYYY HH:MM")
        .Font.Name = "Times New Roman": .Font.Size = 9: .Font.Italic = True
    End With
    With wsPivot.Range("A3")
        .Value = "Срок ответа по районам (Регламентный срок, Портал)"
        .Font.Name = "Times New Roman": .Font.Bold = True
    End With

    ' ===== Сводная 1: Регламентный срок x Район =====
    Set pt1 = pc.CreatePivotTable(TableDestination:=wsPivot.Cells(4, 1), TableName:="PT_Срок")
    With pt1
        .PivotFields("Район").Orientation = xlColumnField
        .PivotFields("Регламентный срок у сообщения (Портал)").Orientation = xlRowField
        .AddDataField .PivotFields("Номер заявки"), "Кол-во", xlCount
        .RowAxisLayout xlTabularRow
        .ColumnGrand = True: .RowGrand = True
    End With
    On Error Resume Next
    pt1.PivotFields("Регламентный срок у сообщения (Портал)").DataRange.Cells(1).Ungroup
    ' Формат дат задаём на самих ячейках-метках (PivotField.NumberFormat у строкового поля даёт 1004)
    pt1.PivotFields("Регламентный срок у сообщения (Портал)").DataRange.NumberFormat = "DD.MM.YYYY"
    On Error GoTo 0
    StylePivot pt1

    ' ===== Сводная 2: Дата отображения x Район =====
    startRow2 = pt1.TableRange2.Row + pt1.TableRange2.Rows.Count + 2
    With wsPivot.Cells(startRow2, 1)
        .Value = "Количество сообщений по дате отображения (Монитор)"
        .Font.Name = "Times New Roman": .Font.Size = 12: .Font.Bold = True
    End With

    Set pt2 = pc.CreatePivotTable(TableDestination:=wsPivot.Cells(startRow2 + 1, 1), TableName:="PT_Дата")
    With pt2
        .PivotFields("Район").Orientation = xlColumnField
        .PivotFields("Дата отображения (Монитор)").Orientation = xlRowField
        .AddDataField .PivotFields("Номер заявки"), "Кол-во", xlCount
        .RowAxisLayout xlTabularRow
        .ColumnGrand = True: .RowGrand = True
    End With
    ' Разгруппировать даты -> вид ДД.ММ.ГГГГ
    On Error Resume Next
    pt2.PivotFields("Дата отображения (Монитор)").DataRange.Cells(1).Ungroup
    On Error GoTo 0
    ' Скрыть пустые значения "(пусто)" в таблице с монитором
    HideBlankItems pt2, "Дата отображения (Монитор)"
    HideBlankItems pt2, "Район"
    On Error Resume Next
    ' Формат дат задаём на самих ячейках-метках (PivotField.NumberFormat у строкового поля даёт 1004)
    pt2.PivotFields("Дата отображения (Монитор)").DataRange.NumberFormat = "DD.MM.YYYY"
    ' Хронология: от ранней даты к поздней (по возрастанию)
    pt2.PivotFields("Дата отображения (Монитор)").AutoSort xlAscending, "Дата отображения (Монитор)"
    On Error GoTo 0
    StylePivot pt2

    ' Столбец A держит подписи полей строк ("Регламентный срок у сообщения (Портал)",
    ' "Дата отображения (Монитор)") — делаем его широким, чтобы названия помещались.
    wsPivot.Columns("A").ColumnWidth = 32

    With wsPivot.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1: .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
    End With
End Sub

Private Sub StylePivot(pt As PivotTable)
    pt.TableStyle2 = "PivotStyleLight16"
    With pt.TableRange2
        .Font.Name = "Times New Roman": .Font.Size = 10
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        ' Границы у всех ячеек таблицы (внешние + внутренние)
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(140, 140, 140)
        End With
    End With

    Dim body As Range, heat As Range
    On Error Resume Next
    Set body = pt.DataBodyRange
    On Error GoTo 0
    If Not body Is Nothing Then
        body.FormatConditions.Delete

        ' Исключаем строку и столбец общего итога из заливки.
        ' DataBodyRange включает итоги: последняя строка = "Общий итог" (строки),
        ' последний столбец = "Общий итог" (столбцы). Обрезаем их.
        Set heat = body
        If pt.ColumnGrand And heat.Rows.Count > 1 Then
            Set heat = heat.Resize(heat.Rows.Count - 1, heat.Columns.Count)
        End If
        If pt.RowGrand And heat.Columns.Count > 1 Then
            Set heat = heat.Resize(heat.Rows.Count, heat.Columns.Count - 1)
        End If

        ' Адаптивный градиент: минимум -> жёлтый, максимум -> красный.
        ' LowestValue/HighestValue пересчитываются от реальных чисел (без итогов),
        ' поэтому диапазон 1..10 и 1..70 красятся одинаково корректно.
        heat.FormatConditions.AddColorScale ColorScaleType:=2
        With heat.FormatConditions(heat.FormatConditions.Count)
            .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 243, 178)  ' пастельно-жёлтый (минимум)
            .ColorScaleCriteria(2).Type = xlConditionValueHighestValue
            .ColorScaleCriteria(2).FormatColor.Color = RGB(244, 169, 160)  ' пастельно-красный (максимум)
        End With
    End If
    ' Автоширину применяем к столбцам данных, но НЕ к A —
    ' столбец A держит длинные подписи полей и задаётся вручную.
    On Error Resume Next
    Intersect(pt.TableRange2, pt.TableRange2.Offset(0, 1)).EntireColumn.AutoFit
    On Error GoTo 0
End Sub

Private Sub HideBlankItems(pt As PivotTable, fieldName As String)
    Dim pi As PivotItem
    On Error Resume Next
    For Each pi In pt.PivotFields(fieldName).PivotItems
        If Trim(pi.Caption) = "" Or pi.Caption = "(пусто)" Or LCase(pi.Caption) = "(blank)" Then
            pi.Visible = False
        End If
    Next pi
    On Error GoTo 0
End Sub
"""


# ──────────────────────────────────────────────
# Обработка файла
# ──────────────────────────────────────────────

def process_lk_prefekta_file(directory: str, selected_district: str, filepath: str):
    df = pd.read_excel(filepath)

    responsible_mapping = {
        'ГБУ «Автомобильные дороги ЮВАО»': 'АВД ЮВАО',
        'ГБУ Жилищник Выхино района Выхино-Жулебино города Москвы': 'Выхино-Жулебино',
        'Управа Выхино-Жулебино': 'Выхино-Жулебино',
        'ГБУ Жилищник Нижегородского района города Москвы': 'Нижегородский',
        'Управа Нижегородский': 'Нижегородский',
        'ГБУ Жилищник района Капотня города Москвы': 'Капотня',
        'Управа Капотня': 'Капотня',
        'ГБУ Жилищник района Кузьминки города Москвы': 'Кузьминки',
        'Управа Кузьминки': 'Кузьминки',
        'ГБУ Жилищник района Лефортово города Москвы': 'Лефортово',
        'Управа Лефортово': 'Лефортово',
        'ГБУ Жилищник района Люблино города Москвы': 'Люблино',
        'Управа Люблино': 'Люблино',
        'ГБУ Жилищник района Марьино города Москвы': 'Марьино',
        'Управа Марьино': 'Марьино',
        'ГБУ Жилищник района Некрасовка города Москвы': 'Некрасовка',
        'Управа Некрасовка': 'Некрасовка',
        'ГБУ Жилищник района Печатники города Москвы': 'Печатники',
        'Управа Печатники': 'Печатники',
        'ГБУ Жилищник района Текстильщики города Москвы': 'Текстильщики',
        'Управа Текстильщики': 'Текстильщики',
        'ГБУ Жилищник Рязанского района города Москвы': 'Рязанский',
        'Управа Рязанский': 'Рязанский',
        'ГБУ Жилищник Южнопортового района города Москвы': 'Южнопортовый',
        'Управа Южнопортовый': 'Южнопортовый'
    }

    # Функция для обновления значений в столбце 'Район'
    def update_region(row):
        if row['Ответственный ОИВ первого уровня'] == 'Префектура Юго-Восточного округа':
            return row['Район']  # Ничего не меняем
        else:
            return responsible_mapping.get(row['Ответственный ОИВ первого уровня'], row['Район'])

    # Применение функции к каждому ряду
    df['Район'] = df.apply(update_region, axis=1)

    df_filtered = df[df['Ответственный за подготовку ответа'] == 'Префектура Юго-Восточного округа']

    columns_to_keep = [
        "Номер заявки",
        "Дата отображения (Монитор)",
        "Регламентный срок у сообщения (Портал)",
        "Признак Монитора",
        "Просрок (Монитор)",
        "Дата публикации сообщения",
        "Район",
        "Проблемная тема",
        "Адрес",
        "Категория объекта",
        "Категория/действие последнего ответа",
        "Ответственный за подготовку ответа",
        "Ответственный ОИВ первого уровня",
        "Статус подготовки ответа на сообщение"
    ]

    # Проверяем, какие столбцы из списка реально существуют
    existing_columns = [col for col in columns_to_keep if col in df_filtered.columns]
    missing_columns = [col for col in columns_to_keep if col not in df_filtered.columns]

    if missing_columns:
        print(f"Внимание! Отсутствуют следующие столбцы: {missing_columns}")
        print("Доступные столбцы:", df_filtered.columns.tolist())

    # Используем только существующие столбцы
    df_filtered = df_filtered[existing_columns]

    if selected_district != "Все районы" and 'Район' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Район'] == selected_district]

    df_filtered = df_filtered.dropna(how='all')
    if df_filtered.empty:
        print("После фильтрации не осталось данных.")
        return None

    now = pd.Timestamp.now()
    base_name = f"{selected_district}_ЛК_Префекта_{datetime.now().strftime('%d.%m')}_на_{now.strftime('%H-%M')}"
    processed_file_path = os.path.join(directory, f"{base_name}.xlsx")
    print(f"Saving processed file to: {processed_file_path}")
    df_filtered.to_excel(processed_file_path, index=False)
    excel_file = processed_file_path

    if not WIN32COM_AVAILABLE:
        print("win32com недоступен — пропускаю создание сводных и PDF.")
        return processed_file_path

    # Запускаем Excel
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = True  # Если нужно скрыть Excel — поставьте False

    # Открываем Excel-файл
    workbook = excel.Workbooks.Open(excel_file)

    # Добавляем новый модуль VBA и вставляем макрос
    vb_module = workbook.VBProject.VBComponents.Add(1)  # 1 = стандартный модуль
    vb_module.CodeModule.AddFromString(VBA_MACRO)

    # Выполняем макрос
    excel.Application.Run("CreateReport")
    print("Сводные таблицы созданы")

    # --- Экспорт первого листа в PDF (.pdf, а не .xlsx!) ---
    pdf_file_name = f"{base_name}.pdf"
    pdf_path = os.path.join(os.path.dirname(processed_file_path), pdf_file_name)
    wsFirst = workbook.Worksheets(1)  # Ссылка на первый лист

    # Настройки страницы для печати
    wsFirst.PageSetup.FitToPagesWide = 1
    wsFirst.PageSetup.FitToPagesTall = 1
    wsFirst.PageSetup.Zoom = False

    # Уменьшаем поля
    wsFirst.PageSetup.LeftMargin   = excel.Application.CentimetersToPoints(0.5)
    wsFirst.PageSetup.RightMargin  = excel.Application.CentimetersToPoints(0.5)
    wsFirst.PageSetup.TopMargin    = excel.Application.CentimetersToPoints(0.5)
    wsFirst.PageSetup.BottomMargin = excel.Application.CentimetersToPoints(0.5)
    workbook.Save()

    try:
        if os.path.exists(pdf_path):
            print(f"Файл {pdf_path} существует. Удаление...")
            os.remove(pdf_path)
            print("Файл успешно удалён.")

        print(f"Сохранение PDF в {pdf_path}...")
        wsFirst.ExportAsFixedFormat(0, pdf_path)  # 0 = xlTypePDF
        print(f"PDF успешно создан: {pdf_path}")
    except Exception as e:
        print(f"Ошибка при сохранении PDF: {e}")

    # Автоширина для листа со сводными
    try:
        sheet = workbook.Worksheets(2)
        sheet.Cells.EntireColumn.AutoFit()
    except Exception as e:
        print(f"Не удалось применить автоширину: {e}")

    # Сохраняем и закрываем
    workbook.Save()
    workbook.Close()
    excel.Quit()

    return processed_file_path