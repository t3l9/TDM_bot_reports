"""
mji_svod.py — адаптирован для TDM.

ИЗМЕНЕНО:
  - parcing_MWI() — убран Telegram-контекст, добавлен on_error callback
  - async-обёртка сохранена для обратной совместимости

БЕЗ ИЗМЕНЕНИЙ (скопируй из оригинала):
  - MWI_choosing_files()
  - MWI_process_file()
  - create_pivot_and_pdf()
"""

import os
import time
from datetime import datetime, timedelta

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
import win32com.client
from functools import reduce
import pythoncom

home_dir  = os.path.expanduser("~")
directory = os.path.join(home_dir, "Downloads")

load_dotenv()
login_NG    = os.getenv("login_NG")
password_NG = os.getenv("password_NG")


# ──────────────────────────────────────────────
# ИЗМЕНЁННАЯ ФУНКЦИЯ: parcing_MWI
# ──────────────────────────────────────────────

def parcing_MWI_sync(on_error=None) -> int:
    """
    Синхронный парсинг СВОДА МЖИ.
    Возвращает количество успешно выгруженных вкладок (0–3).
    on_error(text) — опциональный callback для уведомления об ошибке.
    """
    chrome_install    = ChromeDriverManager().install()
    folder            = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")
    driver            = webdriver.Chrome(service=ChromeService(chromedriver_path))

    processed_tabs = [False, False, False]

    try:
        driver.get("https://gorod.mos.ru/api/service/auth/auth")
        driver.maximize_window()

        driver.find_element(By.XPATH, '//input[@placeholder="Логин *"]').send_keys(login_NG)
        driver.find_element(By.XPATH, '//input[@placeholder="Пароль*"]').send_keys(password_NG)
        driver.find_element(
            By.XPATH, "/html/body/div[1]/div/div/main/div/div/div/div[2]/form[1]/button"
        ).click()

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
            '//div[@class="dashboard__block-link"]'
            '//div[@class="button-big link"]'
            '//div[@class="dashboard-container__links-title" and contains(text(), "Аналитика")]'
        )))

        # Закрываем уведомление если есть
        try:
            btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div/div/div[4]/div/div/div[2]/div[2]/div/div/button")
            ))
            btn.click()
            time.sleep(1)
        except Exception:
            pass

        time.sleep(0.5)
        driver.get("https://er.mos.ru/ker/admin/issues/monitor_mzi?sidebar=organization")
        time.sleep(4)

        # ── Вкладка 1: Ответы на доработке ──────────────────────────────
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class,'v-list-item__title') and text()='Ответы на доработке']")
            )).click()
            time.sleep(5)
            _download_mji_tab(driver)
            processed_tabs[0] = True
            driver.get("https://er.mos.ru/ker/admin/issues/monitor_mzi?sidebar=organization")
            time.sleep(3)
        except Exception as e:
            print(f"❌ Вкладка 1 недоступна: {e}")
            driver.get("https://er.mos.ru/ker/admin/issues/monitor_mzi?sidebar=organization")
            time.sleep(3)

        # ── Вкладка 2: Обещание устранения ──────────────────────────────
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class,'v-list-item__title') and text()='Обещание устранения']")
            )).click()
            time.sleep(5)
            _download_mji_tab(driver)
            processed_tabs[1] = True
            driver.get("https://er.mos.ru/ker/admin/issues/monitor_mzi?sidebar=organization")
            time.sleep(3)
        except Exception as e:
            print(f"❌ Вкладка 2 недоступна: {e}")
            driver.get("https://er.mos.ru/ker/admin/issues/monitor_mzi?sidebar=organization")
            time.sleep(3)

        # ── Вкладка 3: Нарушения для получателя ─────────────────────────
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class,'v-list-item__title') and text()='Нарушения для получателя']")
            )).click()
            time.sleep(3)
            _download_mji_tab(driver)
            processed_tabs[2] = True
        except Exception as e:
            print(f"❌ Вкладка 3 недоступна: {e}")

        # ── Скачивание файлов ────────────────────────────────────────────
        driver.get("https://gorod.mos.ru/admin/ker/olap/downloads")

        row_xpaths = [
            ("/html/body/div[1]/div/div[2]/main/div/div[1]/div/div[2]/div[1]/table/tbody/tr[3]/td[5]/div/i", 0),
            ("/html/body/div[1]/div/div[2]/main/div/div[1]/div/div[2]/div[1]/table/tbody/tr[2]/td[5]/div/i", 1),
            ("/html/body/div[1]/div/div[2]/main/div/div[1]/div/div[2]/div[1]/table/tbody/tr[1]/td[5]/div/i", 2),
        ]
        for xpath, idx in row_xpaths:
            if processed_tabs[idx]:
                try:
                    WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, xpath)))
                    driver.find_element(By.XPATH, xpath).click()
                    time.sleep(1.5)
                except Exception as e:
                    print(f"❌ Ошибка скачивания вкладки {idx+1}: {e}")

        return sum(processed_tabs)

    except Exception as e:
        msg = f"❌ Ошибка при выгрузке СВОДА МЖИ: {e}"
        print(msg)
        if on_error:
            on_error(msg)
        return sum(processed_tabs)
    finally:
        driver.quit()


def _download_mji_tab(driver):
    """Вспомогательная функция — нажать кнопку скачать и выбрать 'Все сообщения'."""
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "//button[.//i[contains(@class,'mdi-download')]]")
    )).click()
    time.sleep(2)

    # Скрываем оверлеи
    try:
        driver.execute_script("document.querySelector('.v-overlay__scrim.white').style.display='none';")
        time.sleep(1)
        driver.execute_script("document.querySelector('.v-overlay.v-overlay--active').style.display='none';")
        time.sleep(1)
    except Exception:
        pass

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, '(//div[@class="v-select__selections"])[4]')
    )).click()
    time.sleep(2)

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "//div[contains(text(),'Все сообщения')]")
    )).click()
    time.sleep(2)

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(span,'Экспорт')]")
    )).click()
    time.sleep(3)


# Async-обёртка — сохранена для обратной совместимости
async def parcing_MWI(context=None, chat_id=None):
    import asyncio
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, parcing_MWI_sync)


def MWI_choosing_files(directory, processed_tabs_count):
    """
    Обрабатывает файлы в зависимости от количества успешно обработанных вкладок

    :param directory: Директория с загруженными файлами
    :param processed_tabs_count: Количество успешно обработанных вкладок (0-3)
    :return: Объединенный DataFrame или None, если файлов нет
    """
    # Получение списка файлов с сортировкой по дате изменения (от последнего к более старым)
    files = sorted([os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.xlsx')],
                   key=os.path.getmtime, reverse=True)

    # Выбираем нужное количество последних файлов
    if processed_tabs_count == 0:
        print("Нет файлов для обработки - ни одна вкладка не была выгружена")
        return None
    elif processed_tabs_count == 1:
        latest_files = files[:1]
        print("Обрабатываем 1 файл:", latest_files)
    elif processed_tabs_count == 2:
        latest_files = files[:2]
        print("Обрабатываем 2 файла:", latest_files)
    else:  # 3 или больше
        latest_files = files[:3]
        print("Обрабатываем 3 файла:", latest_files)

    try:
        # Загрузка таблиц только для существующих файлов
        dfs = []
        for file in latest_files:
            if os.path.isfile(file):
                df = pd.read_excel(file)
                dfs.append(df)

        if not dfs:
            print("Не найдено ни одного файла для обработки")
            return None

        # Объединение DataFrame
        result_df = pd.concat(dfs, ignore_index=True)
        return result_df

    except Exception as e:
        print(f"Ошибка при обработке файлов: {str(e)}")
        return None


def MWI_process_file(df):
    # Фильтрация строк, где "Ответственный ОИВ 1 уровня" содержит "ГБУ"
    # df = df[df['Ответственный ОИВ 1 уровня'].str.contains('ГБУ', na=False)]
    today = datetime.now()
    tomorrow = today + timedelta(days=1)
    df['Тип'] = ''
    df['Дата отображения на мониторе'] = pd.to_datetime(df['Дата отображения на мониторе'], dayfirst=True,
                                                        format='%d.%m.%Y %H:%M:%S')
    df.loc[(df['Дата отображения на мониторе'].dt.date == today.date()) & (
            df['Просрок Монитора'] == 'Нет'), "Тип"] = "Срок сегодня"
    df.loc[(df['Просрок Монитора'] == 'Да'), "Тип"] = "Просрок"
    df.loc[(df['Дата отображения на мониторе'].dt.date == tomorrow.date()), "Тип"] = "Срок завтра"
    # Условие для послезавтра и далее
    df.loc[(df['Дата отображения на мониторе'].dt.date > tomorrow.date()), "Тип"] = "Послезавтра и далее"
    return df


def create_pivot_and_pdf(excel_file, directory):
    """
    Создает сводную таблицу и PDF из Excel файла МЖИ

    Args:
        excel_file: Путь к Excel файлу с данными
        directory: Директория для сохранения результатов

    Returns:
        tuple: (pdf_path, success) - путь к PDF файлу и статус успешности
    """
    try:
        import win32com.client
        WIN32COM_AVAILABLE = True
    except ImportError:
        return None, False, "Модуль win32com не установлен"

    try:
        today = datetime.now()
        timenow = today.strftime("%H-%M")

        vba_macro = """  
        Sub CreatePivotTable()  
            Dim wsData As Worksheet  
            Dim wsPivot As Worksheet  
            Dim pivotCache As PivotCache  
            Dim pivotTable As PivotTable  
            Dim lastRow As Long  
            Dim lastCol As Long  

            ' Укажите лист с данными  
            Set wsData = ThisWorkbook.Sheets("МЖИ") ' Замените на имя вашего листа с данными  

            With wsData.Columns("B")
                .NumberFormat = "DD.MM.YYYY"
            End With

            ' Создаем новый лист для сводной таблицы  
            On Error Resume Next  
            Application.DisplayAlerts = False  
            ThisWorkbook.Sheets("Сводная таблица").Delete ' Удаляем лист, если уже существует  
            Application.DisplayAlerts = True  
            On Error GoTo 0  
            Set wsPivot = ThisWorkbook.Sheets.Add  
            wsPivot.Name = "Сводная таблица"  

            ' Находим последний заполненный ряд и столбец на листе с данными  
            lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row  
            lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column  

            ' Создаем кэш для сводной таблиции  
            Set pivotCache = ThisWorkbook.PivotCaches.Create( _  
                SourceType:=xlDatabase, _  
                SourceData:=wsData.Cells(1, 1).Resize(lastRow, lastCol))  

            ' Создаем сводную таблицу  
            Set pivotTable = pivotCache.CreatePivotTable( _  
                TableDestination:=wsPivot.Cells(3, 1), _  
                TableName:="MyPivotTable")  

            With pivotTable  
                .PivotFields("Район").Orientation = xlRowField  
                .PivotFields("Тип").Orientation = xlColumnField  
                .AddDataField .PivotFields("Номер заявки"), "Количество", xlCount  
            End With  

            wsPivot.Range("A4").Value = "Район" 
            ' Скрываем первую строку  
            wsPivot.Rows(3).Hidden = True

            Dim typePivotField As PivotField
            Dim item As PivotItem
            Set typePivotField = pivotTable.PivotFields("Тип")
            For Each item In typePivotField.PivotItems
                ' Если элемент пустой, скрываем его
                If item.Name = "(blank)" Then
                    item.Visible = False
                End If
            Next item

            ' Обновляем сводную таблицу  
            pivotTable.RefreshTable  

            ' Форматирование сводной таблицы  
            Dim rng As Range  
            Set rng = wsPivot.Range("A4").CurrentRegion  
            With rng  
                .Font.Name = "Times New Roman"  
                .Font.Size = 14  
                .Font.Bold = True  
                .Borders.LineStyle = xlContinuous  
                .WrapText = True ' Перенос текста  
                .HorizontalAlignment = xlCenter ' Выравнивание по центру  
            End With  
            wsPivot.Columns("A").ColumnWidth = 24 ' Установите желаемую ширину столбца  
            wsPivot.Rows(6).RowHeight = 19 ' Установите высоту 6-й строки

        End Sub
        """

        # Запускаем Excel
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False  # Скрываем окно Excel

        # Открываем Excel-файл
        workbook = excel.Workbooks.Open(excel_file)

        # Добавляем новый модуль VBA и вставляем макрос
        vb_module = workbook.VBProject.VBComponents.Add(1)  # 1 = стандартный модуль
        vb_module.CodeModule.AddFromString(vba_macro)

        # Выполняем макрос
        excel.Application.Run("CreatePivotTable")

        # Создаем путь для PDF
        pdf_file_name = f"СВОД МЖИ {datetime.now().strftime('%d.%m.%y')} на {timenow}.pdf"
        pdf_path = os.path.join(directory, pdf_file_name)

        wsFirst = workbook.Worksheets(1)  # Ссылка на первый лист

        # Настройки страницы для печати
        wsFirst.PageSetup.FitToPagesWide = 1  # Устанавливаем количество страниц по ширине
        wsFirst.PageSetup.FitToPagesTall = 1  # Устанавливаем количество страниц по высоте на 1
        wsFirst.PageSetup.Zoom = False  # Отключаем масштабирование

        # Обновляем отступы страницы для уменьшения размера PDF
        wsFirst.PageSetup.LeftMargin = excel.Application.CentimetersToPoints(0.5)
        wsFirst.PageSetup.RightMargin = excel.Application.CentimetersToPoints(0.5)
        wsFirst.PageSetup.TopMargin = excel.Application.CentimetersToPoints(0.5)
        wsFirst.PageSetup.BottomMargin = excel.Application.CentimetersToPoints(0.5)

        # Сохраняем изменения
        workbook.Save()

        # Убираем ошибку, если файл уже существует
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

        # Экспортируем в PDF
        wsFirst.ExportAsFixedFormat(0, pdf_path)  # 0 - это xlTypePDF

        # Автоподбор ширины колонок на втором листе
        if workbook.Worksheets.Count > 1:
            sheet = workbook.Worksheets(2)
            sheet.Cells.EntireColumn.AutoFit()
            workbook.Save()

        # Закрываем файл и Excel
        workbook.Close()
        excel.Quit()

        return pdf_path, True, "PDF успешно создан"

    except Exception as e:
        error_msg = f"Ошибка при создании PDF: {str(e)}"
        print(error_msg)
        return None, False, error_msg
