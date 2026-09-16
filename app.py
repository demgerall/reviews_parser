# Инициализация дизайна
# pyuic6 C:/Users/demge/mainWindow.ui -o C:/Users/demge/PycharmProjects/TalkingBot/designMain.py

# Компилятор exe
# pyinstaller -F -w -i "C:\Users\demge\PycharmProjects\ReviewsParser\dozer.ico" app.py

import datetime
import json
import random
import time
import re
import os
import sys
import logging

from threading import Thread

from PyQt6 import QtWidgets

from selenium import webdriver as wd
from selenium.common import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import (
    PatternFill, Border, Side,
    Alignment, Font
)
from selenium.webdriver.support.wait import WebDriverWait

import designMain

def set_styles_to_sheet(sheet: Worksheet, num: int) -> None:
    sheet.auto_filter.ref = f"A1:G{num - 1}"

    colors = ["DE0F10", "F48A11", "F49407", "ABBF1B", "015423"]
    cols = ["A", "B", "C", "D", "E", "F", "G"]

    for col in cols:
        for row in range(1, num):
            sheet[f"{col}{row}"].alignment = Alignment(wrap_text=True, vertical="top")
            sheet[f"{col}{row}"].font = Font(name="Calibri", size=10)
            sheet[f"{col}{row}"].border = Border(left=Side(border_style="thin", color='000000'),
                                                 right=Side(border_style="thin", color='000000'),
                                                 top=Side(border_style="thin", color='000000'),
                                                 bottom=Side(border_style="thin", color='000000'), )
            if row == 1:
                sheet[f"{col}{row}"].alignment = Alignment(horizontal='center')
                sheet[f"{col}{row}"].font = Font(name="Calibri", size=12, bold=True)
                sheet[f"{col}{row}"].fill = PatternFill(patternType='solid', fgColor="FABF8F")
                sheet[f"{col}{row}"].border = Border(left=Side(border_style="medium", color='000000'),
                                                     right=Side(border_style="medium", color='000000'),
                                                     top=Side(border_style="medium", color='000000'),
                                                     bottom=Side(border_style="medium", color='000000'), )
            if col == "A" and row != 1:
                match sheet.cell(row=row, column=1).value:
                    case "+":
                        sheet[f"{col}{row}"].fill = PatternFill(patternType='solid', fgColor="ABBF1B")
                    case "-":
                        sheet[f"{col}{row}"].fill = PatternFill(patternType='solid', fgColor="DE0F10")
            if col == "D" and row != 1:
                sheet[f"{col}{row}"].alignment = Alignment(horizontal='center', vertical="top")
                sheet[f"{col}{row}"].font = Font(color=colors[int(sheet.cell(row=row, column=4).value) - 1])
            if sheet.cell(row=row, column=cols.index(col) + 1).value == "None":
                sheet[f"{col}{row}"].font = Font(name="Calibri", size=12, bold=True, color="A94123")

    sheet.column_dimensions["A"].width = 10
    sheet.column_dimensions["B"].width = 25
    sheet.column_dimensions["C"].width = 20
    sheet.column_dimensions["D"].width = 12
    sheet.column_dimensions["E"].width = 150
    sheet.column_dimensions["F"].width = 20
    sheet.column_dimensions["G"].width = 100


def check_exists(el: WebElement, path: str) -> bool:
    try:
        el.find_element(By.CSS_SELECTOR, path)
        return True
    except NoSuchElementException:
        return False


def human_like_scroll(driver: WebDriver, element, scrolls=1):
    """Имитация естественной прокрутки с задержками и колебаниями"""
    import random
    actions = ActionChains(driver)

    for _ in range(scrolls):
        # Случайное движение мыши перед скроллом
        offset_x = random.randint(-50, 50)
        offset_y = random.randint(-50, 50)
        try:
            actions.move_by_offset(offset_x, offset_y).perform()
        except:
            pass

        # Прокрутка на случайное расстояние
        scroll_amount = random.randint(300, 800)
        driver.execute_script(f"arguments[0].scrollTop += {scroll_amount};", element)

        # Случайная пауза
        time.sleep(random.uniform(1.5, 3.5))


def random_delay(min_sec: float = 1.0, max_sec: float = 3.0):
    """Случайная задержка для имитации человеческого поведения"""
    time.sleep(random.uniform(min_sec, max_sec))


def create_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")

    driver = wd.Chrome()
    driver.maximize_window()

    return driver


class App(QtWidgets.QMainWindow, designMain.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.base_save_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        self.config = self.load_config('config.json')
        self.waitTime = self.load_waitTime()

        logging.basicConfig(level=logging.DEBUG, filename="logs.log",
                            format="%(levelname)s (%(asctime)s): %(message)s (Line: %(lineno)d) [%(filename)s]",
                            datefmt="%d/%m/%Y %I:%M:%S", encoding='UTF-8', filemode="a")

        self.save_textEdit.setPlaceholderText(
            f"Значение по умолчанию: {self.base_save_path}")
        self.filename_textEdit.setPlaceholderText(
            f"Значение по умолчанию: Отзывы <компания> <дата-время>.xlsx")

        self.start_button.clicked.connect(self.start)

    def load_config(self, config_name: str) -> object:
        try:
            with open(config_name, 'r', encoding='utf-8') as f:
                config = json.load(f)
                self.status_label.setText("--Загрузка конфига прошла успешно--")
                return config
        except Exception as _ex:
            self.status_label.setText("--Загрузка конфига не прошла успешно--")
            self.error_label.setText("Возникла ошибка. Проверьте файл logs.log")
            logging.exception(_ex)

    def load_waitTime(self) -> float | None:
        try:
            with open('waitTime.json', 'r', encoding='utf-8') as f:
                waitTime = json.load(f)
                self.status_label.setText("--Загрузка времени ожидания прошла успешно--")
                return float(waitTime['wait'])
        except Exception as _ex:
            self.status_label.setText("--Загрузка времени ожидания не прошла успешно--")
            self.error_label.setText("Возникла ошибка. Проверьте файл logs.log")
            logging.exception(_ex)

    def create_excel_book(self) -> Workbook | None:
        try:
            book = openpyxl.Workbook()
            book.remove(book.active)
            return book
        except Exception as _ex:
            self.error_label.setText("Возникла ошибка. Проверьте файл logs.log")
            logging.exception(_ex)

    def save_excel_book(self, book: Workbook, path: str, excel_file_name: str) -> None:
        try:
            book.save(os.path.join(path, excel_file_name))
            self.status_label.setText(f"--Данные сохранены в {excel_file_name} по пути: {path}--")
        except Exception as _ex:
            self.error_label.setText("Возникла ошибка. Проверьте файл logs.log")
            logging.exception(_ex)
            self.status_label.setText(f"--Данные не удалось сохранить--")

    def start(self) -> None:
        self.error_label.setText("")

        with open("logs.log", "w", encoding='UTF-8') as f:
            f.write("")
        f.close()

        thread = Thread(target=self.company_start_parsing, daemon=True)
        thread.start()

    def company_start_parsing(self) -> None:

        if self.save_textEdit.text() == "":
            path = self.base_save_path
        else:
            path = self.save_textEdit.text()

        self.start_parsing(self.config[0], path)
        random_delay(8.0, 15.0)
        self.start_parsing(self.config[1], path)
        random_delay(8.0, 15.0)
        self.start_parsing(self.config[2], path)

        # if self.dreamjob_checkBox.isChecked() and company.get('2GIS'):
            

    def start_parsing(self, company: object, path: str) -> None:
        book = self.create_excel_book()

        if self.filename_textEdit.text() == "":
            excel_file_name = f"Отзывы {company['name']} {datetime.datetime.now().strftime('%d-%b-%Y %H;%M;%S')}.xlsx"
        else:
            excel_file_name = self.filename_textEdit.text() + ".xlsx"

        driver = None
        try:
            driver = create_driver()
            driver.maximize_window()

            if self.gis_checkBox.isChecked() and company.get('2GIS'):
                self.search_on_site(driver, company, book, "2GIS")
            if self.yandex_checkBox.isChecked() and company.get('Yandex'):
                self.search_on_site(driver, company, book, "Yandex")

            self.save_excel_book(book=book, path=path, excel_file_name=excel_file_name)
        except Exception as _ex:
            self.error_label.setText("Возникла ошибка при работе с браузером. Проверьте файл logs.log")
            logging.exception(_ex)
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass

        self.filename_textEdit.clear()
        self.save_textEdit.clear()

    def search_on_site(self, driver: WebDriver, company: object, book: Workbook, site: str) -> None:
        try:
            sheet = book.create_sheet(site)

            sheet.cell(row=1, column=1).value = "Ответ"
            sheet.cell(row=1, column=2).value = "Никнейм"
            sheet.cell(row=1, column=3).value = "Дата"
            sheet.cell(row=1, column=4).value = "Оценка"
            sheet.cell(row=1, column=5).value = "Текст отзыва"
            sheet.cell(row=1, column=6).value = "Дата ответа"
            sheet.cell(row=1, column=7).value = "Текст ответа"
            num = 2

            self.status_label.setText(f"--Поиск отзывов для {company['name']} с {site} начался--")
            num = self.get_reviews_elements(company=company, site=site, num=num, sheet=sheet, driver=driver)

            set_styles_to_sheet(sheet=sheet, num=num)

        except Exception as _ex:
            self.error_label.setText("Возникла ошибка. Проверьте файл logs.log")
            logging.exception(_ex)

    def get_reviews_elements(self, company: object, site: str, num: int, sheet: Worksheet, driver: WebDriver) -> int:
        try:
            # Имитация человеческого поведения перед загрузкой
            random_delay(2.5, 4.5)

            driver.get(url=company[site]["url"])
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.1);")
            random_delay(1.5, 3.0)

            # Имитация движения мыши после загрузки страницы
            actions = ActionChains(driver)
            try:
                offset_x = random.randint(-100, 100)
                offset_y = random.randint(-100, 100)
                actions.move_by_offset(offset_x, offset_y).perform()
            except:
                pass

            random_delay(2.0, 4.0)

            action = ActionChains(driver)

            # Имитация наведения на элемент перед кликом
            try:
                clickable_element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, company[site]["clicked_element_css_selector"]))
                )
                action.move_to_element(clickable_element).perform()
                random_delay(0.8, 1.8)
                clickable_element.click()
                random_delay(2.5, 4.5)
            except Exception as click_ex:
                logging.warning(f"Ошибка клика на элемент: {click_ex}")
                pass

            # Ожидание загрузки элемента для прокрутки
            element = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, company[site]["scrolled_element_css_selector"]))
            )

            try:
                count_reviews_elem = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, company[site]["count_reviews_css_selector"]))
                )
                count_reviews = int(count_reviews_elem.text)
            except Exception as _ex:
                try:
                    count_reviews = int("".join(re.findall(r'\d+', driver.find_element(By.CSS_SELECTOR, company[site][
                        "count_reviews_css_selector"]).text)))
                except:
                    count_reviews = 50  # Значение по умолчанию

            timer_start = time.perf_counter()
            max_attempts = 20
            attempts = 0

            while attempts < max_attempts:
                current_reviews = len(
                    driver.find_elements(By.CSS_SELECTOR, company[site]["searched_card_css_selector"]))

                if abs(current_reviews - count_reviews) > 3 and time.perf_counter() - timer_start < 60:
                    self.status_label.setText(
                        f"Прочитано {current_reviews} отзывов из ~{count_reviews} с {site}...")

                    human_like_scroll(driver, element, scrolls=2)

                    random_delay(1.5, 3.5)
                    attempts += 1
                else:
                    # Обработка кнопок "Показать больше"
                    show_more_buttons = driver.find_elements(By.CSS_SELECTOR, company[site]["show_more_button"])
                    if show_more_buttons:
                        for el_to_open in show_more_buttons:
                            try:
                                action.move_to_element(el_to_open).perform()
                                random_delay(0.7, 1.7)

                                WebDriverWait(driver, 5).until(
                                    EC.element_to_be_clickable(el_to_open)
                                ).click()

                                random_delay(self.waitTime + 0.5, self.waitTime + 2.5)
                            except Exception as click_ex:
                                logging.warning(f"Ошибка клика на 'Показать больше': {click_ex}")
                                continue

                    # Обработка кнопок ответов на отзывы
                    answer_buttons = driver.find_elements(By.CSS_SELECTOR, company[site]["review_answer_css_selector"])
                    if answer_buttons:
                        for el_to_open in answer_buttons:
                            try:
                                action.move_to_element(el_to_open).perform()
                                random_delay(0.6, 1.6)
                                driver.execute_script("arguments[0].click();", el_to_open)
                                random_delay(self.waitTime, self.waitTime + 2.0)
                            except Exception as click_ex:
                                logging.warning(f"Ошибка клика на ответ: {click_ex}")
                                continue

                    final_count = len(
                        driver.find_elements(By.CSS_SELECTOR, company[site]["searched_card_css_selector"]))
                    self.status_label.setText(
                        f"--Всего прочитано {final_count} отзывов с {site}--")
                    self.status_label.setText(f"--Поиск отзывов с {site} закончился--")

                    num = self.get_reviews_data(driver=driver, company=company, site=site, num=num, sheet=sheet,
                                                html_els=driver.find_elements(By.CSS_SELECTOR,
                                                                              company[site][
                                                                                  "searched_card_css_selector"]))
                    return num

            # Если вышли по лимиту попыток
            final_count = len(driver.find_elements(By.CSS_SELECTOR, company[site]["searched_card_css_selector"]))
            self.status_label.setText(f"--Достигнут лимит попыток. Прочитано {final_count} отзывов с {site}--")
            num = self.get_reviews_data(driver=driver, company=company, site=site, num=num, sheet=sheet,
                                        html_els=driver.find_elements(By.CSS_SELECTOR,
                                                                      company[site]["searched_card_css_selector"]))
            return num

        except Exception as _ex:
            self.error_label.setText("Возникла ошибка. Проверьте файл logs.log")
            logging.exception(_ex)
            return num

    def get_reviews_data(self, driver: WebDriver, company: object, site: str, num: int, sheet: Worksheet,
                         html_els: list[WebElement]) -> int:
        self.status_label.setText(f"--Обработка данных для {company['name']} с {site}--")

        for idx, el in enumerate(html_els):
            # Добавляем случайные задержки при обработке каждого 3-го отзыва
            if idx > 0 and idx % 3 == 0:
                random_delay(0.3, 0.8)

            if check_exists(el, company[site]["review_answer_css_selector"]):
                sheet.cell(row=num, column=1).value = "+"  # review answer
            else:
                sheet.cell(row=num, column=1).value = "-"  # review answer

            if check_exists(el, company[site]["review_name_css_selector"]):
                sheet.cell(row=num, column=2).value = driver.execute_script('return arguments[0].textContent.trim();',
                                                                            el.find_element(By.CSS_SELECTOR,
                                                                                            company[site][
                                                                                                "review_name_css_selector"]))  # name
            else:
                sheet.cell(row=num, column=2).value = "None"

            if check_exists(el, company[site]["review_date_css_selector"]):
                sheet.cell(row=num, column=3).value = driver.execute_script('return arguments[0].textContent.trim();',
                                                                            el.find_element(By.CSS_SELECTOR,
                                                                                            company[site][
                                                                                                "review_date_css_selector"])).replace(
                    ", отредактирован", "(отредактирован)")  # date
            else:
                sheet.cell(row=num, column=3).value = "None"

            if check_exists(el, company[site]["review_rate_css_selector"]):
                sheet.cell(row=num, column=4).value = len(
                    el.find_elements(By.CSS_SELECTOR, company[site]["review_rate_css_selector"]))  # rate
            else:
                sheet.cell(row=num, column=4).value = "None"

            if check_exists(el, company[site]["review_text_css_selector"]):
                sheet.cell(row=num, column=5).value = driver.execute_script('return arguments[0].textContent.trim();',
                                                                            el.find_element(By.CSS_SELECTOR,
                                                                                            company[site][
                                                                                                "review_text_css_selector"]))  # text review
            else:
                sheet.cell(row=num, column=5).value = "None"

            if check_exists(el, company[site]["review_answer_date_css_selector"]):
                sheet.cell(row=num, column=6).value = driver.execute_script('return arguments[0].textContent.trim();',
                                                                            el.find_element(By.CSS_SELECTOR,
                                                                                            company[site][
                                                                                                "review_answer_date_css_selector"]))  # answer date
            else:
                sheet.cell(row=num, column=6).value = "None"

            if check_exists(el, company[site]["review_answer_text_css_selector"]):
                sheet.cell(row=num, column=7).value = driver.execute_script('return arguments[0].textContent.trim();',
                                                                            el.find_element(By.CSS_SELECTOR,
                                                                                            company[site][
                                                                                                "review_answer_text_css_selector"]))  # answer text
            else:
                sheet.cell(row=num, column=7).value = "None"

            num += 1

        self.status_label.setText(f"--Обработка данных с {site} завершена--")
        return num


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = App()
    window.show()
    app.exec()


if __name__ == '__main__':
    main()