from bs4 import BeautifulSoup
import requests
from datetime import datetime
import json

class Parser_worker():

    def __init__(self, headers, html_folder):

        self.headers = headers  # headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36'}
        self.html_folder = html_folder  # Папка для скалдывания html-файлов

    def get_till_page(self, baseurl):
        html = self.get_html(baseurl)
        soup = BeautifulSoup(html, "html.parser")
        pages = soup.find('div', {'class': 'uni-paging'})  # Находим тег со страницами
        if pages:
            try:
                last_page = pages.text.split("... ")[1].strip()  # 1, 2, 3, 4... 78
                return int(last_page)
            except:
                last_page = pages.text.split(" ")[-2].strip()  # 1, 2, 3, 4, 8
                return int(last_page)


    def get_html(self, url):
        try:
            res = requests.get(url, self.headers)
        except requests.ConnectionError:
            return
        if res.status_code < 400:
            return res.content

    def get_today_date(self):  # Получаем егодняшнюю дату в формате строки '13.07.2017'
        date = datetime.strftime(datetime.now(), "%d.%m.%Y")
        return date

    def convert_realtdate_to_str(self, realtdate):  # Конвертируем дату в строку
        date = datetime.strptime(realtdate, "%Y-%m-%d")
        date_new = datetime.strftime(date, "%d.%m.%Y")
        return date_new

    def get_kurs(self, valuta):  # Получаем курс доллара На вход - запрос апи с сайта нацбанка
        kurs = None
        if valuta == 'EUR':
            byte_kurs = self.get_html('http://www.nbrb.by/API/ExRates/Rates/292')  # b'{...} 292 - код валюты евро
            dict_kurs = json.loads(byte_kurs)  # {...}
            kurs = dict_kurs['Cur_OfficialRate']
        elif valuta == 'USD':
            byte_kurs = self.get_html('http://www.nbrb.by/API/ExRates/Rates/145')  # b'{...} 145 - код валюты долл
            dict_kurs = json.loads(byte_kurs)  # {...}
            kurs = dict_kurs['Cur_OfficialRate']
        return kurs  # 1.9750

    def del_space(self, string):  # Для удаления спец пробела в ЦЕНЕ - 1 670. Возвращает число в формате float без пробела. На вход - строка с ценой '1 670' or '879'
        string = str(string)
        if ' ' in string:
            new_string = string.replace(' ', '')
            string = float(new_string)
            return string
        else:
            if ' млн' in string:  # "1.37 млн" - на вход - еще нужно проверить цену на наличие слова "млн" - такие есть в ресторанах - когда стоимость дана за весь уасток.
                string = string.split(' млн')[0]  # "1.37 млн" - "1.37"
                string = float(string) * 1000000  # переводим строку в число и умножаем на 1 млн, т.к. стоимость была дана в млн.
                return string
            else:
                string = float(string)
                return string

    def del_coma(self, string):  # Для удаления запятой  в ЦЕНЕ - 5,76 и замене ее на точку. Возвращает строку
        if ',' in string:
            string = string.replace(',', '.')
            return string
        else:
            return string

    def get_finish_area(self, realt_answer, project, Excel_field1, Excel_field2):  # Получаем из ответа только площадь - избавляемся от м², где realt_answer - ответ на реалте
        if "от" in realt_answer:  # Если площадь указана "от 34 до 56 м²"
            area_from = realt_answer.split('от')[1].split('до')[0].strip()
            area_till = realt_answer.split('до')[1].split('м²')[0].strip()
        elif "до" in realt_answer:  # Если площадь указана "до 403 м²"
            area_from = realt_answer.split('до')[1].split('м²')[0].strip()
            area_till = realt_answer.split('до')[1].split('м²')[0].strip()
        else:  # Если площадь указана "403 м²" Получаем из ответа только площадь - избавляемся от м², где realt_answer - ответ на реалте и в поле от и в поле до записываем одну площадь
            area_from = realt_answer.split('м²')[0].strip()
            area_till = realt_answer.split('м²')[0].strip()
        project[Excel_field1] = float(area_from)
        project[Excel_field2] = float(area_till)

    def get_contacts(self, realt_answer, project, Excel_field):  # Если в поле ответа "Контактные данные" есть 'Пожалуйста, скажите что Вы нашли это объявление на сайте Realt.by', то удаляется эта часть
        if 'Пожалуйста, скажите что Вы нашли это объявление на сайте Realt.by' in realt_answer:
            project[Excel_field] = realt_answer.split('Пожалуйста, скажите что Вы нашли это объявление на сайте Realt.by')[1].strip()
        else:
            project[Excel_field] = realt_answer

    def get_hight(self, realt_answer):  # Получаем из ответа только высоту - избавляемся от м, где realt_answer - ответ на реалте
        realt_answer = realt_answer.split('м')[0].strip()
        return round(float(realt_answer), 2)

    def get_zu_area(realt, realt_answer):  # Получаем площадь земельного участка(в гектарах), которая записана 8 соток
        realt_answer = realt_answer.split(' ')[0]
        return round(float(realt_answer)/100, 4)

    def write_webpage_to_html(self, id_object_name, html_obj, html_folder, todays_date):  # Для записи веб-страницы в папку на локальном компьютере в файл формата .html "24100_29.12.2017.html"
        name_html = '{}/{}_{}.html'.format(html_folder, id_object_name, todays_date)  # "24100_29.12.2017.html"
        with open(name_html, 'wb') as file:
            file.write(html_obj)

    # НОВАЯ ФУНКЦИЯ ПО ЗАПИСИ КООРДИНАТ Х и У
    def get_coords_new(self, soup, project):  # НОВОЕ Для каждого объявления находит координату с Яндекс карты, На вход - объект суп страницы с объявлением и словарь, куда записываются координаты
        table = soup.find_all('div', {'id': 'map-center'})  # Координаты записаны под данны тегом
        for i in table:
            # str(i) = <div data-center='{"distance":"1000","position.":{"x":"27.491466","y":"53.892999"},"image":"&lt;i class=\"icomoon-house\"&gt;&lt;\/i&gt;","name":"г. Минск, Мавра ул., 41"}' id="map-center"> </div>
            coords_str = str(i).split('"position.":{')[1].split('},"image')[0]  # "x":"27.491466","y":"53.892999"
            # print(coords_str)
            X = float(coords_str.split('x":"')[1].split('",')[0])  # 27.491466
            Y = float(coords_str.split('"y":"')[1].split('"')[0])  # 53.892999
            # print(X, Y)
            project['XCoord'] = X
            project['YCoord'] = Y


    # 3 ФУНКЦИИ ДЛЯ ОПРЕДЕЛЕНИЯ ВИДА и ВСПОМОГАТЕЛЬНЫХ ВИДОВ
    def write_into_project_all_vidy(self, osnov_vid, project, Excel_field, Excel_field2, Excel_field3, excel_options_dict):
        project[Excel_field] = excel_options_dict[Excel_field][osnov_vid]
        project[Excel_field2] = excel_options_dict[Excel_field2][osnov_vid]
        project[Excel_field3] = excel_options_dict[Excel_field3][osnov_vid]

    def get_vidy_in_brackets(self, realt_answer, project, Excel_field, Excel_field2, Excel_field3, Excel_field4, excel_options_dict):
        osnov_vid = realt_answer.split(")")[-2].split("(")[1].lower()  # Нужно взять только то, что в последних скобках - Делим по последней скобке и берем предпоследний эл-т - [-2] - второй элемент с конца, т.к. первый элемент с конца - пустая строка
        if ',' in osnov_vid:  # если в скобочках записано более чем один доп вид - т.е. есть запятая. ПОЧТИ ВСЕГДА
            osnov_vid = osnov_vid.split(",")[0]  # если в скобочках записано более чем один доп вид. ПОЧТИ ВСЕГДА
            if osnov_vid == 'помещение':  # Иногда первым в скобчках записано помещение и оно склад, т.е. основной вид идет после помещения. поэтому ужно взять второй элемент в скобках
                osnov_vid = realt_answer.split(")")[-2].split("(")[1].lower()  # получается "помещение, склад, ...
                project[Excel_field4] = osnov_vid.split(',', 2)[-1]  # "помещение, склад, холодильник, офис, складик" сплитим на 3 объекта: помещение склад и остальное и берем последнее
                osnov_vid = osnov_vid.split(', ')[1]  # из "помещение, склад, ... получаем склад
            else:
                vidy = realt_answer.split(")")[-2].split("(")[1].lower()  # склад, холодильник, офис, складик
                project[Excel_field4] = vidy.split(',', 1)[-1]  # "склад, холодильник, офис, складик" сплитим на 2 объекта: псклад и остальное - и берем последнее "холодильник, офис, складик"
            self.write_into_project_all_vidy(osnov_vid, project, Excel_field, Excel_field2, Excel_field3, excel_options_dict)

        else:  # если в скобочках записан один вид
            self.write_into_project_all_vidy(osnov_vid, project, Excel_field, Excel_field2, Excel_field3, excel_options_dict)

    def get_finish_vid_object(self, realt_answer, project, Excel_field, Excel_field2, Excel_field3, Excel_field4, excel_options_dict):  # из многообразия того что записано в поле Вид объекта, нужно определить основной вид и в определить по нему Вид объекта, Наименование и Назначение
        if "(" in realt_answer:  # такой вид: "Продажа офисов от застройщика в новостройке по пр.Дзержинского (офис)"
            osnov_vid = realt_answer.split(" (")[0].strip().lower()  # "Продажа офисов от застройщика в новостройке по пр.Дзержинского"
            if len(osnov_vid) <= 18:  # "торговое помещение" - 18 символов. Самое длинное из возможных вариантов на реалте, которое можеть быть записано до скобки. Но бывает что до скобок записана ерунда и она тоже меньше 18 символов
                try:
                    self.write_into_project_all_vidy(osnov_vid, project, Excel_field, Excel_field2, Excel_field3, excel_options_dict)
                    project[Excel_field4] = realt_answer.split("(")[1].split(")")[0]  # Записываем в проект вспомогательные виды - они в скобках
                except KeyError:  # Если все-таки до скобок хоть и меньше 18 символов, но записана ерунда, основной вид записан в скобках
                    self.get_vidy_in_brackets(realt_answer, project, Excel_field, Excel_field2, Excel_field3, Excel_field4, excel_options_dict)
            else:  # если символов до "(" большее 18, то основной вид записан первым после послдней скобки "("
                self.get_vidy_in_brackets(realt_answer, project, Excel_field, Excel_field2, Excel_field3, Excel_field4, excel_options_dict)
        else:
            osnov_vid = realt_answer.lower()  # Если в поле нет скобочек, значит записан только один основной вид
            self.write_into_project_all_vidy(osnov_vid, project, Excel_field, Excel_field2, Excel_field3, excel_options_dict)


    # 5 ФУНКЦИИ ДЛЯ ПОЛУЧЕНИЯ ЦЕНЫ И РАСЧЕТОВ ЦЕНЫ

    def calculate_price_whole_lot(self, project, Excel_field1, Excel_field2, Excel_field3, Excel_field4, Excel_field5):
        if 'Общая площадь до, кв.м' in project:  # Проверка! Т.к. плоащди может не быть!!!
            if project['Общая площадь до, кв.м'] != project['Общая площадь от, кв.м'] and project[Excel_field4] != project[Excel_field5]:
                pass  # В таком случае ничего зваписвывпать неправильно, так как логика неправильная выходит
            else:
                project[Excel_field1] = round(project[Excel_field4] * project['Общая площадь до, кв.м'], 2)
                project[Excel_field2] = round(project[Excel_field5] * project['Общая площадь от, кв.м'], 2)
                project[Excel_field3] = 'НКА'

    def calculate_price_metr(self, project, Excel_field1, Excel_field2, Excel_field4, Excel_field5, Excel_field6):
        if 'Общая площадь до, кв.м' in project:  # Проверка! Т.к. плоащди может не быть!!!
            project[Excel_field4] = round(project[Excel_field1] / project['Общая площадь до, кв.м'], 2)
            project[Excel_field5] = round(project[Excel_field2] / project['Общая площадь от, кв.м'], 2)
            project[Excel_field6] = 'НКА'

    def get_price(self, price_from, price_till, project,Exc_field1, Exc_field2, Exc_field3, valuta):  # (Если цена указана за 1 кв.м.) Функция удаляет запятую и пробел в цене если они есть и записывает в project цену переведенную в долларах с округлением до 2 знаков после запятой
        price_from = self.del_coma(price_from)
        price_till = self.del_coma(price_till)
        price_from = self.del_space(price_from)
        price_till = self.del_space(price_till)
        project[Exc_field1] = round((price_from / self.get_kurs(valuta)), 2)
        project[Exc_field2] = round((price_till / self.get_kurs(valuta)), 2)
        project[Exc_field3] = 'Сайт'
        print(Exc_field1, Exc_field2, Exc_field3, project[Exc_field1], project[Exc_field2], project[Exc_field3])

    def check_price(self, realt_answer, project, Exc_field1, Exc_field2, Exc_field3, valuta):  # realt_answer = 'до 92 750', '355—395' or '356' or '1 876' or '1 355—1 395' or 'до 2 157'(Если цена указана за 1 кв.м.)  Функция проверяет есть ли в цене строки '—', 'до ' или нет. и записывает корректную (если были строки) в долларах
        if '—' in realt_answer:  # Цена записана в таком виде: 114 911 — 120 658 руб, 1 973—2 072 руб/кв.м
            price_from = realt_answer.split('—')[0].strip()
            price_till = realt_answer.split('—')[1].strip()
            self.get_price(price_from, price_till, project, Exc_field1, Exc_field2, Exc_field3, valuta)
        elif 'до ' in realt_answer:  # Цена записана в таком виде: до 92 750 руб, до 2 157 руб/кв.м
            realt_answer = realt_answer.split('до ')[1].strip()
            price_from = realt_answer
            price_till = realt_answer
            self.get_price(price_from, price_till, project, Exc_field1, Exc_field2, Exc_field3, valuta)
        else:
            price_from = realt_answer
            price_till = realt_answer
            self.get_price(price_from, price_till, project, Exc_field1, Exc_field2, Exc_field3, valuta)

    def get_finish_price(self, realt_answer, project, Excel_field1, Excel_field2, Excel_field3, Excel_field4, Excel_field5, Excel_field6, valuta):  # где realt_answer - ответ на реалте, project - словарь, куда записывается итоговая цена,   Excel_field = Realt_Excel_dict[option] - Получаем название поля в Excel, option - название поля на реалте
        # print(realt_answer) - сто пятьсот вариантов
        if "руб, " in realt_answer:  # Цена записана в таком виде: 11 740 руб, 118 руб/кв.м
            realt_answer_metr = realt_answer.split('руб, ')[1].split('руб/')[0].strip()  # 118    '355—395' or '356' or '1 876' or '1 355—1 395' or 'до 2 157'
            self.check_price(realt_answer_metr, project, Excel_field4, Excel_field5, Excel_field6, valuta)
            # Еще нужно записать цену за весь участок
            realt_answer_whole_lot = realt_answer.split('руб, ')[0].strip()  # 11 740    '355—395' or '356' or '1 876' or '1 355—1 395' or 'до 2 157'
            self.check_price(realt_answer_whole_lot, project, Excel_field1, Excel_field2, Excel_field3, valuta)
        elif 'договор' in realt_answer:  # Цена записана в таком виде: Цена договорная
            project[Excel_field3] = 'Цена договорная'
            project[Excel_field6] = 'Цена договорная'
        else:  # Значит цена указана либо за весь объект либо только за 1 кв.м. Варианты: 118 руб/кв.м, 118 руб/м², 22-33 руб/кв.м, 345 руб
            if ' руб/' in realt_answer:  # есть цена за 1 кв.м
                realt_answer = realt_answer.split(' руб/')[0].strip()
                self.check_price(realt_answer, project, Excel_field4, Excel_field5, Excel_field6, valuta)
                # Если цена указана только за 1 кв. м, то теперь нужно рассчитать цену за весь земельный участок, НО в случае если Sот не равна Sдо и Цена за 1 кв.м от  не равна Цене за 1 кв.м до ТОГДА НЕ РССЧИТЫВАЕМ
                self.calculate_price_whole_lot(project, Excel_field1, Excel_field2, Excel_field3, Excel_field4, Excel_field5)
            else:  # значит стоимость дана только за весь земельный участок и в строке есть слово "руб"
                realt_answer = realt_answer.split(' руб')[0].strip()
                self.check_price(realt_answer, project, Excel_field1, Excel_field2, Excel_field3, valuta)
                # Если цена указана только за весь объект, то теперь нужно рассчитать цену за 1 кв.м
                self.calculate_price_metr(project, Excel_field1, Excel_field2, Excel_field4, Excel_field5, Excel_field6)
    # 3 функции для определения адреса

    def get_street(self, realt_answer):  # Получаем название улицы
        if len(realt_answer.split(".")) > 2:  # Значит "С. Ковалевской ул."
            elems_by_point = realt_answer.split(".") # разбиваем ответ по точке и получаем список (из 3 элементов например)
            street_elem = realt_answer.split(elems_by_point[-1])[0]  # разбиваем ответ последним элементом из elems_by_point чтобы получить улицу и эудс
            street_elem_by_point = street_elem.split[' ']
            realt_street_name = street_elem.split(street_elem_by_point[-1])[0]
        else:
            street_elem = realt_answer.split(".")[0]  # Никольская ул
            realt_elem_name = street_elem.split(' ')[-1]  # ул
            realt_street_name = street_elem.split(realt_elem_name)[0].strip()  # 40 лет Победы
        return realt_street_name

    def get_elem(self, realt_answer):  # Получаем название ЭУДС
        if len(realt_answer.split(".")) > 2:  # Значит "С. Ковалевской ул."
            elems_by_point = realt_answer.split(".") # разбиваем ответ по точке и получаем список (из 3 элементов например)
            street_elem = realt_answer.split(elems_by_point[-1])[0]  # разбиваем ответ последним элементом из elems_by_point чтобы получить улицу и эудс
            realt_elem_name = street_elem.split(' ')[-1]
        else:
            street_elem = realt_answer.split(".")[0]  # Никольская ул
            realt_elem_name = street_elem.split(' ')[-1]  # ул
        return realt_elem_name

    def get_full_address(self, realt_answer, project, Excel_field1, Excel_field2, Excel_field3, Excel_field4, id_object_name, excel_options_dict):  # Никольская ул., 66-2, 40 лет Победы ул., 66-2,
        if "." in realt_answer:  # Сначала ищем в ответе точку - она всегда после ЭУДС (ВОПРОС - точно ли ВСЕГДА - например есть улица "С. Ковалевской" или "Меньковский тракт"
            try:
                realt_street_name = self.get_street(realt_answer)
                project[Excel_field2] = excel_options_dict[Excel_field2][realt_street_name]
            except IndexError:
                print("Для объекта с номером {} / Невозможно определить улицу".format(id_object_name))
            except KeyError:  # улицы нет в словаре Offices_Realt_Fields_Options
                realt_street_name = self.get_street(realt_answer)
                project[Excel_field2] = "{} / не из классификатора".format(realt_street_name)
            try:
                realt_elem_name = self.get_elem(realt_answer)
                project[Excel_field1] = excel_options_dict[Excel_field1][realt_elem_name]
            except IndexError:
                print("Для объекта с номером {} / Невозможно определить ЭУДС".format(id_object_name))
            except KeyError:
                realt_elem_name = self.get_elem(realt_answer)
                project[Excel_field1] = "{} / не из классификатора".format(realt_elem_name)
        else:  # обработка если в ЭУДС нет точки - например Меньковский тракт
            if "," in realt_answer:  # Меньковский тракт, 43
                try:
                    street_elem = realt_answer.split(",")[0]
                    realt_elem_name = street_elem.split(' ')[-1]  # ул
                    realt_street_name = street_elem.split(realt_elem_name)[0].strip()  # 40 лет Победы
                    project[Excel_field2] = excel_options_dict[Excel_field2][realt_street_name]
                    project[Excel_field1] = excel_options_dict[Excel_field1][realt_elem_name]
                except KeyError:  # сделано для адреса Центральная, 13
                    print("Неправильная структура УДС")  # сделано для адреса "Центральная, 13
            else:  # Меньковский тракт
                try:
                    realt_elem_name = realt_answer.split(' ')[-1]  # ул
                    realt_street_name = realt_answer.split(realt_elem_name)[0].strip()  # 40 лет Победы
                    project[Excel_field2] = excel_options_dict[Excel_field2][realt_street_name]
                    project[Excel_field1] = excel_options_dict[Excel_field1][realt_elem_name]
                except KeyError:  # сделано для адреса "Брест"
                    print("Неправильная структура УДС")
        if "," in realt_answer:  # если в ответе есть запятая, значит указаны номер дома/корпус
            house_korp = realt_answer.split(',')[1].strip()  # 66-2 или 66 or "66-2 Информация о доме"
            if "-" in house_korp:  # значит есть корпус
                house = house_korp.split('-')[0]
                project[Excel_field3] = int(house)
                korp = house_korp.split('-')[1].strip()
                if " " in korp:  # если после адреса есть строка "Информация о доме"
                    korp = korp.split(' ')[0]
                project[Excel_field4] = korp  # корпус может быть не только integer
            else:  # значит нет корпуса
                house = house_korp
                if " " in house_korp:  # если после адреса есть строка "Информация о доме"
                    house = house.split(' ')[0]
                project[Excel_field3] = int(house)

    def parse_object(self, obj_url: object, excel_objects: object, Realt_Excel_fields_dict: object, realt_fields_list: object,
                     excel_options_dict: object) -> object:
        html_obj = self.get_html(obj_url)
        soup = BeautifulSoup(html_obj, "html.parser")
        table = soup.find_all('tr', {'class': 'table-row'})  # Получаем список со всеми необходимыми данными объявления
        id_object_name = int(obj_url.split('object/')[1][:-1])  # из url страницы объявления оставляем только уникальный номер. Из этого https://realt.by/sale/offices/object/712024/ - получаем 712024
        object_date = None
        try:
            for row in table:  # Проходимся по каждой строке на странице объявления
                option = row.find('td',{'class': "table-row-left"}).text  # в строке выделяем левую часть - т.е. название поля
                if "Дата обновления" in option:
                    realt_answer = row.find('td', {'class': "table-row-right"}).text  # в строке выделяем правую часть - т.е. ответ
                    print('Date of object update', realt_answer)
                    object_date = self.convert_realtdate_to_str(realt_answer)
                    print('Date of writing update date', object_date)
                    break
            object_initials = (id_object_name, object_date)
            print("\nObject initials (№ and Update date: ", object_initials)
            if object_initials in excel_objects:
                print('Yes. Object IS in Excel\n')
                project = (excel_objects[object_initials], id_object_name, object_date)
            else:
                excel_objects[object_initials] = None # Добавляем в словарь excel_objects (номер объявления, дата объявления), чтобы исключить дубли с разных страниц сайта при парсинге
                project = {}
                print('No. Object is NOT in excel')
                project['№ Объявления'] = id_object_name
                todays_date = self.get_today_date()
                project['Дата парсинга'] = todays_date
                project['Ссылка на HTML-страницу'] = '{}/{}_{}.html'.format(self.html_folder, id_object_name, object_date)
                try:
                    text_above_object = soup.find('div', {'class': 'text-12 mb20 fl wp100'}).text  # Описание над объявлением записано под данным тегом
                    project['Описание над объявлением'] = text_above_object
                    print(text_above_object)
                except:
                    with open("mistakes.txt", "a") as myfile:
                        myfile.write('No text above object {} {} {}\n'.format(Exception.__class__, Exception.args, id_object_name))
                try: # Для бизнес-центров текст над объявлением хранится в другом теге
                    text_above_object = soup.find('div', {'class': 'object-desc'}).text  # Описание над объявлением записано под данным тегом
                    project['Описание над объявлением'] = text_above_object
                    print(text_above_object)
                except:
                    with open("mistakes.txt", "a") as myfile:
                        myfile.write('No text above object {} {} {}\n'.format(Exception.__class__, Exception.args, id_object_name))

                # Для бизнес-центров и торговых центров в объявлении нет строки "Вид объекта". ПОэтому записываем его таким образом
                if 'malls' in obj_url:  # Если парсим Торговые центры
                    Excel_field = Realt_Excel_fields_dict['Вид объекта']
                    Excel_field2 = Realt_Excel_fields_dict['Вид объекта2']
                    Excel_field3 = Realt_Excel_fields_dict['Вид объекта3']
                    osnov_vid = "торговый центр"
                    self.write_into_project_all_vidy(osnov_vid, project, Excel_field, Excel_field2, Excel_field3, excel_options_dict)

                if 'newoffices' in obj_url:  # Если парсим Бизнес-центры
                    Excel_field = Realt_Excel_fields_dict['Вид объекта']
                    Excel_field2 = Realt_Excel_fields_dict['Вид объекта2']
                    Excel_field3 = Realt_Excel_fields_dict['Вид объекта3']
                    osnov_vid = "бизнес-центр"
                    self.write_into_project_all_vidy(osnov_vid, project, Excel_field, Excel_field2, Excel_field3, excel_options_dict)
                try:
                    self.get_coords_new(soup, project)
                except:
                    with open("mistakes.txt", "a") as myfile:
                        myfile.write('Error in COORDS {} {} {}\n'.format(Exception.__class__, Exception.args, id_object_name))

                self.write_webpage_to_html(id_object_name, html_obj, self.html_folder, object_date)  # write web page to html file

                for row in table:  # Проходимся по каждой строке на странице объявления

                    option = row.find('td', {'class': "table-row-left"}).text  # в строке выделяем левую часть - т.е. название поля
                    if option in realt_fields_list:  # если полю на реалте есть соответствие в словаре Offices_Realt_Excel - значит его обрабатываем, т.к. в объявлении есть лишние поля, которые нам не нужны
                        realt_answer = row.find('td',{'class': "table-row-right"}).text  # в строке выделяем правую часть - т.е. ответ
                        print("", option, ": ", realt_answer)
                        Excel_field = Realt_Excel_fields_dict[option]

                        if option == "Площадь":
                            # Excel_field = 'Площадь от' - уже есть
                            Excel_field2 = Realt_Excel_fields_dict['Площадь до']
                            self.get_finish_area(realt_answer, project, Excel_field, Excel_field2)

                        elif option == 'Дата обновления':
                            project['Дата обновления'] = object_date

                        elif option == "Вид объекта":
                            Excel_field2 = Realt_Excel_fields_dict['Вид объекта2']
                            Excel_field3 = Realt_Excel_fields_dict['Вид объекта3']
                            Excel_field4 = Realt_Excel_fields_dict['Вспомогательные виды']
                            self.get_finish_vid_object(realt_answer, project, Excel_field, Excel_field2, Excel_field3, Excel_field4, excel_options_dict)

                        elif option == "Ориентировочная стоимость эквивалентна": # Необходимым условием является наличине в project['Тип предложения'] т.к. если аренда- то нужно переводить в евро, если продажа, то в доллары
                            # realt_answer = 1 677 руб/кв.м 1 677 руб/кв.м  Цена сделки определяется по соглашению сторон. Расчеты осуществляются в белорусских рублях в соответствии с законодательством Республики Беларусь.
                            Excel_field1 = Realt_Excel_fields_dict['Цена от']
                            Excel_field2 = Realt_Excel_fields_dict['Цена до']
                            Excel_field3 = Realt_Excel_fields_dict['Маркер Цена']
                            Excel_field4 = Realt_Excel_fields_dict['Цена за 1м2 от']
                            Excel_field5 = Realt_Excel_fields_dict['Цена за 1м2 до']
                            Excel_field6 = Realt_Excel_fields_dict['Маркер Цена а 1м2']
                            Excel_field7 = Realt_Excel_fields_dict['Валюта']
                            valuta = 'USD' if project['Тип предложения'] == 'Продажа' else 'EUR'
                            project[Excel_field7] = valuta
                            self.get_finish_price(realt_answer, project, Excel_field1, Excel_field2, Excel_field3, Excel_field4,Excel_field5, Excel_field6, valuta)

                        elif option == "Телефоны":
                            self.get_contacts(realt_answer, project, Excel_field)

                        elif option == "Вода":
                            Excel_field1 = Realt_Excel_fields_dict['Вода холодная']
                            Excel_field2 = Realt_Excel_fields_dict['Вода горячая']  # Горяее водоснабжение
                            project[Excel_field1] = excel_options_dict[Excel_field1][realt_answer]
                            project[Excel_field2] = excel_options_dict[Excel_field2][realt_answer]

                        elif option == "Высота потолков":
                            project[Excel_field] = self.get_hight(realt_answer)

                        elif option == "Площадь участка":
                            project[Excel_field] = self.get_zu_area(realt_answer)

                        elif option == "Адрес":  # Никольская ул., 66-2, 40 лет Победы ул., 66-2,
                            try:
                                Excel_field1 = Realt_Excel_fields_dict['ЭУДС'] # элемент улицы
                                Excel_field2 = Realt_Excel_fields_dict['улица']  # название улицы
                                Excel_field3 = Realt_Excel_fields_dict['дом']  # номер дома
                                Excel_field4 = Realt_Excel_fields_dict['корпус']  # корпус
                                project['Полный адрес'] = realt_answer  #добавляем полный адрес
                                self.get_full_address(realt_answer, project, Excel_field1, Excel_field2, Excel_field3, Excel_field4, id_object_name, excel_options_dict)
                            except:
                                with open("mistakes.txt", "a") as myfile:
                                    myfile.write('Error in Address {} {} {}\n'.format(Exception.__class__, Exception.args, id_object_name)) # Если случай, который не предусмотрен - это уже из ряда вон выходящий адрес


                        elif option == "Район области":
                            project[Excel_field] = realt_answer.split('район')[0].strip()

                        elif option == "Дополнительно" or option == "Примечания": # Данные поля склеиваем и записываем в Ексель в одно поле "Примечание"
                            if Excel_field in project:  # Если в проект уже было занесено поле "Примечание" (Сначала записывается Дополнительно, а потом Примечания)
                                project[Excel_field] = project[Excel_field] + realt_answer
                            else: # Если в проект еще не было занесено поле "Примечание"
                                project[Excel_field] = realt_answer

                        elif Realt_Excel_fields_dict[option] in excel_options_dict:  # доп действий производить не нужно, ответ записан в той форме, в которой он в словаре Offices_Realt_Fields_Options
                            try:
                                project[Excel_field] = excel_options_dict[Excel_field][realt_answer]
                            except KeyError:  # есть объявление https://realt.by/sale/shops/object/1106690/ у которого Материал стен не из классификатора
                                project[Excel_field] = '{} / не из классификатора'.format(realt_answer)

                        else:  # Записываем ответ как он есть
                            project[Excel_field] = realt_answer
            return project, id_object_name
        except:
            with open("mistakes.txt", "a") as myfile:
                myfile.write('Error in all object {} {} {}\n'.format(Exception.__class__, Exception.args, id_object_name))


    def get_page_projects(self, page_url, excel_objects, Realt_Excel_fields_dict, realt_fields_list, excel_options_dict): # Парсим одну страницу
        url = self.get_html(page_url)
        soup = BeautifulSoup(url, "html.parser")
        table = soup.find_all('div', {'class': 'bd-item'})  # Получаем список с объявлениями на странице
        page_projects_in_excel = []
        new_page_projects_for_excel = []
        id_object_name_at_page = []
        for row in table:
            # получаем ссылку каждого объявления
            href_name = row.find('a')
            obj_url = href_name.get("href")
            project, id_object_name = self.parse_object(obj_url, excel_objects, Realt_Excel_fields_dict, realt_fields_list, excel_options_dict)
            if isinstance(project, tuple):  # Если проект сущесмтвует в ексель(значит вернется кортеж - (3, 12365, "01.01.2017")
                page_projects_in_excel.append(project)
            else:
                new_page_projects_for_excel.append(project)
            id_object_name_at_page.append(id_object_name)
        return page_projects_in_excel, new_page_projects_for_excel, id_object_name_at_page
