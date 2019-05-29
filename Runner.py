import parse_utils, excel_utils
import json
from datetime import date
import time, random
from ast import literal_eval


class Runner(object):
    # Задаваемые атрибуты

    # список урлов для дальнейших парсингов. В существующем екселе должны быть отпаршены именно эти урлы и именно столько. Потому что если в этой переменной меньше урлов, то в екселе для некоторых объектов будет рассчитана экспозиция. А это неправильно

    # Общественно-деловая недвижимость
    # baseurls = ['https://realt.by/sale/offices/', 'https://realt.by/rent/offices/', 'https://realt.by/sale/shops/', 'https://realt.by/rent/shops/',
    #             'https://realt.by/sale/restorant-cafe', 'https://realt.by/rent/restorant-cafe']
    #   офисы(аренда) + БЦ
    #baseurls = ['https://realt.by/rent/offices/', 'https://realt.by/newoffices/', 'https://realt.by/malls/']

    # Бизнес-центры
    # baseurls = ['https://realt.by/newoffices/', 'https://realt.by/malls/']

    # Склады
    # baseurls = ['https://realt.by/sale/warehouses/', 'https://realt.by/rent/warehouses/']

    #всё
    baseurls = ['https://realt.by/sale/warehouses/', 'https://realt.by/rent/warehouses/', 'https://realt.by/sale/offices/', 'https://realt.by/rent/offices/', 'https://realt.by/sale/shops/',
                'https://realt.by/rent/shops/', 'https://realt.by/sale/restorant-cafe', 'https://realt.by/rent/restorant-cafe','https://realt.by/newoffices/', 'https://realt.by/malls/']

    realt_Excels_fields_json = 'Offices_Realt_Excel'  # файл .json с соответствиями м/д названиями полей на realt.by и excel
    realt_Excels_fields_Options_json = 'Offices_Realt_Fields_Options'  # файл .json с соответствиями м/д вариантами ответов на реалте и ответами, которые мы хотим записывать в ексель

    excel_path = "Форма_04_Предложения_продажи_и_аренды.xlsx"  # файл excel куда парсить
    excel_sheet = "Предложения"  # лист файла excel куда парсить

    excel_headers_row = 3  # номер строки с заголовками в ексель

    # Названия полей, которые используются в парсинге (однако не все) еще есть поля в фале parse_utils, где используются навзания полей из екселя
    num_object_excel_col_name = "№ Объявления"
    update_object_excel_col_name = "Дата обновления"
    parsing_date_excel_col_name = "Дата парсинга"
    html_excel_col_name = "Ссылка на HTML-страницу"
    expozition_col_name = "Расчетная экспозиция"

    html_folder = 'HTMLs'

    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36'}

    from_page = 1
    # till_page = 2  # При повторном парсинге, когда сразу парсится всё (много базовых урлов) данная переменная определяется кодом и данную строку необходимо закомментировать


    def __init__(self) -> object:
        # При инициализации класса


        # создаем excel_worker и parser_worker для дальнейшего использования функций из других модулей (классов)
        self.excel_worker = excel_utils.Excel_worker(Runner.excel_path, Runner.excel_sheet, Runner.excel_headers_row,
                                                     Runner.num_object_excel_col_name, Runner.update_object_excel_col_name, Runner.parsing_date_excel_col_name, Runner.html_excel_col_name)

        self.parser_worker = parse_utils.Parser_worker(Runner.headers, Runner.html_folder)

        with open(Runner.realt_Excels_fields_json, 'r', encoding='utf-8') as jf:  # открываем файл на чтение
            self.Realt_Excel_fields_dict = json.load(jf)  # загружаем из файла данные в словарь Realt_Excel_fields_dict = {'Вид объекта': 'Наименование', 'Вид объекта2': 'Назначение', 'Условия сделки': 'Тип предложения', ...
        self.excel_fields_list = list(self.Realt_Excel_fields_dict.values())  # Вроде как нигде не нужно: Cоздаем лист с полями Ексель - ['Наименование', 'Назначение', 'Тип предложения', 'Контактные данные'...
        self.realt_fields_list = list(self.Realt_Excel_fields_dict.keys())

        with open(Runner.realt_Excels_fields_Options_json, 'r', encoding='utf-8') as jf:  # открываем файл на чтение
            self.excel_options_dict = json.load(jf)  # загружаем из файла данные в словарь excel_options_dict = {'Вид объекта': 'Наименование',"Электроснабжение": { "есть": "Да ","нет": "Нет","220В": "Да ","380В": "Да"}, ...}

    def calculate_expozition(self, date_first, date_second):  # Расчет экспозиции - расчет разности (в днях) между двумя датами
        d0 = date(date_first)
        d1 = date(date_second)
        delta = d0 - d1
        return delta.days

    def parse_page(self, page_url, excel_objects):
        page_projects_in_excel, new_page_projects_for_excel, id_object_name_at_page = self.parser_worker.get_page_projects(page_url, excel_objects, self.Realt_Excel_fields_dict, self.realt_fields_list, self.excel_options_dict)  # Получаем со страницы  списки объектов, которые есть в екселе,  те которые надо записать и список номеров объявлений со страницы (они нужны для дальнейшего определения объявлений которых уже нет на сайте и для которых надо рассчитать экспозицию)
        print('new_page_projects_for_excel: ', new_page_projects_for_excel)
        self.excel_worker.add_projects_into_existing_excel(new_page_projects_for_excel)  # Записываем объекты, которых нет в екселе в сущ ексель
        return page_projects_in_excel, id_object_name_at_page  # Возвращаем те объекты со страницы, которые есть в ексель, чтобы заменить в ексель для них дату парсинга на сегодняшнюю и список номеров объявлений, чтобы выявить те объявления которые пропали и для которых надо будет рассчитать экспозицию


    def parse_pages(self):
        excel_objects = self.excel_worker.get_present_object_list()  # получаем словарь {(1206333, '15.12.2017'): 4, (1035344, '15.12.2017'): 5, ...} - все объявления, которые уже записаны в екселе (номер обьъявления, дата обновления и номер строки)
        with open("excel_objects.txt", "w") as myfile:
            myfile.write(str(excel_objects))
        print("In Excel already exist: ", excel_objects)
        all_projects_in_excel = []  # переменная, в которую будут записываться (номер строки в ексель, номер объявления, дата обновления) с сайта, которые уже есть в екселе, для них в дальнейшем нужно будет заменить дату парсинга на сегодняшнюю
        id_object_name_at_pages = []  # переменная, в которую будут записываться номера объявлений с сайта, в дальнейшем будут находится номера объявлений в екселе, которых уже нет на сайте (завершенные) и для них будет рассчитана экспозиция
        for baseurl in self.baseurls: # парсинг осуществляется для всех базовых урлов в списке (по типу недвижимости)
            till_page = self.parser_worker.get_till_page(baseurl)  # находим последнюю страницу в базовом урле
            print('The last page in ', baseurl, " is", till_page)
            page_projects_in_excel = []  # переменная, в которую будут записываться (номер строки в ексель, номер объявления, дата обновления) со СТРАНИЦЫ сайта, которые уже есть в екселе, для них в дальнейшем нужно будет заменить дату парсинга на сегодняшнюю
            id_object_name_at_page = []  # переменная, в которую будут записываться номера объявлений со СТРАНИЦЫ сайта
            for page_num in range(self.from_page, till_page+1): # проходимся по каждой странице
                if page_num == 1:  # Парсим первую страницу (базовый урл)
                    print("\nParsing page №: ", page_num)
                    page_projects_in_excel, id_object_name_at_page = self.parse_page(baseurl, excel_objects)
                else:  # Парсим все остальные страницы кроме(первой)
                    print("\nParsing page №: ", page_num)
                    page_url = "{}?page={}".format(baseurl, page_num - 1)
                    try:
                        page_projects_in_excel, id_object_name_at_page = self.parse_page(page_url, excel_objects)  # Парсим каждую страницу (в ней записываем новые объекты в ексель), а возвращаем объявления, которые уже есть в екселе и номера объхявлений со старницы
                    except:
                        print('ERROR')
                        with open("mistakes.txt", "a") as myfile:
                            myfile.write('ERROR in parse page {} {}\n'.format(Exception.__class__.__name__, TypeError, page_num))
                all_projects_in_excel.extend(page_projects_in_excel)
                id_object_name_at_pages.extend(id_object_name_at_page)
                waiting_time = random.randint(1, 10)
                print("Waiting time is {}".format(waiting_time))
                time.sleep(waiting_time)
                with open("object_on_realt.txt", "w") as myfile:
                    myfile.write(str(all_projects_in_excel))
                with open("id_object_on_realt.txt", "w") as myfile:
                    myfile.write(str(id_object_name_at_pages))


        #  Для последующих парсингов (не для первого)
        todays_date = self.parser_worker.get_today_date()  # Получаем сегодняшнюю дату парсинга

        # Для завершенных объявлений: СЕЙЧАС НЕ ИСПОЛЬЗУЕТСЯ
        # print("\nCalculating Expozition and update Parsing date for closed objects ... ")
        # self.excel_worker.write_expozition(excel_objects, id_object_name_at_pages, self.expozition_col_name, todays_date)  # для объявлений номеров из ексель которых уже нет на сайте записывается экспозиция и заменяется дата парсинга на сегодняшнюю


if __name__ == "__main__":
    realt_parser = Runner()
    realt_parser.parse_pages()

    # with open("текст ОД/excel_objects.txt", 'r') as mf:
    #     excel_objects = literal_eval(mf.read())
    # with open("текст ОД/id_object_on_realt.txt", 'r') as mf:
    #     id_object_name_at_pages =literal_eval(mf.read())
    # with open("текст ОД/object_on_realt.txt", 'r') as mf:
    #     all_projects_in_excel = literal_eval(mf.read())
    #
    # todays_date = realt_parser.parser_worker.get_today_date()
    # print('todas day is', todays_date)
    #
    # print(excel_objects, type(excel_objects))
    #
    # realt_parser.excel_worker.rewrite_parsing_date(all_projects_in_excel, todays_date)
    #
    # realt_parser.excel_worker.write_expozition(excel_objects, id_object_name_at_pages, realt_parser.expozition_col_name, todays_date)


