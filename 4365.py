from time import sleep
from selenium import webdriver
import collections
import pandas as pd
import xlwt

HEADERS = [
    'Название',
    'Категория',
    'Тип',
    'Ссылки на изображение',
    'Рейтинг заведения',
    'Номер телефона',
    'Адрес',
    'Метро',
    'Средняя цена',
    'Рабочее время',
    'Сайт',
    'Описание',
]


class Building:
    def __init__(self, title, description, categories, rating, phone, address, list_metro,
                 avg_price, work_time, images, website, types):
        self.title = title
        self.description = description
        self.categories = categories
        self.rating = rating
        self.phone = phone
        self.address = address
        self.list_metro = list_metro
        self.avg_price = avg_price
        self.work_time = work_time
        self.images = images
        self.website = website
        self.types = types


Main_URL = 'https://zoon.ru/'


class Client:

    def __init__(self):
        self.driver = webdriver.Chrome()
        self.driver.get(Main_URL)
        self.result = []
        self.categories = self.find_categories()
        self.point = 0
        self.wb = xlwt.Workbook()
        self.ws = self.wb.add_sheet('1', cell_overwrite_ok=True)
        self.l = 1


    def get_url(self, URL):
        self.driver.get(URL)

    def use_button(self, URL):
        self.get_url(URL)
        while True:
            try:
                button = self.driver.find_element_by_class_name(
                    'js-next-page.button.button-show-more.button-block.button40.button-primary')
                button.click()
                sleep(10)
            except:
                break

    def find_all_urls(self, URL):
        self.use_button(URL)
        urls = self.driver.find_elements_by_class_name('js-item-url')
        list = []
        for url in urls:
            k = url.get_attribute('href')
            if k not in list:
                list.append(k)
        return list

    def find_categories(self):
        category_block = self.driver.find_elements_by_class_name('nav-categories__title')
        categories = []
        for item in category_block:
            t = item.text
            categories.append(t)
        categories = categories[1:-1]
        categories[5] = 'Развлечения'
        categories[6] = 'Ремонт'
        return categories

    def con_str(self, lst):
        str = ''
        for item in lst:
            str += f'{item},'
        return str[:-1]

    def find_information(self, URL):
        sleep(5)
        urls = self.find_all_urls(URL)
        for url in urls:
            self.get_url(url)
            try:
                brand_block = self.driver.find_element_by_class_name('H1.m0')
                brand = brand_block.find_element_by_tag_name('span').text
            except:
                brand = ''
            try:
                rating_block = self.driver.find_element_by_class_name('rating-value')
                rating = rating_block.text
            except:
                rating = ''
            try:
                phone_block = self.driver.find_element_by_class_name('js-phone.phoneView.phone-hidden')
                phone = phone_block.get_attribute('data-number')
            except:
                phone = ''
            try:
                address = self.driver.find_element_by_class_name('iblock').text
            except:
                address = ''
            try:
                metro_block = self.driver.find_elements_by_class_name('address-metro.invisible-links')
                list_metro = []
                for metro_ in metro_block:
                    metro = metro_.find_element_by_tag_name('a').text
                    list_metro.append(metro)
            except:
                list_metro = []
            try:
                work_time = self.driver.find_element_by_class_name('upper-first').text
            except:
                work_time = ''
            try:
                avg_price_block = self.driver.find_element_by_class_name('tooltip-target.rel')
                avg_price = avg_price_block.find_element_by_tag_name('span').text
            except:
                avg_price = ''
            try:
                website_block = self.driver.find_element_by_class_name('service-website')
                website = website_block.find_element_by_tag_name('a').get_attribute('href')
            except:
                website = ''
            try:
                description_block = self.driver.find_elements_by_class_name('description-text')
                description = ''
                for description_ in description_block:
                    description = description + ' ' + description_.text
            except:
                description = ''
            try:
                type_org = self.driver.find_elements_by_class_name('first-p')
                types = ''
                for type_ in type_org:
                    q = type_.find_elements_by_tag_name('a')
                    for a in q:
                        types = types + ',' + a.get_attribute('title')
                    types = types + '\n'
                types = types[1:]
            except:
                types = ''
            try:
                pic = self.driver.find_element_by_class_name('gallery-more')
                pic.click()
                sleep(10)
                list_pictures = []
                main_window = self.driver.find_element_by_class_name(
                    'popup.popup-photoview.js-popup-photoview.layer-content')
                pictures_window = main_window.find_element_by_tag_name('ul')
                pictures_block = pictures_window.find_elements_by_tag_name('li')
                icon = self.driver.find_element_by_class_name('s-icons-arrow-left')
                while True:
                    icon.click()
                    page = main_window.find_element_by_class_name('sign').find_element_by_tag_name('span').text
                    if page[0] == '1':
                        break
                    sleep(10)
                i = 0
                for picture_ in pictures_block:
                    i = i + 1
                    if i == 4:
                        break
                    sleep(10)
                    picture = picture_.find_element_by_tag_name('img').get_attribute('src')
                    if picture not in list_pictures:
                        list_pictures.append(picture)

            except:
                list_pictures = []

            building = Building(brand, description, self.categories[self.point], rating, phone, address,
                                self.con_str(list_metro), avg_price, work_time, self.con_str(list_pictures), website,
                                types)
            self.ws.write(self.l, 0, building.title)
            self.ws.write(self.l, 1, building.categories)
            self.ws.write(self.l, 2, building.types)
            self.ws.write(self.l, 3, building.images)
            self.ws.write(self.l, 4, building.rating)
            self.ws.write(self.l, 5, building.phone)
            self.ws.write(self.l, 6, building.address)
            self.ws.write(self.l, 7, building.list_metro)
            self.ws.write(self.l, 8, building.avg_price)
            self.ws.write(self.l, 9, building.work_time)
            self.ws.write(self.l, 10, building.website)
            self.ws.write(self.l, 11, building.description)


            self.l = self.l + 1

    def run(self, URL):
        self.find_information(URL)

    def find_web(self):
        List_urls = []
        for i in range(0, 12):
            self.ws.write(0, i, HEADERS[i])
        urls_block = self.driver.find_elements_by_class_name('nav-categories__item.js-categories-item')
        urls_block = urls_block[1:-1]
        for item in urls_block:
            url = item.find_element_by_tag_name('a').get_attribute('href')
            if url not in List_urls:
                List_urls.append(url)
        for item in List_urls:
            self.run(item)
            self.point += 1
        self.wb.save('test.xls')


client = Client()
client.find_web()
