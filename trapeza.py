#!/usr/bin/python3

import re
import sys
import xlrd
import copy
import datetime
import argparse
import urllib.request


def warning(message):
    warning.list.append(message)
    sys.stderr.write("WARNING: " + str(message) + "\n\n")
warning.list = []


class Item(dict):
    """ 
    Menu item
    
    """
    def __init__(self, name, size, price, category=None):
        super().__init__()
        self['name'] = name
        self['size'] = size
        self['price'] = price
        self['category'] = category
        
    def __str__(self):
        return "%s | %s | %dр." % (self['name'], self['size'], self['price'])
        
class Menu(list):
    """ 
    Menu for one day
    
    """
    def __init__(self, date=None):
        super().__init__()
        self.date = date
        
    def categoris(self):
        """
        Return list of all categoris in menu
        
        """
        return sorted(list(set(i['category'] for i in self)))
    
    def compex_by_price(self, price):
        """
        Return complex dinner item by price
        
        """
        complexs =  [i for i in self if 'комплексные обеды' in i['category']]
        for c in complexs:
            if c['price'] == price:
                return c
            
    def submenu(self, category):
        """
        Return menu list by category
        
        """
        return [i for i in self if category in i['category']]
    
    def find(self, value):
        # try find price at the end and remove from serach pattern
        with_price = re.search(r'(\D+)\s+(\d+)',value)
        price = None
        if with_price:
            value = with_price.groups()[0] # only string with no price
            price = int(with_price.groups()[1])
        for i in self:
            if i['name'].lower().startswith(value.lower()):
                if price and price != i['price']:
                    warning("Предполагаемая цена %d блюда '%s' не равна цене блюда из меню: %s" % (price, value, str(i)))
                return i
        
        
class ParseDateException(Exception):
    pass
        
def parse_date(s):
    """
    Parse date string by format "DD month_name YYYY"
    
    """
    months = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря']
    g = re.search(r"(\d+)\s+(\w+)\s+(\d+)", s)
    if g is None:
        raise ParseDateException("В строке '%s' не удалось найти сопоставлению даты 'DD название_месяца YYYY'" % s)
    grp = g.groups()
    if grp[1].lower() not in months:
        raise ParseDateException("Месяц '%s' не удалось найти в " % grp[1].lower(), months)
    return datetime.date(int(grp[2]),months.index(grp[1].lower()) + 1, int(grp[0]))
        

def read_all(file_path):
    """
    Read XLS and return generator throw the rows
    
    """
    rb = xlrd.open_workbook(file_path,formatting_info=True)
    sheet = rb.sheet_by_index(0)
    return (sheet.row_values(rownum) for rownum in range(sheet.nrows))

def parse_menu(file_path):
    """
    Parse XML and return menu dict {date: Menu}
    
    """
    rows = read_all(file_path)
    menu = {}
    current_date = None
    current_category = None
    for row in rows:
        #print(row[0])
        if "меню" in str(row[0]).lower():
            current_date = parse_date(rows.__next__()[0])
            rows.__next__() # skip phones
            rows.__next__() # skip Наименование
            menu[current_date] = Menu(current_date)
            continue
            
        if current_date is not None:
            if row[1] == '' or row[1] == None: # detect category by size less in cell 2
                current_category = row[0].lower()
            elif row[1] != '' and row[1] is not None and row[0] != '' and row[0] is not None:
                menu[current_date].append(Item(row[0], row[1], int(re.search(r"\d+", str(row[2])).group()), current_category))
    return menu

# === Order functions ===

class Order(Menu):
    """
    Order list

    """
    def __init__(self, date=None):
        super().__init__(date)

    def append(self, item, owner=None):
        super().append(copy.copy(item))
        self[-1]['owner'] = owner

    def print_by_owner(self):
        print("Заказ на %s " % self.date.strftime("%d.%m.%y"))
        owner = None
        total = 0
        for o in self:
            if o['owner'] != owner:
                if owner is not None:
                    print("Итого %s: " % owner, total)
                owner = o['owner']
                total = 0
                print(owner + ":")
            total += o['price']
            print(o.__str__())
        print("Итого %s: " % owner, total)


class OrderCell:
    """
    Represent order cell in order file

    """
    def __init__(self, price=None, comments=None, owner=None):
        self.price = int(price)
        self.comments = comments
        self.owner = owner

    def __str__(self):
        return "%s: %d р. на: %s" % (self.owner, self.price, self.comments)


def parse_comments(cell, order, menu):
    """
    Разбиваем комментарий по строкам и пробуем найти в меню

    """
    total_price = 0
    for row in cell.comments.split("\n"):
        item = menu.find(re.sub(r'(^\s+)|(\s+$)', '', row).lower())
        if item is None:
            warning(str(cell.owner) + ": Для строки %s блюдо не найдено")
        else:
            order.append(item, cell.owner)
            total_price += item['price']
    if total_price != cell.price:
        warning(str(cell.owner) + ": итоговая сумма %d р. не равна сумме ячейки %d р." % (total_price, cell.price))


def parse_order(cell, order, menu):
    """
    Парсим заказ из одной ячейки

    """
    if cell.price is None or cell.price == 0:
        warning(str(cell.owner) + " ничего не заказал!")

    elif cell.comments is None:
        complex = menu.compex_by_price(int(cell.price))
        if complex is None:
            warning(str(cell.owner) + " за %d рублей комплекс не найден" % cell.price)
        else:
            order.append(complex, cell.owner)

    else:
        parse_comments(cell, order, menu)


def parse_order_file(file_path):
    """
    Возвращает список ячеек с заказами

    """
    return [OrderCell(0), OrderCell(209, owner="Иванов"), OrderCell(100, comments="Салат 100",owner="Петров")]


def parse_order_list(order_cells, menu):
    """
    Парсим список ячеек заказов и создаем общий заказ

    """
    order = Order(menu.date)
    for cell in order_cells:
        parse_order(cell, order, menu)

    return order


            
# === Service functions ===

def valid_date(s):
    try:
        return datetime.datetime.strptime(s, "%d.%m.%y").date()
    except ValueError:
        msg = "For date option not a valid date: '{0}'. Should be DD.MM.YY".format(s)
        raise argparse.ArgumentTypeError(msg)

def menu_file_name(date):
    start = date - datetime.timedelta(days=date.weekday())
    end = start + datetime.timedelta(days=6)
    return start.strftime("%d.%m") + '-' + end.strftime("%d.%m") + ".xls"

def get_menu(menu, date):
    menu_name = menu_file_name(date)
    url = "http://www.dobraja-trapeza.ru/menju/%s" % menu_name
    if menu.lower().startswith("http"):
        url = menu
    if menu == "download" or menu.lower().startswith("http"):
        print("Download menu from %s" % url)
        urllib.request.urlretrieve(url, menu_name)
        return menu_name 
    return menu


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Automotion dinner order generator")

    parser.add_argument('-d', "--date", 
                    help="The date for order in format DD.MM.YY", 
                    default=(datetime.date.today() + datetime.timedelta(days=1)), # tomorrow
                    type=valid_date)    
    parser.add_argument("-m","--menu", type=str, default = "download", 
                        help="Path or url(http start) for menu download. Try to download automatical if not specified")

    parser.add_argument("-o","--order", type=str, default = "order", 
                        help="Path order file")


    args = parser.parse_args()
    
    menu = parse_menu(get_menu(args.menu, args.date))
    order = parse_order_list(parse_order_file(args.order), menu[args.date])
    order.print_by_owner()



