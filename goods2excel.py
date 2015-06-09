#!/usr/bin/env python3
# -*- coding: utf-8 -*-

""" Parse a goods database table and create an ms excel's workbook from it
"""

__author__ = 'remico'


from xml.etree.ElementTree import ElementTree
from abc import ABC, abstractmethod
from bs4 import BeautifulSoup, Tag
import xlsxwriter
import sys, os, glob


class IBuilder(ABC):
    @abstractmethod
    def convert_articul(self, text):
        pass
    @abstractmethod
    def convert_sizes(self, text):
        pass
    @abstractmethod
    def convert_description(self, text):
        pass
    @abstractmethod
    def convert_price(self, text):
        pass
    @abstractmethod
    def convert_price_retail(self, text):
        pass
    @abstractmethod
    def increment_row(self):
        pass


class XlsxBuilder(IBuilder):
    def __init__(self):
        self.filename = "output_.xlsx"
        self.book = xlsxwriter.Workbook(self.filename)
        self.sheet = self.book.add_worksheet("goods")
        self.fill_header()
        self.current_row = 2  # there is the header in the first row

        self.cell_format = self.book.add_format()
        self.cell_format.set_text_wrap()
        self.cell_format.set_align('vjustify')
        # self.cell_format.set_align('top')

    def fill_header(self):
        header_format = self.book.add_format()
        header_format.set_align('center')
        header_format.set_align('vcenter')
        header_format.set_bg_color('yellow')
        header_format.set_bold()

        self.sheet.write_string('A1', 'Артикул')
        self.sheet.write_string('B1', 'Описание')
        self.sheet.write_string('C1', 'Цена')
        self.sheet.write_string('D1', 'Розничная цена')
        self.sheet.write_string('E1', 'Размеры')

        self.sheet.set_column('A:A', 50)
        self.sheet.set_column('B:B', 80)
        self.sheet.set_column('C:C', 20)
        self.sheet.set_column('D:D', 20)
        self.sheet.set_column('E:E', 20)

        self.sheet.set_row(0, 25, header_format)
        self.sheet.set_default_row(35)

    def get_result(self):
        self.book.close()
        print("'%s' created" % self.filename)
        return self.book

    def increment_row(self):
        self.current_row += 1

    def convert_articul(self, text=""):
        cleantext = text.replace('&#34;', '"') if text is not None else ""
        self.sheet.write('A%d' % self.current_row, cleantext, self.cell_format)

    def convert_description(self, text=""):
        cleantext = ""
        if text is not None:
            soup = BeautifulSoup(text)

            rows = []
            # utilize the direct child objects
            for tag in soup.children:
                if not isinstance(tag, Tag):
                    continue
                # parse an html table
                if tag.name == 'table':
                    for row in tag.find_all('tr'):
                        r = '   '.join([col.get_text().strip()
                                       for col in row.find_all('td')])
                        rows.append(r)
                # parse simple html paragraphs
                else:
                    rows.append(tag.get_text().strip())
            cleantext = "\n".join(rows).strip()

        self.sheet.write('B%d' % self.current_row, cleantext, self.cell_format)

    def convert_price(self, text=""):
        self.sheet.write('C%d' % self.current_row, text, self.cell_format)

    def convert_price_retail(self, text=""):
        self.sheet.write('D%d' % self.current_row, text, self.cell_format)

    def convert_sizes(self, text=""):
        self.sheet.write('E%d' % self.current_row, text, self.cell_format)


class GoodsReader(object):
    def __init__(self, filename, IBuilder_builder):
        self.doc = ElementTree(file=filename)
        self.database = self.doc.find("database")
        if self.database is None:
            raise LookupError("It seems that the input file is not a dump of "
                              "'gloowi_goods' database table")
        print("Database: '%s'" % self.database.get("name"))
        self.builder = IBuilder_builder

    def parse_goods(self):
        goods = self.database.findall('table')
        len_ = len(goods)
        denominator_ = 20
        part_ = len_ // denominator_
        records = ({column.get('name'): column.text
                    for column in item.getiterator('column')}
                        for item in goods)
        for i, rec in enumerate(records):
            self.builder.convert_articul(rec['name'])
            self.builder.convert_description(rec['content'])
            self.builder.convert_price(rec['price'])
            self.builder.convert_price_retail(rec['price_retail'])
            self.builder.convert_sizes(rec['har_size'])
            self.builder.increment_row()
            # indicate progress
            if not i % part_:
                print('#', end='' if i < part_*denominator_ else '\n')
                sys.stdout.flush()


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: %s <xmlFile>" % (sys.argv[0],))
        sys.exit(-1)

    # clear garbage
    for file in glob.glob("output_*.xlsx"):
        os.remove(file)
        print("'%s' removed" % file)

    input_file = sys.argv[1]
    try:
        builder = XlsxBuilder()
        parser = GoodsReader(input_file, builder)
        parser.parse_goods()
    finally:
        builder.get_result()
