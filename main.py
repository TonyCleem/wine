import datetime
import pandas
import collections
import os
from dotenv import load_dotenv
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


FOUNDATION_DATE = 1920


def get_correct_tense_declination(total_years_together):
    declensions = ['год', 'года']
    correct_tense_declination = 'лет'

    remainder = total_years_together % 10
    tens_remainder = total_years_together % 100

    if remainder == 1 and tens_remainder != 11:
        correct_tense_declination = declensions[0]
    elif 2 <= remainder <= 4 and not (12 <= tens_remainder <= 14):
        correct_tense_declination = declensions[1]

    return correct_tense_declination


def sort_wines_by_category(wines):
    wines_by_category = collections.defaultdict(list)
    for wine in wines:
        wines_by_category[wine['Категория']].append(wine)
    return wines_by_category


def main():
    load_dotenv()

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
        )
    template = env.get_template('template.html')

    now = datetime.datetime.now()
    total_years_together = now.year - FOUNDATION_DATE
    correct_tense_declination = get_correct_tense_declination(
        total_years_together
    )

    file = os.getenv("EXCEL_FILE")
    wine_from_excel = pandas.read_excel(
        file,
        sheet_name='Лист1',
        keep_default_na=False,
        usecols=['Категория', 'Название', 'Сорт', 'Цена', 'Картинка', 'Акция']
        )
    wines = wine_from_excel.to_dict(orient='records')
    wines_by_category = sort_wines_by_category(wines)

    rendered_page = template.render(
        years=total_years_together,
        years_string=correct_tense_declination,
        wines=wines_by_category
        )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
