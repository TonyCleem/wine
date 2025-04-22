import datetime
import pandas
import collections
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_needed_form(duration):
    declensions = ['год', 'года']
    needed_form = 'лет'
    remainder = duration % 10
    tens_remainder = duration % 100

    if remainder == 1 and tens_remainder != 11:
        needed_form = declensions[0]
    elif 2 <= remainder <= 4 and not (12 <= tens_remainder <= 14):
        needed_form = declensions[1]

    return needed_form


def update_range_wines(wines):
    updated_wines = collections.defaultdict(list)
    for wine in wines:
        updated_wines[wine['Категория']].append(wine)

    return updated_wines


def main():
    env = Environment(
        loader=FileSystemLoader('.'), 
        autoescape=select_autoescape(['html', 'xml'])
        )
    template = env.get_template('template.html')

    now = datetime.datetime.now()
    duration = now.year - 1920
    needed_form = get_needed_form(duration)

    excel_file = 'wines.xlsx'
    wine_from_excel = pandas.read_excel(
        excel_file, 
        sheet_name='Лист1', 
        keep_default_na=False, 
        usecols=['Категория', 'Название', 'Сорт', 'Цена', 'Картинка', 'Акция']
        )
    wines = wine_from_excel.to_dict(orient='records')
    wines = update_range_wines(wines)
    rendered_page = template.render(
        years=duration,
        years_string=needed_form,
        wines=wines
        )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
