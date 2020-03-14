import argparse
from datetime import datetime
from pandas import read_excel
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape

YEAR_OF_CREATION_OF_THE_WINERY = 1920


def load_drinks_from_excel_file(file_name):
    drinks_excel_table = read_excel(
        file_name,
        sheet_name='Лист1',
        na_values=['N/A', 'NA'],
        keep_default_na=False
    )
    drink_types = defaultdict(list)
    for drink in drinks_excel_table.to_dict(orient='record'):
        drink_types[drink['Категория']].append({
            'name': drink['Название'],
            'price': drink['Цена'],
            'grape_variety': drink['Сорт'],
            'image': drink['Картинка'],
            'sales': drink['Акция']
        })
    return dict(sorted(drink_types.items()))


def generate_html_template(template_file_name, drinks_file_name):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml']))
    template = env.get_template(template_file_name)
    age_of_the_winery = datetime.now().year - \
        YEAR_OF_CREATION_OF_THE_WINERY
    drink_types = load_drinks_from_excel_file(drinks_file_name)
    rendered_page = template.render(
        age_of_the_winery=age_of_the_winery,
        drink_types=drink_types
    )
    return rendered_page


def read_links_from_args():
    args_parser = argparse.ArgumentParser(description='''This script can
    load wine price list from file,
    generate html page and start simple web server''')
    args_parser.add_argument("wine_price_file", type=str,
                             help="Enter file name")
    return args_parser.parse_args().wine_price_file


def main():
    template_file_name = 'template.html'
    drinks_file_name = read_links_from_args()
    rendered_page = generate_html_template(
        template_file_name, drinks_file_name)
    drinks = load_drinks_from_excel_file(drinks_file_name)
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
