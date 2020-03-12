from datetime import datetime
from pandas import read_excel
from collections import defaultdict, OrderedDict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape

YEAR_OF_CREATION_OF_THE_WINERY = 1920


def load_drinks_from_excel_file(file_name):
    drinks_data = read_excel(
        file_name,
        sheet_name='Лист1',
        na_values=['N/A', 'NA'],
        keep_default_na=False
    )
    drink_types = defaultdict(list)
    for drink_data in drinks_data.to_dict(orient='record'):
        drink_types[drink_data['Категория']].append({
            'name': drink_data['Название'],
            'price': drink_data['Цена'],
            'grape_variety': drink_data['Сорт'],
            'image': drink_data['Картинка'],
            'sales': drink_data['Акция']
        })
    return OrderedDict(sorted(drink_types.items()))


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


def main():
    template_file_name = 'template.html'
    drinks_file_name = 'wine3.xlsx'
    rendered_page = generate_html_template(
        template_file_name, drinks_file_name)
    drinks = load_drinks_from_excel_file('wine3.xlsx')
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
