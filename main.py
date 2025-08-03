from http.server import HTTPServer, SimpleHTTPRequestHandler

import configargparse
from dotenv import load_dotenv
import datetime
import collections
import pandas
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape


def calculate_the_age(current_year):
    year_of_foundation = 1920
    age_of_the_company = current_year - year_of_foundation
    return age_of_the_company


def change_the_word_year(age):
    if 11 <= age % 100 <= 20:
        return f"{age} лет"
    elif age % 10 == 1:
        return f"{age} год"
    elif 2 <= age % 10 <= 4:
        return f"{age} года"
    else:
        return f"{age} лет"


def create_parser():
    parser = configargparse.ArgParser()

    parser.add_argument(
        '-p', '--excel_file_path',
        env_var='WINE_EXCEL_PATH',
        default='wine3.xlsx',
        type=Path
    )

    return parser


def main():
    load_dotenv()
    parser = create_parser()
    options = parser.parse_args()

    wines_df = pandas.read_excel(
        options.excel_file_path,
        sheet_name='Лист1',
        keep_default_na=False
        )

    sorted_wines_df = wines_df.sort_values('Категория')

    sorted_wines_by_category = collections.defaultdict(list)

    for sort in sorted_wines_df.to_dict(orient='records'):
        sorted_wines_by_category[sort['Категория']].append(sort)

    groups_of_wines = []
    for group in sorted_wines_by_category.values():
        groups_of_wines.extend(group)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    current_year = datetime.datetime.now().year

    rendered_page = template.render(
        sorted_wines_by_category=sorted_wines_by_category,
        groups_of_wines=groups_of_wines,
        age=change_the_word_year(calculate_the_age(current_year)),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
