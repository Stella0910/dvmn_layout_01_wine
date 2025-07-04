from http.server import HTTPServer, SimpleHTTPRequestHandler

import datetime
import collections
import pandas

from jinja2 import Environment, FileSystemLoader, select_autoescape


def calculate_the_age():
    now_year = datetime.datetime.now().year
    year_of_foundation = 1920
    age_of_the_company = now_year - year_of_foundation
    return age_of_the_company



def change_the_word_year(age=calculate_the_age()):
    if 11 <= age % 100 <= 20:
        return f"{age} лет"
    elif age % 10 == 1:
        return f"{age} год"
    elif 2 <= age % 10 <= 4:
        return f"{age} года"
    else:
        return f"{age} лет"


def main():
    wines_df = pandas.read_excel('wine3.xlsx',
                                sheet_name='Лист1',
                                keep_default_na=False)

    sorted_wines_df = wines_df.sort_values('Категория')

    dict_sorted_wines = collections.defaultdict(list)

    for sort in sorted_wines_df.to_dict(orient='records'):
        dict_sorted_wines[sort['Категория']].append(sort)

    groups_of_wines = []
    for group in dict_sorted_wines.values():
        groups_of_wines.extend(group)


    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        dict_sorted_wines=dict_sorted_wines,
        groups_of_wines=groups_of_wines,
        age=change_the_word_year(calculate_the_age()),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
