import collections
import pandas as pd
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import argparse

def main():
    env = Environment(
        loader=FileSystemLoader(''),
        autoescape=select_autoescape(['html', 'xml'])
    )

    year = datetime.now().year - 1920
    match year % 10:
        case 1:
            form = "год"
        case 2 | 3 | 4:
            form = "года"
        case _:
            form = "лет"

    parser = argparse.ArgumentParser(description='Запуск локального сервера')
    parser.add_argument("--path_to_file", help="Путь к файлу", default="wine3.xlsx")
    file = parser.parse_args().path_to_file
    excel_data_df = pd.read_excel(file, header=None, keep_default_na=False)
    wines_list = excel_data_df.values

    wines = collections.defaultdict(list)
    params = wines_list[0]

    for wine in wines_list[1:]:
        wines[wine[0]].append({_param: _item for _param, _item in zip(params, wine)})

    template = env.get_template('template.html')
    rendered_page = template.render(
        years=year,
        year_form=form,
        wines=wines,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
