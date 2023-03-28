import collections
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import pandas as pd

env = Environment(
    loader=FileSystemLoader(''),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

year = datetime.now().year - 1920
match year % 10:
    case 1:
        form = "год"
    case 2 | 3 | 4:
        form = "года"
    case _:
        form = "лет"

excel_data_df = pd.read_excel('wine3.xlsx', header=None, keep_default_na=False)
list_wines = excel_data_df.values

wines = collections.defaultdict(list)
params = list_wines[0]

for item in list_wines[1:]:
    wines[item[0]].append({p: i for p, i in zip(params, item)})

rendered_page = template.render(
    years=year,
    year_form=form,
    wines=wines,
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
