import pandas as pd
import os
import jinja2
from bs4 import BeautifulSoup

latex_jinja_env = jinja2.Environment(
    block_start_string = '\BLOCK{',
    block_end_string = '}',
    variable_start_string = '\VAR{',
    variable_end_string = '}',
    comment_start_string = '\#{',
    comment_end_string = '}',
    line_statement_prefix = '%%',
    line_comment_prefix = '%#',
    trim_blocks = True,
    autoescape = False,
    loader = jinja2.FileSystemLoader(os.path.abspath('.'))
)

data = {}

template = latex_jinja_env.get_template(f"static/template.tex")
texfile = template.render(data = data)



def xls_to_html_list(file_path):
    # Excel-Datei einlesen
    df = pd.read_excel(file_path, dtype=str).fillna("")
    df = df.sort_values(by=["Name", "Vorname"], ascending=[True, True])

    # Überprüfen, ob die benötigten Spalten vorhanden sind
    required_columns = {"Name", "Vorname", "Affiliation", "URL", "Zusage"}
    if not required_columns.issubset(df.columns):
        raise ValueError(f"Die Datei muss die Spalten {required_columns} enthalten.")
    
    # HTML-Liste generieren
    ul = BeautifulSoup(features="html.parser").new_tag("ul", **{"class": "list-group"})
    
    for _, row in df.iterrows():
        if row["Zusage"].strip().lower() != "ja":
            continue

        name = row["Name"]
        vorname = row["Vorname"]
        affiliation = row["Affiliation"]
        url = row["URL"]
        
        li = BeautifulSoup(features="html.parser").new_tag("li", **{"class": "list-group-item"})
        if url != "":
            link = f'<a href="{url}" target="_blank">{vorname} {name}</a>'
        else:
            link = f'{vorname} {name}'
        if affiliation != "":
            link = link + f" ({affiliation})"
        li.append(BeautifulSoup(link, "html.parser"))
        ul.append(li)
    
    # HTML-Dokument formatieren
    html = BeautifulSoup(features="html.parser")
    html.append(ul)
    return html.prettify()

if __name__ == "__main__":
    file_path = "sprecher.xlsx"  # Datei-Pfad anpassen
    html_output = xls_to_html_list(file_path)
    
    with open("sprecher.html", "w", encoding="utf-8") as f:
        f.write(html_output)
    
    print("HTML-Datei erfolgreich erstellt: sprecher.html")
