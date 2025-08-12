import pandas as pd
import os
import jinja2
import json

latex_jinja_env = jinja2.Environment(
    block_start_string     = r'\BLOCK{',
    block_end_string       = '}',
    variable_start_string  = r'\VAR{',
    variable_end_string    = '}',
    comment_start_string   = r'\#{',
    comment_end_string     = '}',
    line_statement_prefix  = '%%',
    line_comment_prefix    = '%#',
    trim_blocks            = True,
    autoescape             = False,
    loader                 = jinja2.FileSystemLoader(os.path.abspath('.'))
)

data = {}

days = [
        {
            "name" : "Tuesday", 
             "longname": "Tuesday, September 9", 
             "slots": ["14:00--14:30", "14:30--15:00", "15:00--15:30", "15:30--16:00", "16:00--16:30", "16:30--17:00", "17:00--17:30", "18:00"]
        },
        {   
            "name" : "Wednesday", 
            "longname": "Wednesday, September 10", 
            "slots": ["9:00--9:30", "9:30--10:00", "10:00--10:30", "10:30--11:00", "11:00--11:30", "11:30--12:00", "12:00", "18:00", "19:30--21:00"]
        },
        {
            "name" : "Thursday", 
            "longname": "Thursday, September 11", 
            "slots": ["9:00--9:30", "9:30--10:00", "10:00--10:30", "10:30--11:00", "11:00--11:30", "11:30--12:00", "12:00", "14:00--14:30", "14:30--15:00", "15:00--15:30", "15:30--16:00", "16:00--16:30", "16:30--17:00", "17:00--17:30", "18:00"]
        },
        {
            "name" : "Friday", 
            "longname": "Friday, September 12", 
            "slots": ["9:00--9:30", "9:30--10:00", "10:00--10:30", "10:30--11:00", "11:00--11:30", "11:30--12:00", "12:00"]
        }
]

def checkavailable(day, time):
    res = False
    x = [d for d in days if d.get("name") == day]
    if len(x): 
        if time in x[0]["slots"]:
            res = True
    return res 

extra = [
    {"day": "Tuesday", "time": "15:30--16:00", "title": "Coffee"},
    {"day": "Tuesday", "time": "18:00", "title": "Dinner"},
    {"day": "Wednesday", "time": "10:30--11:00", "title": "Coffee"},
    {"day": "Wednesday", "time": "12:00", "title": "Lunch"},
    {"day": "Wednesday", "time": "18:00", "title": "Dinner"},
    {"day": "Wednesday", "time": "19:30--21:00", "title": "Poster session"},
    {"day": "Thursday", "time": "10:30--11:00", "title": "Coffee"},
    {"day": "Thursday", "time": "12:00", "title": "Lunch"},
    {"day": "Thursday", "time": "15:30--16:00", "title": "Coffee"},
    {"day": "Thursday", "time": "18:00", "title": "Dinner"},
    {"day": "Friday", "time": "10:30--11:00", "title": "Coffee"},
    {"day": "Friday", "time": "12:00", "title": "Lunch"}
]

for e in extra:
    if not checkavailable(e["day"], e["time"]):
        print(f"Warnung (extra): Der Slot {e["day"]}, {e["time"]} ist nicht verfügbar.")

with open("program.json", "r", encoding="utf-8") as f:
    talks = json.load(f)
    
for t in talks: 
    if not checkavailable(t["day"], t["time"]):
        print(f"Warnung (talks): Der Slot {e["day"]}, {e["time"]} ist nicht verfügbar.")
    
for d in days: 
    for slot in d["slots"]:
        if not any(t.get("day") == d["name"] and t.get("time") == slot for t in talks + extra):
            print(f"Warnung: Der Slot {d["name"]}, {slot} ist nicht belegt.")

data["abstracts"] = []
xls_path = "sprecher.xlsx"  # Datei-Pfad anpassen
df = pd.read_excel(xls_path, dtype=str).fillna("")
df = df.sort_values(by=["Name", "Vorname"], ascending=[True, True])
speakers = df.to_dict(orient="records")
for t in talks:
    s = [s for s in speakers if s["Name"] == t["lastname"] and s["Vorname"] == t["firstname"]]
    if s !=[]:
        t["affiliation"] = s[0]["Affiliation"]
        t["url"] = s[0]["URL"]
        t["mail"] = s[0]["Mail"]
    else:
        print(f"Warnung: {t["firstname"]} {t["lastname"]} nicht in der xls-Tabelle gefunden.")

for d in days:
    data["abstracts"].append({"day": d["name"], "talks" : []})
    d["programlines"] = []
    d["abstractlines"] = []
    for s in d["slots"]:
        x = [e for e in extra if e["day"] == d["name"] and e["time"] == s]
        if x != []:
            d["programlines"].append(f"{x[0]["time"]} & {x[0]["title"]}")
        x = [t for t in talks if t["day"] == d["name"] and t["time"] == s]
        if x != []:
            d["programlines"].append(f"{x[0]["time"]} &\\textbf{{{x[0]["firstname"]} {x[0]["lastname"]} }} {x[0]["title"]}")
            d["abstractlines"].append(f"\\noindent {{\\Large {x[0]["time"]}: {x[0]["title"]} }}\\\\[1ex]{{ \\large \\textbf{{ {x[0]["firstname"]} {x[0]["lastname"]}}}}}, {x[0]["affiliation"]} \\\\[2ex] {x[0]["abstract"]}")

data["program"] = days

template = latex_jinja_env.get_template(f"template.tex")
texfile = template.render(data = data)

with open("program.tex", "w", encoding="utf-8") as f:
    f.write(texfile)




