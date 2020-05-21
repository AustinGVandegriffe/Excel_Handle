import numpy as np
import pandas as pd
import xlwings as xw
from openpyxl.utils.cell import get_column_letter

class App_Handle(object):
    def __init__(self, app : xw.main.App):
        self.app : xw.App = app
    ##############################################################
    # Define class handle for use of "with App_Handle as app: ..."
    def __enter__(self) -> xw.main.App:
        """
            Define initiation of class handle.
        """
        return self.app
    def __exit__(self, *args) -> None:
        """
            Define exit of class handle.
        """
        self.app.quit()
        return
    ##############################################################

import requests as rq
from bs4 import BeautifulSoup as bs

            
base_url = 'https://github.com/AustinGVandegriffe/Supplemental_Files/blob/master/Random_Names'
files_of_interest = ["first_names.txt", "middle_names.txt", "last_names.txt"]

# List to store all First, Middle, and Last names (fml)
fml_names = []
for f in files_of_interest:
    r = rq.get(f"{base_url}/{f}")
    html = bs(r.text, features="lxml")

    # Navigate to content table (inspect site HTML)
    contents = html.find("div", {"itemprop":"text"})

    # All names lie in <td> with IDs starting with LC
    # Note:
    ##  >>> lambda i: i and "LC" in i
    ##  guarentees that i in not empty and contains "LC"
    names = [t.text for t in contents.find_all("td", id = lambda i: i and "LC" in i)]

    # Store retrieved names and remove pointer
    fml_names.append(names)
    names = None

df = pd.DataFrame(
            data    = {
                    "lname" : np.random.choice(fml_names[2], size=100),
                    "fname" : np.random.choice(fml_names[0], size=100),
                    "mname" : np.random.choice(fml_names[1], size=100),
                    "empid" : np.random.choice(range(10000,20000,1), size=100, replace=False),
                    "hrrate": np.round(np.random.uniform(low=15, high=30, size=100), decimals=2),
                    "hrs"   : np.random.uniform(low=20, high=80, size=100).astype(np.int64)
                }
            # index = ["fname","mname","lname"]
    )
df.drop_duplicates(subset=["lname","fname","mname"], keep="first", inplace=True)

with App_Handle(xw.App(visible = True, add_book=False)) as xwapp:

    wb = xwapp.books.add()
    wb.activate()

    ws = wb.sheets['Sheet1']
    ws.activate()
    ws.name = "EmpWeeklyReport"

    column_names = ["First_Name","Middle_Name","Last_Name","Emp_ID","Hourly_Rate","Hours"]
    rng = xw.Range((1,1),(1,len(column_names)))
    rng.value = column_names
    rng.api.Font.Bold = True
    rng.api.Font.Size = 14
    rng.columns.autofit()

    xw.Range((2,1)).options(index=False, header=False).value = df

    wb.save("Weekly_Payroll_Report.xlsx")
    
    print("Press <Enter> to close program...")
    input()