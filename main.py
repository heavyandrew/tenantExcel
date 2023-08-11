import os
from excel import *

path = str(os.path.abspath('tenants.xlsx'))

if __name__ == "__main__":
    xl = load_workbook(path)
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    insert_new(xl, writer)

    xl = load_workbook(path)
    edit_last(xl)
    xl.save(path)