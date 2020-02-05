from pyscopus import Scopus
import pandas as pd

#this reads from an excel file authors.xlsx, from sheet name Sheet1 the author first and last name (columns "Last name", "First name")
# update your city accordingly

data = pd.ExcelFile("authors.xlsx")
dfs = pd.read_excel("authors.xlsx", sheet_name='Sheet1')

key = 'your scopus key here'
scopus = Scopus(key)
city = "Eindhoven"

with open('emtic-scopus.csv', encoding='utf-8', mode='w+') as file:
    for tindex, trow in dfs.iterrows():
        scopus_id = "not found"
        try:
            df = scopus.search_author("AUTHLASTNAME("+trow['Last name']+") and AUTHFIRST("+trow['First name']+") and AFFILCITY("+city+")")
            scopus_id = ""
            for index, row in df.iterrows():
                #print(row['name'], row['author_id'])
                if (scopus_id == ""):
                    scopus_id = row['author_id']
                else:
                    scopus_id += ","+row['author_id']
        except:
            scopus_id = "not found"
            pass
        last_name = trow['Last name']
        if pd.isnull(trow['First name']):
            first_name = ""
        else:
            first_name = trow['First name']
        print(last_name+" "+first_name+" "+scopus_id)
        file.write("'"+last_name+"','"+first_name+"','"+scopus_id+"'\n")
        file.flush()
    file.close()
