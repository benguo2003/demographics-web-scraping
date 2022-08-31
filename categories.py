import requests
from bs4 import BeautifulSoup
import openpyxl

cookies = {}
headers = {}

wb = openpyxl.load_workbook("C:\\Users\\ben.guo\\Deskt"
                            "op\\populationCollege.xlsx", data_only=True)
sheet = wb["Colleges"]

for i in range(1,5):
    college = sheet[f"A{i}"]
    print(type(college.value))
    college = college.value.replace(",", "")
    college = college.replace(" ", "-")
    college = college.replace("&", "-")
    college = college.lower()
    url = f"https://www.collegefactual.com/colleges/{college}/student-life/diversity/"
    r = requests.get(url, cookies=cookies, headers=headers)
    if r.status_code == 200:
        soup = BeautifulSoup(r.content, 'html.parser')

        # population
        temp = soup.findAll("div", attrs={"class": "quick-stats row"})
        temp2 = temp[1].find("span", attrs={"class": "stat-highlight"})
        population = temp2.text.replace(",","")

        sheet[f"E{i}"].value = population
        print(college + ": " + str(population))
        wb.save("C:\\Users\\ben.guo\\Desktop\\populationCollege.xlsx")
    else:
        print(college + " failed")
        continue










