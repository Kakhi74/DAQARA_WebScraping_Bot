from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
from lxml import etree
import time
import pandas as pd


field_1 = []
No_ferme = []
Nom_ferme = []
Responsable = []
Rue = []
Municipalite = []
Province = []
Code_postal = []
Year = []
No_champ = []
Nom_champ = []
No_point = []
Nom_point = []
Etat = []
Date = []
Nom_du_point = []
Nom_de_l__ = []
Longitude = []
Latitude = []
Commentaire = []
Code_QR = []
FID = []


driver = webdriver.Chrome(service=Service("C:\Program Files (x86)\chromedriver.exe"))
driver.get("https://www.arcgis.com/home/webmap/viewer.html?webmap=249640db0015458f96dd442180e98dde&extent=-74.6875,46.1045,-74.6558,46.1168")
driver.maximize_window()

time.sleep(5)

button1 = driver.find_element(By.XPATH, "//*[@id='legendContentButtons']/div[1]/span[2]")

time.sleep(5)

print("\nAccessing Soil_Sample grid ...\n")
ActionChains(driver).move_to_element(button1).click(button1).perform()

time.sleep(5)

button2 = driver.find_element(By.XPATH, "//*[@id='Soil_sample_GPS_LCC_shapefile_5155_tableTool']/span")

time.sleep(5)

print("Connection Established ...")
ActionChains(driver).move_to_element(button2).click(button2).perform()

time.sleep(5)

counter = 0
verical_ordinate = 1000
scroller = driver.find_element(By.XPATH, "//*[@id='dgrid_0']/div[2]")

while True:
    try:
        driver.find_element(By.XPATH, f"//*[@id='dgrid_0-row-{counter}']/table/tr/td[1]/div")
    except:
        print("\nWARNING")
        print("\n[Could not find data on the current HTML file]\n")
        print("\nStarting scrolling mode ...\n")
        while True:
            print("     ... Scrolling down ...\n")
            driver.execute_script("arguments[0].scrollTop = arguments[1]", scroller, verical_ordinate)
            verical_ordinate += 1000
            time.sleep(5)
            try:
                driver.find_element(By.XPATH, f"//*[@id='dgrid_0-row-{counter}']/table/tr/td[1]/div")
            except:
                continue
            else:
                break
    finally:
        print("        DATA FOUND")
        time.sleep(0.5)
        content = driver.page_source
        soup = BeautifulSoup(content, "html.parser")
        dom = etree.HTML(str(soup))
        data = []
        print("     ... Retrieving data ...")
        for table in range(1, 23):
            xpath_address = f"//*[@id='dgrid_0-row-{str(counter)}']/table/tr/td[{str(table)}]/div"
            webdata = dom.xpath(xpath_address)[0].text
            data.append(webdata)

        field_1.append(data[0])
        No_ferme.append(data[1])
        Nom_ferme.append(data[2])
        Responsable.append(data[3])
        Rue.append(data[4])
        Municipalite.append(data[5])
        Province.append(data[6])
        Code_postal.append(data[7])
        Year.append(data[8])
        No_champ.append(data[9])
        Nom_champ.append(data[10])
        No_point.append(data[11])
        Nom_point.append(data[12])
        Etat.append(data[13])
        Date.append(data[14])
        Nom_du_point.append(data[15])
        Nom_de_l__.append(data[16])
        Longitude.append(data[17])
        Latitude.append(data[18])
        Commentaire.append(data[19])
        Code_QR.append(data[20])
        FID.append(data[21])

        counter += 1

    if counter == 1387:
        print("\n\n DATA WAS SUCCESSFULLY COLLECTED !\n\n")
        break

driver.close()

websitedata = {
    "field_1":field_1,
    "No_ferme":No_ferme,
    "Nom_ferme":Nom_ferme,
    "Responsable":Responsable,
    "Rue":Rue,
    "Municipalite":Municipalite,
    "Province":Province,
    "Code_postal":Code_postal,
    "Year":Year,
    "No_champ":Nom_champ,
    "Nom_champ":Nom_champ,
    "No_point":No_point,
    "Nom_point":Nom_point,
    "Etat":Etat,
    "Date":Date,
    "Nom_du_point":Nom_du_point,
    "Nom_de_l__":Nom_de_l__,
    "Longitude":Longitude,
    "Latitude":Latitude,
    "Commentaire":Commentaire,
    "Code_QR":Code_QR,
    "FID":FID
}

print("Creating Dataframe ...\n")
df = pd.DataFrame(websitedata)

print("Exporting Dataframe to Excel format as ---> [ArcGIS_data.xlsx]\n")
df.to_excel('ArcGIS_data.xlsx')

print("--- Program executed without any issue ---/n")
print()
print("Karim Khiari")
print("Ulaval 2022")