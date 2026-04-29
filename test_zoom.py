import urllib.parse
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
import time

options = EdgeOptions()
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--mute-audio")
options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Edge(options=options)
url = "https://etbcsj-my.sharepoint.com/:f:/g/personal/j08cmpaliba_cendoj_ramajudicial_gov_co/IgDVrdmf5-DVSoWqBT9ky7s1AcOrR12wiq_D6dHPioVOI5U?e=viiudX"
driver.get(url)
print("Page loading...")
time.sleep(10)

print("Zooming out to 10%...")
driver.execute_script("document.body.style.zoom = '10%'")
time.sleep(2)

elems = driver.find_elements(By.CSS_SELECTOR, "button[name='C01Principal']")
if elems:
    print("Found C01Principal, navigating...")
    driver.execute_script("arguments[0].click();", elems[0])
    time.sleep(10)
    
    print("Inside C01Principal. Setting Zoom to 10%...")
    driver.execute_script("document.body.style.zoom = '10%'")
    time.sleep(5)
    
    text = driver.find_element(By.TAG_NAME, "body").text
    lines = text.split('\n')
    archivos = [line for line in lines if "kb" in line.lower() or "mb" in line.lower() or "bytes" in line.lower()]
    print(f"Found {len(archivos)} files using ZOOM HACK.")
else:
    print("Could not find C01Principal")

driver.quit()
