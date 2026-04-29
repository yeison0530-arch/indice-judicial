import urllib.parse
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
import time
import json

options = EdgeOptions()
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--mute-audio")
options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Edge(options=options)
url = "https://etbcsj-my.sharepoint.com/:f:/g/personal/j08cmpaliba_cendoj_ramajudicial_gov_co/IgDVrdmf5-DVSoWqBT9ky7s1AcOrR12wiq_D6dHPioVOI5U?e=viiudX"
driver.get(url)
print("Loading UI...")
time.sleep(12)

# Navigate to C01Principal
elems = driver.find_elements(By.CSS_SELECTOR, "button[name='C01Principal']")
if elems:
    print("Clicking folder...")
    driver.execute_script("arguments[0].click();", elems[0])
    time.sleep(12)
    
    print("Dumping Javascript Contexts...")
    
    # Dump window keys
    res = driver.execute_script("""
        var items = [];
        for(var key in window) {
            if(key.toLowerCase().includes('sp') || key.toLowerCase().includes('react') || key.includes('__')) {
                items.push(key);
            }
        }
        return items;
    """)
    print("Keys:", res)
    
    # Try to find elements with React Props
    react_props = driver.execute_script("""
        var objs = [];
        document.querySelectorAll('*').forEach(el => {
            Object.keys(el).forEach(k => {
                if(k.startsWith('__reactProps')) {
                    if(el.className && typeof el.className === 'string' && el.className.includes('odspSpartanList')) {
                        objs.push('Found List React Props on ' + el.tagName);
                    }
                }
            })
        });
        return objs;
    """)
    print("React Props:", react_props)

    # Dump fetch API test
    try:
        api_res = driver.execute_script("""
            var callback = arguments[arguments.length - 1];
            // Get current folder ID from URL
            var urlParams = new URLSearchParams(window.location.search);
            var folderId = urlParams.get('id');
            if(!folderId) { callback("No ID"); return; }
            
            // Try SP Web API
            fetch("/personal/j08cmpaliba_cendoj_ramajudicial_gov_co/_api/web/GetFolderByServerRelativeUrl('" + folderId + "')/Files", {
                headers: { "Accept": "application/json;odata=verbose" }
            }).then(r => r.json()).then(d => callback(JSON.stringify(d))).catch(e => callback("Error: " + e.message));
        """)
        print("API Direct Response:", api_res[:500])
    except Exception as e:
        print("API extraction failed:", e)

    # Let's also do the PAGE_DOWN key on the List itself
    print("Testing native scroll again...")
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.keys import Keys
        # Focus list element
        list_el = driver.find_element(By.CSS_SELECTOR, "div[data-automationid='DetailsList']")
        list_el.click()
        time.sleep(1)
        ActionChains(driver).send_keys(Keys.END).perform()
        time.sleep(5)
        print("Body text length after END:", len(driver.find_element(By.TAG_NAME, "body").text))
    except Exception as e:
        print("Scroll Focus Error:", e)

driver.quit()
