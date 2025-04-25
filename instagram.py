import os
import time
import base64
from io import BytesIO
from PIL import Image
import hashlib
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def open_instagram():
    chrome_options = Options()
    # chrome_options.add_argument("--no-sandbox")
    # chrome_options.add_argument("--disable-dev-shm-usage")
    # chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--start-maximized")
    # options.add_argument("--start-maximized")  # Start in maximized mode
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")

    chrome_options.add_argument("user-data-dir=./profile")

    # driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    service = Service("./chromedriver-mac-arm64/chromedriver")
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        driver.get("https://www.instagram.com/karawangkab.go.id/")

        # class name _aagv
        boxes =  WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="_aagv"]'))
        )

        for box in boxes:
            time.sleep(13) 
            # get the alt
            img = box.find_element(By.TAG_NAME, 'img')

            # Get the src and alt attributes
            img_src = img.get_attribute('src')
            img_alt = img.get_attribute('alt')

            # Print or process the image source and alt text
            print(f"Image URL: {img_src}")
            print(f"Alt Text: {img_alt}")

            # do screenshoot
    
    except Exception as e:
        print(f"Error: {e}")

def main():
    t = open_instagram()
    print(t)

if __name__ == "__main__":
    main()