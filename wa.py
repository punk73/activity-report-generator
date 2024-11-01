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

# chrome_options = Options()
# chrome_options.add_argument("--no-sandbox")
# chrome_options.add_argument("--disable-dev-shm-usage")

# # Enforce ARM-based ChromeDriver installation
# # service = Service(ChromeDriverManager(platform="mac_arm64").install())
# service = Service("./chromedriver-mac-arm64/chromedriver")
# driver = webdriver.Chrome(service=service, options=chrome_options)

# Cache untuk hash gambar
downloaded_hashes = set()

def sanitize_folder_name(folder_name):
    """Sanitasi nama folder agar tidak mengandung karakter invalid."""
    invalid_chars = r'\/:*?"<>|'
    return ''.join(c for c in folder_name if c not in invalid_chars).strip() or 'tanpa_keterangan'

def calculate_image_hash(image_data):
    """Menghitung hash MD5 dari byte gambar langsung."""
    try:
        image = Image.open(BytesIO(image_data)).convert('RGB')
        image = image.resize((800, 800))  # Resize agar konsisten
        image_bytes = BytesIO()
        image.save(image_bytes, format='JPEG')
        return hashlib.md5(image_bytes.getvalue()).hexdigest()
    except Exception as e:
        print(f"Gagal menghitung hash gambar: {e}")
        return None

def save_image(image_data, folder_name, image_name):
    """Simpan gambar jika belum pernah diunduh sebelumnya."""
    global downloaded_hashes

    base_folder = os.path.join(os.getcwd(), "WhatsAppImages")
    folder_path = os.path.join(base_folder, sanitize_folder_name(folder_name))
    os.makedirs(folder_path, exist_ok=True)

    image_hash = calculate_image_hash(image_data)
    if not image_hash:
        print("Gagal menghitung hash gambar. Melewatkan...")
        return

    if image_hash in downloaded_hashes:
        print("Gambar duplikat terdeteksi. Melewatkan.")
        return  # Jangan simpan jika duplikat

    try:
        image = Image.open(BytesIO(image_data)).convert('RGB')
        file_path = os.path.join(folder_path, image_name)
        image.save(file_path, "JPEG")
        downloaded_hashes.add(image_hash)  # Simpan hash gambar
        print(f"Gambar berhasil disimpan di: {file_path}")
    except Exception as e:
        print(f"Gagal menyimpan gambar: {e}")

def get_image_caption_or_next(messages, index):
    """Mencari caption dari pesan atau pesan berikutnya."""
    try:
        caption = messages[index].find_element(By.XPATH, './/span[contains(@class, "selectable-text")]').text.strip()
        if caption:
            return caption
    except:
        pass

    if index + 1 < len(messages):
        try:
            next_message = messages[index + 1]
            return next_message.find_element(By.XPATH, './/span[contains(@class, "selectable-text")]').text.strip()
        except:
            pass

    return "tanpa_keterangan"

def download_images_with_captions(contact_name):
    """Mengunduh gambar beserta caption dari chat WhatsApp."""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("user-data-dir=./profile")

    # driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    service = Service("./chromedriver-mac-arm64/chromedriver")
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        driver.get("https://web.whatsapp.com")
        input("Scan QR Code di WhatsApp Web dan tekan Enter...")

        # Cari kontak
        search_box = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
        )
        search_box.click()
        search_box.send_keys(contact_name)

        contact = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, f'//span[@title="{contact_name}"]'))
        )
        contact.click()

        last_message_ids = set()
        scrolling = True

        # Find the chat container where messages are loaded
        # x3psx0u xwib8y2 xkhd6sd xrmvbpv
        # chat_container = driver.find_element(By.XPATH, '//div[@class="copyable-area"]')
        # fifth_child = driver.find_element(By.XPATH, '//*[@id="main"]/div[5]')
        # chat_container = WebDriverWait(driver, 5).until(
        #     EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[4]' ))
        # )

        # Wait for the #main element to be present
        main_div = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "main"))
        )

        # Use JavaScript to locate the targetDiv within #main
        chat_container = driver.execute_script("""
            const mainDiv = arguments[0];

            const targetDiv = Array.from(mainDiv.querySelectorAll('div[tabindex="0"]'))
                .find(div => {
                    const style = window.getComputedStyle(div);
                    const parent = div.closest('.xnpuxes.copyable-area');
                    return style.display === 'flex' && parent;
                });

            return targetDiv;
        """, main_div)
        

        prev_height = 0
        for _ in range(10):  # Scroll 10 times max
            print(f'Scrolling attempt {_ + 1}')

            driver.execute_script("arguments[0].scrollTop = 0;", chat_container)
            # driver.execute_script("window.scrollTo({top: 900, behavior: 'smooth'});")
        
            time.sleep(10)  # Adjust for slow loads

            # Get new scroll height to check if more messages were loaded
            new_height = driver.execute_script("return arguments[0].scrollHeight;", chat_container)

            if new_height == prev_height:
                print("No new messages loaded, stopping scroll.")
                # break  # Stop if no new messages are loaded

            # print(f'ini udah masuk perulangan ke {_+1}')
            
            # while scrolling:
            messages = driver.find_elements(By.XPATH, 
                '//div[contains(@class,"message-in") or contains(@class,"message-out")]'
            )

            messageLen = len(messages)
            print(f'it has {messageLen} message')



            # driver.execute_script("arguments[0].scrollTop = 0;", chat_container)

            new_messages_found = False
            

            for index, message in enumerate(messages):
                # message_id = message.get_attribute("data-id")

                # Get the data-id from the parent div
                message_id = message.find_element(By.XPATH, './..').get_attribute("data-id")

                # message_text = message.find_element(By.XPATH, './/span[contains(@class, "selectable-text")]').text
                message_text = WebDriverWait(message, 10).until(
                    EC.presence_of_element_located((By.XPATH, './/span[contains(@class, "selectable-text")]'))
                ).text

                print(f'pesan ke {index} berisi {message_text}')

                if message_id in last_message_ids:
                    print(f'message id {message_id} sudah ada di last message ids')
                    continue

                last_message_ids.add(message_id)
                print(f'message id {message_id} belum ada')
                # new_messages_found = True

                try:
                    # image_elements = message.find_elements(By.XPATH, './/img[contains(@src, "blob:")]')
                    image_elements = WebDriverWait(message, 10).until(
                        EC.presence_of_all_elements_located((By.XPATH, './/img[contains(@src, "blob:")]'))
                    )

                    if not image_elements:
                        print(f'not image elements')
                        continue

                    caption = get_image_caption_or_next(messages, index)
                    folder_name = sanitize_folder_name(caption)

                    for image in image_elements:
                        image_src = image.get_attribute('src')

                        try:
                            base64_data = driver.execute_script('''\
                                return new Promise((resolve, reject) => {
                                    var xhr = new XMLHttpRequest();
                                    xhr.open('GET', arguments[0], true);
                                    xhr.responseType = 'blob';
                                    xhr.onload = function() {
                                        if (xhr.status === 200) {
                                            var reader = new FileReader();
                                            reader.readAsDataURL(xhr.response);
                                            reader.onloadend = function() {
                                                var base64data = reader.result.split(',')[1];
                                                resolve(base64data);
                                            };
                                        } else {
                                            reject('Gagal mengunduh blob.');
                                        }
                                    };
                                    xhr.onerror = function() {
                                        reject('Gagal mengakses blob URL.');
                                    };
                                    xhr.send();
                                });
                            ''', image_src)

                            if base64_data:
                                image_name = f"{folder_name}_{int(time.time())}.jpg"
                                save_image(base64.b64decode(base64_data), folder_name, image_name)

                        except Exception as e:
                            print(f"Error saat memproses gambar {image_src}: {e}")

                except Exception as e:
                    print(f"Error saat memproses pesan {index}: {e}")

            if not new_messages_found:
                scrolling = False  # Hentikan scrolling jika tidak ada pesan baru
            else:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(15)  # Tambahkan jeda agar pesan benar-benar muncul

    except Exception as e:
        print(f"Error saat memproses chat: {e}")
    finally:
        driver.quit()

def main():
    contact_name = "Magang comcen" #input("Masukkan nama kontak: ")
    download_images_with_captions(contact_name)

if __name__ == "__main__":
    main()