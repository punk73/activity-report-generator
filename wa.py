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
from selenium.common.exceptions import StaleElementReferenceException

# Set untuk menyimpan hash gambar yang telah diunduh
downloaded_hashes = set()

def sanitize_folder_name(folder_name):
    """Membersihkan nama folder dari karakter tidak valid."""
    invalid_chars = r'\\/:*?"<>|'
    return ''.join(c for c in folder_name if c not in invalid_chars).strip() or 'tanpa_keterangan'

def calculate_image_hash(image_data):
    """Menghitung hash dari data gambar untuk deteksi duplikasi."""
    try:
        image = Image.open(BytesIO(image_data)).convert('RGB')
        image = image.resize((800, 800))
        image_bytes = BytesIO()
        image.save(image_bytes, format='JPEG')
        return hashlib.md5(image_bytes.getvalue()).hexdigest()
    except Exception as e:
        print(f"Gagal menghitung hash gambar: {e}")
        return None

def save_image(image_data, folder_name, image_name):
    """Menyimpan gambar ke folder yang sesuai."""
    base_folder = os.path.join(os.getcwd(), "WhatsAppImages")
    folder_path = os.path.join(base_folder, sanitize_folder_name(folder_name))
    os.makedirs(folder_path, exist_ok=True)

    image_hash = calculate_image_hash(image_data)
    if not image_hash:
        print("Gagal menghitung hash gambar. Melewatkan...")
        return

    if image_hash in downloaded_hashes:
        print("Gambar duplikat terdeteksi. Melewatkan.")
        return

    try:
        image = Image.open(BytesIO(image_data)).convert('RGB')
        file_path = os.path.join(folder_path, image_name)
        image.save(file_path, "JPEG")
        downloaded_hashes.add(image_hash)
        print(f"Gambar berhasil disimpan di: {file_path}")
    except Exception as e:
        print(f"Gagal menyimpan gambar: {e}")

def download_images_with_captions(group_or_contact_name):
    """Mengunduh gambar dan keterangan dari WhatsApp Web."""
    chrome_options = Options()
    chrome_options.add_argument("--user-data-dir=C:/path/to/your/Chrome/User/Data")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        driver.get("https://web.whatsapp.com")
        input("Scan QR Code di WhatsApp Web dan tekan Enter...")

        # Cari kontak atau grup
        search_box = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
        )
        search_box.click()
        search_box.send_keys(group_or_contact_name)

        # Klik hasil pencarian
        chat = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, f'//span[@title="{group_or_contact_name}"]'))
        )
        chat.click()

        last_message_ids = set()
        pending_images = []
        scroll_count = 0

        while scroll_count <= 10:
            try:
                messages = driver.find_elements(By.XPATH, '//div[contains(@class,"message-in") or contains(@class,"message-out")]')
                driver.execute_script("arguments[0].scrollIntoView();", messages[0])
                time.sleep(1)

                new_messages_found = False

                for index, message in enumerate(messages):
                    try:
                        message_id = message.get_attribute("data-id")
                        if message_id in last_message_ids:
                            continue

                        last_message_ids.add(message_id)
                        new_messages_found = True

                        # Proses gambar
                        image_elements = message.find_elements(By.XPATH, './/img')
                        if image_elements:
                            for image in image_elements:
                                image_src = image.get_attribute('src')
                                try:
                                    base64_data = driver.execute_script('''
                                        return new Promise((resolve, reject) => {
                                            fetch(arguments[0])
                                                .then(response => response.blob())
                                                .then(blob => {
                                                    const reader = new FileReader();
                                                    reader.onloadend = () => resolve(reader.result.split(',')[1]);
                                                    reader.onerror = () => reject('Gagal membaca blob.');
                                                    reader.readAsDataURL(blob);
                                                })
                                                .catch(() => reject('Gagal mengunduh gambar.'));
                                        });
                                    ''', image_src)

                                    if base64_data:
                                        pending_images.append(base64.b64decode(base64_data))

                                except Exception as e:
                                    print(f"Error saat memproses gambar {image_src}: {e}")

                        # Proses keterangan setelah gambar
                        if pending_images and index + 1 < len(messages):
                            caption_elements = messages[index + 1].find_elements(By.XPATH, './/span[contains(@class, "selectable-text") or contains(@class, "message-text")]')
                            if caption_elements:
                                folder_name = " ".join([element.text.strip() for element in caption_elements if element.text.strip()])
                            else:
                                folder_name = "tanpa_keterangan"

                            for image_data in pending_images:
                                image_name = f"{int(time.time())}.jpg"
                                save_image(image_data, folder_name, image_name)

                            pending_images.clear()

                    except StaleElementReferenceException:
                        print("Elemen pesan berubah, mencoba lagi...")
                        break

            except Exception as e:
                print(f"Error saat memproses chat: {e}")

            if not new_messages_found:
                scroll_count += 1
                time.sleep(1)
            else:
                scroll_count = 0

    except Exception as e:
        print(f"Error saat memproses chat: {e}")
    finally:
        driver.quit()

def main():
    group_or_contact_name = input("Masukan Nama Kontak Atau Grup: ")
    download_images_with_captions(group_or_contact_name)

if __name__ == "__main__":
    main()
