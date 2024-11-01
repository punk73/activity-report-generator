import instaloader
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.drawing.image import Image as ExcelImage
import locale
from datetime import datetime
import os
import requests


# Mengatur locale ke bahasa Indonesia
# locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')

# Fungsi untuk membuat direktori jika belum ada
def create_directory_if_not_exists(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

# Fungsi untuk mengunduh gambar dari URL
def download_image(url, filename):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            with open(filename, 'wb') as f:
                f.write(response.content)
            return True
        else:
            print(f"Error mengunduh gambar: status code {response.status_code}")
            return False
    except Exception as e:
        print(f"Error saat mengunduh gambar: {e}")
        return False

# Fungsi untuk mengambil data postingan dari akun Instagram publik
def get_posts_from_account(username, month, year):
    try:
        print("Mengambil postingan dari akun Instagram...")
        L = instaloader.Instaloader(download_pictures=False, download_videos=False, quiet=True)
        # Load cookies from the file exported from Chrome in Netscape format
        # L.context.load_cookiejar('cookies.txt')
        
        # this should be 
        # L.load_session_from_file('teguh_setiawan73')

        profile = instaloader.Profile.from_username(L.context, username)
        # posts = profile.get_posts()


        posts = instaloader.Profile.from_username(L.context, "instagram").get_posts()

        SINCE = datetime(2015, 5, 1)
        UNTIL = datetime(2015, 3, 1)


        # Pastikan direktori images/ ada
        create_directory_if_not_exists('images')

        filtered_posts = []
        for post in posts:
            post_date = post.date_utc
            # Memfilter postingan sesuai bulan dan tahun
            if post_date.month == month and post_date.year == year:
                image_filename = f"images/{post.shortcode}.jpg"
                # Mengunduh gambar menggunakan URL langsung
                if download_image(post.url, image_filename):
                    post_data = {
                        'Tanggal': post_date.strftime('%Y-%m-%d'),
                        'Hari': post_date.strftime('%A'),
                        'Caption': post.caption if post.caption else '',
                        'Likes': post.likes,
                        'Comments': post.comments,
                        'IsVideo': post.is_video,  # Menyimpan informasi apakah postingan adalah video
                        'URL': f"https://www.instagram.com/p/{post.shortcode}/",
                        'Image': image_filename  # Menyimpan lokasi gambar
                    }
                    filtered_posts.append(post_data)
                else:
                    print(f"Error: Gambar tidak berhasil diunduh untuk postingan {post.shortcode}")

        print(f"Berhasil mengambil {len(filtered_posts)} postingan dari akun {username}.")
        return filtered_posts

    except Exception as e:
        print(f"Error saat mengambil data: {e}")
        return []

# Mengkonversi ukuran dari cm ke piksel (1 cm ≈ 37.8 piksel)
def cm_to_px(cm):
    return cm * 37.8

# Fungsi untuk menyimpan data postingan ke file Excel dengan kop tetap ada
def save_posts_to_new_excel(posts, template_file, output_file):
    try:
        print("Menyimpan data postingan ke file Excel...")
        # Load workbook dari file template
        workbook = load_workbook(template_file)
        sheet = workbook.active

        # Menentukan baris kosong pertama mulai dari baris ke-25
        next_row = 25
        while sheet[f'A{next_row}'].value is not None:
            next_row += 1

        # Menyimpan setiap postingan yang telah difilter ke dalam file Excel
        for post_data in posts:
            # Mengisi kolom Hari dan Tanggal dengan alignment dan bold
            sheet[f'A{next_row}'] = post_data['Hari']  # Kolom Hari
            sheet[f'A{next_row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            sheet[f'A{next_row}'].font = Font(bold=True)
            
            sheet[f'B{next_row}'] = post_data['Tanggal']  # Kolom Tanggal
            sheet[f'B{next_row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            sheet[f'B{next_row}'].font = Font(bold=True)

            # Menyimpan caption dan mengatur format
            caption_cell = sheet[f'E{next_row}']
            caption_cell.value = post_data['Caption']  # Kolom Caption
            caption_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            caption_cell.font = Font(bold=True)

            # Menentukan nilai keluaran berdasarkan apakah postingan adalah video atau foto
            if post_data['IsVideo']:  # Jika postingan adalah video
                sheet[f'F{next_row}'] = 'VIDEO'  # Kolom Keluaran
            else:
                sheet[f'F{next_row}'] = 'LEAFLET'  # Kolom Keluaran
            sheet[f'F{next_row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            sheet[f'F{next_row}'].font = Font(bold=True)

            # Menentukan apakah postingan "ORIGINAL" atau "REPOST"
            if 'repost' in post_data['Caption'].lower():  # Jika caption mengandung kata 'repost'
                sheet[f'G{next_row}'] = 'REPOST'  # Kolom Keterangan
            else:
                sheet[f'G{next_row}'] = 'ORIGINAL'  # Kolom Keterangan
            sheet[f'G{next_row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            sheet[f'G{next_row}'].font = Font(bold=True)

            # Mengisi kolom H dengan link postingan
            sheet[f'H{next_row}'] = post_data['URL']  # Kolom Link Postingan
            sheet[f'H{next_row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            sheet[f'H{next_row}'].font = Font(bold=True)

            # Memasukkan likes ke kolom Progress (kolom I)
            sheet[f'I{next_row}'] = post_data['Likes']  # Kolom Progress
            sheet[f'I{next_row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            sheet[f'I{next_row}'].font = Font(bold=True)

            # Memasukkan gambar ke kolom "Dokumentasi" (kolom D)
            if os.path.exists(post_data['Image']):
                img = ExcelImage(post_data['Image'])  # Memastikan menggunakan openpyxl.drawing.image.Image
                
                # Mengatur ukuran gambar menjadi 6 cm x 6 cm
                img.width = cm_to_px(6)  # Lebar gambar
                img.height = cm_to_px(6)  # Tinggi gambar

                # Tempatkan gambar di kolom D pada baris yang sesuai
                sheet.add_image(img, f'D{next_row}')

                # Menyesuaikan tinggi baris dengan ukuran gambar
                sheet.row_dimensions[next_row].height = img.height * 0.75
            else:
                print(f"Gambar {post_data['Image']} tidak ditemukan, tidak bisa dimasukkan ke Excel")

            next_row += 1

        # Simpan ke file output baru
        workbook.save(output_file)
        print(f"Data berhasil disimpan di {output_file}")

    except Exception as e:
        print(f"Error saat menyimpan data ke Excel: {e}")

if __name__ == "__main__":
    # Masukkan username dari akun Instagram dan bulan serta tahun yang ingin diambil
    username = input("Masukkan username Instagram: ")
    month = int(input("Masukkan bulan (angka): "))
    year = int(input("Masukkan tahun: "))

    # Template file dan output file
    template_file = "TEMPLATE LAPORAN PUBLIKASI SOSMED.xlsx"

    # Membuat objek tanggal untuk mendapatkan nama bulan dalam format bahasa Indonesia
    date_object = datetime(year, month, 1)
    month_name = date_object.strftime('%B')

    output_file = f"LAPORAN PUBLIKASI SOSMED BULAN {month_name.capitalize()}_{year}.xlsx"

    # Mengambil postingan dari akun sesuai bulan dan tahun
    posts = get_posts_from_account(username, month, year)

    if posts:
        # Simpan data postingan ke file Excel baru
        save_posts_to_new_excel(posts, template_file, output_file)
    else:
        print("Tidak ada postingan yang bisa disimpan untuk bulan ini.")