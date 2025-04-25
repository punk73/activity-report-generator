from datetime import datetime
import os
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import re
import glob
# from moviepy.editor import VideoFileClip
from moviepy import VideoFileClip
from cover import insert_dates_and_places_in_existing_table

def renderTanggal(tgl):
    result = tgl
    if len(tgl) > 10:
        tgl = tgl[:10]
    try:
        dt = datetime.strptime(tgl, "%Y-%m-%d")
    except:
        return result

    days = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']
    months = ['empty', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
    day = dt.weekday()
    month = dt.month
    hari = days[day]
    tanggal = dt.day
    bulan = months[month]
    tahun = dt.year
    result = f"{hari}, {tanggal} {bulan} {tahun}"
    return result

def generate():
    monthFolder = input("Masukkan direktori untuk dijadikan laporan: ")  # Meminta input direktori
    doc = DocxTemplate("TEMPLATE LAPORAN KEGIATAN.docx")
    context = getContent(monthFolder)  # Kirim folder sebagai argumen ke getContent

    # Loop over the context data and replace image paths with InlineImage
    for c in context:
        tmp = []
        # Loop through all images (including screenshots)
        for img in c['images']:
            tmp.append(InlineImage(doc, img, height=Mm(54)))  # Add each image
        c['images'] = tmp if tmp else ["-"]  # Use placeholder if no images

        if c['daftar_tamu']:
            dft = InlineImage(doc, c['daftar_tamu'], height=Mm(54))
            c['daftar_tamu'] = dft
        else:
            c['daftar_tamu'] = '-'

    # Render the doc with the updated context
    doc.render({"data": context})

    # Dapatkan nama folder sebagai nama bulan
    month_name = os.path.basename(monthFolder)  # Ambil nama folder sebagai nama bulan

    # Simpan dokumen dengan nama sesuai nama folder
    doc.save(f"LAPORAN KEGIATAN BALEPRASUTI SINGAPERBANGSA BULAN {month_name}.docx")
    print("Dokumen berhasil disimpan!")

def get_immediate_subdirectories(a_dir):
    res = [name for name in os.listdir(a_dir) if os.path.isdir(os.path.join(a_dir, name))]
    return sorted(res)

def getData(folder, reg, index=0, defRes=''):
    res = re.search(reg, folder)
    if res:
        res = res.group(index)
        return res.strip()
    else:
        return defRes

def getTempat(folder):
    try:
        # Split berdasarkan kata "di" tanpa memperhatikan huruf besar
        matches = re.split(r"\b[Dd][Ii]\b", folder)
        if len(matches) > 1:
            # Ambil bagian setelah "di" terakhir
            tempat = matches[-1].strip()

            # Cek apakah tempat valid
            if len(tempat) == 0 or "Kecamatan" in tempat or "aplikasi" in tempat:
                return 'Zoom Meeting'

            return tempat if tempat else 'Zoom Meeting'

        # Jika tidak ada "di", kembalikan "Zoom Meeting"
        return 'Zoom Meeting'
    except Exception as e:
        print(f"Error dalam mengambil tempat: {e}")
        return 'Zoom Meeting'

def getContent(monthFolder):
    subdir = get_immediate_subdirectories(monthFolder)

    # Tambahkan print untuk memverifikasi subdir
    print("Subfolder yang ditemukan:", subdir)

    doc_path = 'COVER.docx'
    insert_dates_and_places_in_existing_table(doc_path, monthFolder)
    res = []

    for folder in subdir:
        tanggal = getData(folder, r'\d+-\d+-\d+ (\d+.\d+.\d+.)?')
        tanggal = renderTanggal(tanggal)  # Directly use the result

        nama_acara = getData(folder, r'(?:\d{4}-\d{2}-\d{2}(?:\s+\d{2}\.\d{2}\.\d{2})?\s+)?(.+)', 1)
        tempat = getTempat(folder)  # Pastikan tempat didapatkan dari fungsi getTempat

        image_list = []
        for ext in ["jpg", "jpeg", "png", "jfif", "gif", "bmp"]:
            image_list.extend(glob.glob(os.path.join(monthFolder, folder, f'*.{ext}')))
        
        image_list = [img for img in image_list if "daftar_hadir" not in img]  # Filter images

        # Jika tidak ada gambar, cek video dan ambil screenshot
        if not image_list:
            video_filenames = [glob.glob(os.path.join(monthFolder, folder, f'*.{ext}')) for ext in ["mp4", "mkv"]]
            video_filenames = [vf for sublist in video_filenames for vf in sublist]  # Flatten the list
            
            if video_filenames:
                try:
                    video_path = video_filenames[0]
                    video_clip = VideoFileClip(video_path)

                    # Ambil screenshot di tengah dan dekat akhir video
                    video_duration = video_clip.duration
                    middle_time = video_duration / 2
                    first_time = video_duration * 0.9  # 90% dari durasi video

                    # Simpan screenshot dengan nama unik
                    screenshot_middle_path = os.path.join(monthFolder, folder, "screenshot_middle.png")
                    screenshot_first_path = os.path.join(monthFolder, folder, "screenshot_first_.png")

                    video_clip.save_frame(screenshot_middle_path, t=middle_time)
                    video_clip.save_frame(screenshot_first_path, t=first_time)

                    image_list = [screenshot_middle_path, screenshot_first_path]
                except Exception as e:
                    print(f"Error mengambil screenshot dari video: {e}")
                    image_list = []

        # Cari file gambar daftar tamu dengan nama "daftar_hadir"
        daftarTamuFileName = glob.glob(os.path.join(monthFolder, folder, "daftar_hadir.*"))
        daftarTamu = daftarTamuFileName[0] if daftarTamuFileName else False

        res.append({
            'tanggal': tanggal,
            'nama_acara': nama_acara,
            'tempat': tempat,
            'images': image_list,
            'daftar_tamu': daftarTamu  # Daftar tamu hanya mengambil gambar "daftar_hadir"
        })

    return res

generate()