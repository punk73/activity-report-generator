from datetime import datetime
import os
from docxtpl import *
from docx.shared import Mm
import re
import glob
from cover import insert_dates_and_places_in_existing_table

def renderTanggal(tgl):
    # print(tgl)
    # make sure it is tanggal
    result = tgl
    if (len(tgl) > 10):
        tgl = tgl[:10]
    try:
        dt = datetime.strptime(tgl, "%Y-%m-%d")
    except:
        return result

    days = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']
    months = ['empty','Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli','Agustus','September', 'Oktober', 'November', 'Desember']
    day = dt.weekday()
    month = dt.month
    hari = days[day]
    tanggal = dt.day
    bulan = months[month]
    tahun = dt.year
    result = f"{hari}, {tanggal} {bulan} {tahun}"
    # print(result)
    return result


def generate():
    doc = DocxTemplate ("TEMPLATE LAPORAN KEGIATAN.docx")
    # context = {
    #     'data' : [
    #     { 
    #         'tanggal' : "2022-08-23",
    #         'tempat'  : "Command Center",
    #         'daftar_tamu' : "some img that we need to think later",
    #         'nama_acara' : 'materi that consist folder name',
    #         'images': 
    #         [InlineImage(doc, image_descriptor='./test.jpeg', width=Mm(120), height=Mm(80))],
    #     }
    # ]}
    monthFolder = input("input directory untuk dijadikan laporan : ") # "./agustus"
    
    new_data = {"data": getContentWithoutImages(monthFolder)}
    # new_data = {'data': [{'tanggal': 'Senin, 6 Januari 2025', 'nama_acara': 'Rapat Koordinasi Pengendalian Inflasi Daerah di Command Center', 'tempat': 'Command Center'}, {'tanggal': 'Selasa, 7 Januari 2025', 'nama_acara': 'Koordinasi Tindaklanjut Penyusunan Profil Infrastruktur Penunjang Investasi di Kabupaten Karawang', 'tempat': 'Kabupaten Karawang'}, {'tanggal': 'Rabu, 8 Januari 2025', 'nama_acara': 'Rapat penyelesaian penataan tenaga non ASN di Instansi Pemerintah DAerah', 'tempat': 'Instansi Pemerintah DAerah'}, {'tanggal': 'Kamis, 9 Januari 2025', 'nama_acara': '08-33-16 rapat koordinasi pengendalian inflasi daerah nasional (COMCEN)', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Kamis, 9 Januari 2025', 'nama_acara': 'sosialisasi teknis pengusulan NIPPPK guru', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Jumat, 10 Januari 2025', 'nama_acara': 'Sosialisasi Kamus Usulan Tahun 2026', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Senin, 13 Januari 2025', 'nama_acara': 'Validasi Status Istithaah Kesehatan Jemaah Haji Menjelang Pelunasan BPIH 2025', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Senin, 13 Januari 2025', 'nama_acara': 'rakor dengan kementan untuk kesiapan penanaman jagung serentak 1 juta hektar di lahan perkebunan guna mendukung swasembada pangan 2025', 'tempat': 'lahan perkebunan guna mendukung swasembada pangan 2025'}, {'tanggal': 'Senin, 13 Januari 2025', 'nama_acara': 'rapat koordinasi pengendalian inflasi daerah nasional', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Selasa, 14 Januari 2025', 'nama_acara': 'Orientasi Penyusunan RPJMD Kabupaten Karawang Tahun 2025-2029 dan Penyusunan RKPD Kabupaten Karawang Tahun 2026', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Rabu, 15 Januari 2025', 'nama_acara': 'Rapat Tindak Lanjut Persiapan Gala Diner Kolaborasi Pembangunan Jawa Barat', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Rabu, 15 Januari 2025', 'nama_acara': 'Pendampingan Pelaporan Penurunan Emisi GRK untuk mendukung pencapaian target RPJPD Karawang Tahun 2025 - 2045', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Senin, 20 Januari 2025', 'nama_acara': '08-06-17 Rapat Koordinasi pengendalian Inflasi daerah nasional (COMCEN).mkv', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Selasa, 21 Januari 2025', 'nama_acara': 'TINDAK LANJUT H2H TETEH SIPADI', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Rabu, 22 Januari 2025', 'nama_acara': 'Pemantauan & Validasi Status Istithaah Kesehatan Jemaah Haji Menjelang Pelunasan BPIH 2025', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Rabu, 22 Januari 2025', 'nama_acara': '13-16-09 Launching Hasil Survei Penilaian Integritas (SPI) 2024 (COMCEN)', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Kamis, 23 Januari 2025', 'nama_acara': '09-05-35 Launching Hasil Survei Penilaian Integritas (SPI) 2024 (COMCEN)', 'tempat': 'Zoom Meeting'}, {'tanggal': 'Jumat, 24 Januari 2025', 'nama_acara': 'Pertemuan Persiapan POPM Cacingan Tahun 2025', 'tempat': 'Zoom Meeting'}]}

    print(new_data)

    tableTemplate = DocxTemplate('table_template.docx')
    tableTemplate.render(new_data)
    bulan = datetime.today().strftime("%B-%Y")
    tableTemplate.save(f"TABLE KEGIATAN TEST {bulan}.docx")
    
    context = getContent(monthFolder)
    # print(context)

    for c in context:
        tmp = []
        for img in c['images']:
            if os.path.exists(img):
                tmp.append(InlineImage(doc, img, height= Mm(54)))
        if tmp :            
            c['images'] = tmp
        else :
            c['images'] = ["-"]

        if c['daftar_tamu'] :
            if os.path.exists(c.get('daftar_tamu', '')):
                dft = InlineImage(doc, c['daftar_tamu'], height= Mm(54))
                c['daftar_tamu'] = dft
        else:
            c['daftar_tamu'] = '-'
    c['page_break'] =  R('\f')
    context = {"data" : context}
    # print(new_data)
    # render nya duluan di table
    


    doc.render(context)
    doc.save(f"LAPORAN KEGIATAN BALEPRASUTI SINGAPERBANGSA BULAN {bulan}.docx")

    print("Docx Generated!")


def get_immediate_subdirectories(a_dir):
    
    res = [name for name in os.listdir(a_dir)
            if os.path.isdir(os.path.join(a_dir, name))]

    return sorted(res)

def getData(folder, reg, index=0, defRes = ''):
    res = re.search(reg, folder)
    if res:
        res = res.group(index)
        return res.strip()
    else:
        return defRes

def getContent(monthFolder):
    # monthFolder= input("input directory untuk dijadikan laporan : ") # "./agustus"
    subdir = get_immediate_subdirectories(monthFolder)

    # monthFolder = input("Input nama folder untuk dijadikan cover: ")
    doc_path = 'COVER.docx'
    # we don't really need the cover anymore since jan 2025
    # insert_dates_and_places_in_existing_table(doc_path, monthFolder)
    # print(subdir)
    res = []
    for folder in subdir:
        tanggal = getData(folder,  r'\d+-\d+-\d+ (\d+.\d+.\d+.)?')
        try:
            tanggal = renderTanggal(tanggal)
            # tanggal = datetime.strptime(tanggal, "%Y-%m-%d").strftime("%A, %d %B %Y")
        except:
            tanggal = tanggal
        # masih harus disempurnakan
        nama_acara = getData(folder, r'(?:\d{4}-\d{2}-\d{2}(?:\s+\d{2}\.\d{2}\.\d{2})?\s+)?(.+)', 1)
        # print(nama_acara)
        tempat = getData(folder, '(?<= di | Di | DI | dI ).+', defRes='Zoom Meeting')
        # get the image from every sub dir
        
        image_list = []
        for filename in [glob.glob(monthFolder +'/'+ folder+'/*.%s' % ext) for ext in ["jpg","png","jpeg"]]: #assuming gif
            if filename:
                image_list = filename

        
        daftarTamuFileName = 'daftar_hadir.png'
        daftarTamu = monthFolder + '/' + folder+'/'+ daftarTamuFileName if os.path.isfile(monthFolder + '/' + folder+'/'+daftarTamuFileName) else False

        res.append({
            'tanggal' : tanggal,
            'nama_acara' : nama_acara,
            'tempat' : tempat,
            'images' : image_list,
            'daftar_tamu' : daftarTamu
        })
    return res
    # list all sub folder here
    # for every sub folder, get tanggal, tempat, dan judul acara di foldername
    #get photo and list tamu di dalam folder tersebut.
# print(getContent())

def getContentWithoutImages(monthFolder):
    # monthFolder= input("input directory untuk dijadikan laporan : ") # "./agustus"
    subdir = get_immediate_subdirectories(monthFolder)

    # monthFolder = input("Input nama folder untuk dijadikan cover: ")
    doc_path = 'COVER.docx'
    # we don't really need the cover anymore since jan 2025
    # insert_dates_and_places_in_existing_table(doc_path, monthFolder)
    # print(subdir)
    res = []
    for folder in subdir:
        tanggal = getData(folder,  r'\d+-\d+-\d+ (\d+.\d+.\d+.)?')
        try:
            tanggal = renderTanggal(tanggal)
            # tanggal = datetime.strptime(tanggal, "%Y-%m-%d").strftime("%A, %d %B %Y")
        except:
            tanggal = tanggal
        # masih harus disempurnakan
        nama_acara = getData(folder, r'(?:\d{4}-\d{2}-\d{2}(?:\s+\d{2}\.\d{2}\.\d{2})?\s+)?(.+)', 1)
        # print(nama_acara)
        tempat = getData(folder, '(?<= di | Di | DI | dI ).+', defRes='Zoom Meeting')
        # get the image from every sub dir
        
        
        res.append({
            'tanggal' : tanggal,
            'nama_acara' : nama_acara,
            'tempat' : tempat,
        })
    return res
    # list all sub folder here
    # for every sub folder, get tanggal, tempat, dan judul acara di foldername
    #get photo and list tamu di dalam folder tersebut.


generate()