from datetime import datetime
import os
from docxtpl import *
from docx.shared import Mm
import re
import glob

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
    doc = DocxTemplate("TEMPLATE LAPORAN KEGIATAN.docx")
    context = {
        'data' : [
        { 
            'tanggal' : "2022-08-23",
            'tempat'  : "Command Center",
            'daftar_tamu' : "some img that we need to think later",
            'nama_acara' : 'materi that consist folder name',
            'images': 
            [InlineImage(doc, image_descriptor='./test.jpeg', width=Mm(120), height=Mm(80))],
        }
    ]}

    context = getContent()
    # print(context)
    for c in context:
        tmp = []
        for img in c['images']:
            tmp.append(InlineImage(doc, img, height= Mm(54)))
        if tmp :            
            c['images'] = tmp
        else :
            c['images'] = ["-"]
        if c['daftar_tamu'] :
            dft = InlineImage(doc, c['daftar_tamu'], height= Mm(54))
            c['daftar_tamu'] = dft
        else:
            c['daftar_tamu'] = '-'
    c['page_break'] =  R('\f')
    context = {"data" : context}
    # print(context)
    doc.render(context)
    bulan = datetime.today().strftime("%B-%Y")
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

def getContent():
    monthFolder= input("input directory untuk dijadikan laporan : ") # "./agustus"
    subdir = get_immediate_subdirectories(monthFolder)
    # print(subdir)
    res = []
    for folder in subdir:
        tanggal = getData(folder, '\d+-\d+-\d+ (\d+.\d+.\d+.)?')
        try:
            tanggal = renderTanggal(tanggal)
            # tanggal = datetime.strptime(tanggal, "%Y-%m-%d").strftime("%A, %d %B %Y")
        except:
            tanggal = tanggal
        # masih harus disempurnakan
        nama_acara = getData(folder, '(?:\d{4}-\d{2}-\d{2}(?:\s+\d{2}\.\d{2}\.\d{2})?\s+)?(.+)', 1)
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
generate()