from datetime import datetime
import os
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import re
import glob

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
            tmp.append(InlineImage(doc, img, Mm(100), Mm(54)))
        if tmp :            
            c['images'] = tmp
        else :
            c['images'] = ["-"]
        if c['daftar_tamu'] :
            dft = InlineImage(doc, c['daftar_tamu'], Mm(100), Mm(54))
            c['daftar_tamu'] = dft
        else:
            c['daftar_tamu'] = '-'

    context = {"data" : context}
    print(context)
    doc.render(context)
    bulan = datetime.today().strftime("%B-%Y")
    doc.save(f"LAPORAN KEGIATAN BALEPRASUTI SINGAPERBANGSA BULAN {bulan}.docx")

    print("Docx Generated!")


def get_immediate_subdirectories(a_dir):
    
    res = [name for name in os.listdir(a_dir)
            if os.path.isdir(os.path.join(a_dir, name))]

    return sorted(res)

def getData(folder, reg):
    res = re.search(reg, folder)
    if res:
        res = res.group(0)
        return res.strip()
    else:
        return ''

def getContent():
    monthFolder= input("input directory untuk dijadikan laporan : ") # "./agustus"
    subdir = get_immediate_subdirectories(monthFolder)
    # print(subdir)
    res = []
    for folder in subdir:
        tanggal = getData(folder, '\d+-\d+-\d+ (\d+.\d+.\d+.)?')
        try:
            tanggal = datetime.strptime(tanggal, "%Y-%m-%d").strftime("%A, %d %B %Y")
        except:
            tanggal = tanggal
        # masih harus disempurnakan
        nama_acara = getData(folder, '([a-zA-Z]+\s?)+')
        tempat = getData(folder, '(?<=di).+')
        # get the image from every sub dir
        
        image_list = []
        for filename in [glob.glob(monthFolder +'/'+ folder+'/*.%s' % ext) for ext in ["jpg","png","jpeg"]]: #assuming gif
            if filename:
                image_list = filename
        
        daftarTamu = monthFolder + '/' + folder+'/daftar_tamu.png' if os.path.isfile(monthFolder + '/' + folder+'/daftar_tamu.png') else False

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