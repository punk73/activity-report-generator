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
    for c in context:
        tmp = []
        for img in c['images']:
            tmp.append(InlineImage(doc, img, Mm(120), Mm(80)))            
        c['images'] = tmp
        if c['daftar_tamu'] :
            dft = InlineImage(doc, c['daftar_tamu'], Mm(120), Mm(80))
            c['daftar_tamu'] = dft
        else:
            c['daftar_tamu'] = '-'

    context = {"data" : context}
    doc.render(context)
    doc.save("generated_doc.docx")

    print("Docx Generated!")


def get_immediate_subdirectories(a_dir):
    
    res = [name for name in os.listdir(a_dir)
            if os.path.isdir(os.path.join(a_dir, name))]

    return sorted(res)

def getData(folder, reg):
    res = re.search(reg, folder)
    if res:
        res = res.group(0)
        return res
    else:
        return ''

def getContent():
    monthFolder= "./agustus"
    subdir = get_immediate_subdirectories(monthFolder)
    res = []
    for folder in subdir:
        tanggal = getData(folder, '\d+-\d+-\d+ (\d+.\d+.\d+.)?')
        nama_acara = getData(folder, '[a-zA-Z\s]+') #masih harus disempurnakan
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
g = generate()
print(g)