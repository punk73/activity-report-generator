from index import getContentWithoutImages
from docx import Document

def fill_existing_table(filename, data):
    doc_path = 'table_template_2.docx'
    doc = Document(doc_path)
    tables = doc.tables
    if len(tables) > 0:
        target_table = tables[0]
        
        for index, value in enumerate(data):
            row_cells = target_table.add_row().cells
            no = index +1
            row_cells[0].text = f"{no}"  # Menambahkan nomor di depan nama tempat
            row_cells[1].text = value['tanggal']
            row_cells[2].text = value['nama_acara']
            row_cells[3].text = value['tempat']

        doc.save(f'{filename}.docx')


def generate_table(filename):
    data = getContentWithoutImages(monthFolder)
    fill_existing_table(filename, data)

    print("data saved!")


# Main function
if __name__ == "__main__":
    monthFolder = input("Input path folder untuk dijadikan cover: ")
    # monthFolder = '/Volumes/192.168.93.99/DOKUMENTSI VIDCON BPS/1. LAPORAN KEGIATAN & VIDCON BALE PRASUTI SINGAPERBANGSA/2025/4. April'
    bulan = monthFolder.split('/')[-1]
    filename = f'LAPORAN TABLE BULAN {bulan}';
    generate_table(filename)
    print('finish')