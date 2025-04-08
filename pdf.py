import pypandoc
import os
from tkinter import Tk
from tkinter.filedialog import askdirectory
from PyPDF2 import PdfMerger

# Fungsi untuk mengonversi DOCX ke PDF menggunakan pypandoc
def docx_to_pdf(docx_file, pdf_file):
    output = pypandoc.convert_file(docx_file, 'pdf', outputfile=pdf_file)
    assert output == ""
    print(f"Konversi {docx_file} ke {pdf_file} berhasil.")

# Membuka jendela dialog untuk memilih folder
root = Tk()
root.withdraw()  # Menyembunyikan jendela Tkinter utama
folder_path = askdirectory(title="Pilih Folder yang Berisi File DOCX")  # Membuka dialog folder

# Memeriksa apakah folder dipilih
if not folder_path:
    print("Tidak ada folder yang dipilih.")
else:
    # Membuat objek PdfMerger untuk menggabungkan file PDF nanti
    merger = PdfMerger()

    # Menyaring semua file DOCX di folder, menghindari file sementara (~$)
    docx_files = [f for f in os.listdir(folder_path) if f.endswith('.docx') and not f.startswith('~$')]

    if not docx_files:
        print("Tidak ada file DOCX di folder tersebut.")
    else:
        print(f"Menemukan file DOCX: {docx_files}")  # Menampilkan daftar file DOCX yang ditemukan
        # Mengonversi file DOCX ke PDF dan menggabungkannya
        for docx in docx_files:
            docx_path = os.path.join(folder_path, docx)  # Menyusun path lengkap file DOCX
            pdf_temp_path = os.path.join(folder_path, f"{os.path.splitext(docx)[0]}.pdf")  # Path sementara untuk file PDF

            # Mengonversi DOCX ke PDF
            docx_to_pdf(docx_path, pdf_temp_path)

            # Menambahkan PDF ke objek PdfMerger
            merger.append(pdf_temp_path)

        # Tentukan path untuk menyimpan file gabungan
        output_path = os.path.join(folder_path, "gabungan.pdf")
        
        # Menyimpan hasil gabungan ke file baru
        merger.write(output_path)

        # Menutup objek merger
        merger.close()

        # Menghapus file PDF sementara
        for docx in docx_files:
            pdf_temp_path = os.path.join(folder_path, f"{os.path.splitext(docx)[0]}.pdf")
            os.remove(pdf_temp_path)

        print(f"File gabungan disimpan di: {output_path}")

    print("Proses penggabungan selesai.")
