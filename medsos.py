import openpyxl

# Memuat workbook dan sheet
workbook = openpyxl.load_workbook('LAPORAN PUBLIKASI SOSMED.xlsx')
sheet = workbook.active  # Atau sheet spesifik jika perlu

# Mengakses sel yang tergabung
merged_cell = sheet['A1']  # Misalnya A1 adalah sel yang tergabung

# Menghapus isi dari sel utama dari merged cell
main_cell = sheet.cell(row=merged_cell.row, column=merged_cell.column)
main_cell.value = None

# Simpan workbook
workbook.save('LAPORAN.xlsx')
