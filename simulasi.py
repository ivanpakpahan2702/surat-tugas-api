import json
from utils.docx_parser import generate_surat_tugas
from model_surat_tugas.data_types import parse_surat_tugas_data
from docx2pdf import convert

with open("data_simulasi.json") as f:
    data = json.load(f)

# with open("data_simulasi_1.json") as f:
#     data = json.load(f)

def docx_to_pdf(docx_path, pdf_path):
    convert(docx_path, pdf_path)

peserta_objs = parse_surat_tugas_data(data["peserta"])
surat_info = {k: data[k] for k in data if k != "peserta"}

output_file = generate_surat_tugas(peserta_objs, surat_info)
pdf_file = output_file.replace(".docx", ".pdf")
docx_to_pdf(output_file, pdf_file)

print("Surat tugas berhasil dibuat.")