from docx import Document
import docx.oxml.shared
import os
from docx.oxml.ns import qn
import string


def replace_all_placeholders(doc, replacements):
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                inline_replace(p, key, val)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in replacements.items():
                        if key in p.text:
                            inline_replace(p, key, val)

def inline_replace(paragraph, key, val):
    text = paragraph.text.replace(key, val)
    for i in range(len(paragraph.runs) - 1, -1, -1):
        paragraph._element.remove(paragraph.runs[i]._element)
    paragraph.add_run(text)

def insert_numbered_paragraphs_in_tables(doc, placeholder, items, style):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for i, para in enumerate(cell.paragraphs):
                    for run in para.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, "")

                            # Simpan referensi paragraf placeholder yang akan dihapus
                            p_element = para._element

                            # Sisipkan semua item list
                            for item in items:
                                try:
                                    cell.add_paragraph(item, style=style)
                                except KeyError:
                                    cell.add_paragraph(item)

                            # Hapus paragraf placeholder (kosong)
                            p_element.getparent().remove(p_element)
                            return

def generate_surat_tugas(peserta, surat_info, template_path=None, output_path=None):
    if template_path is None:
        # template_path = "templates/templates.docx"
        template_path = "templates/surat_tugas_template.docx"
    if output_path is None:
        output_path = "output.docx"

    doc = Document(template_path)

    # Ganti placeholder lain (kecuali menimbang dan dasar hukum)
    replacements = {
        "{{nomor_surat_tugas}}": surat_info.get("nomor_surat_tugas", ""),
        "{{NAMA_KEGIATAN}}": surat_info.get("nama_kegiatan", ""),
        "{{TAHUN_ANGGARAN_KEGIATAN}}": surat_info.get("tahun_anggaran_kegiatan", ""),
        "{{HARI_PELAKSANAAN}}": surat_info.get("hari_pelaksanaan", ""),
        "{{TANGGAL_PELAKSANAAN}}": surat_info.get("tanggal_pelaksanaan", ""),
        "{{TEMPAT_PELAKSANAAN}}": surat_info.get("tempat_pelaksanaan", ""),
        "{{KOTA}}": surat_info.get("kota", ""),
        "{{TANGGAL}}": surat_info.get("tanggal", ""),
    }

    replace_all_placeholders(doc, replacements)

    menimbang = surat_info.get("menimbang", [])
    if isinstance(menimbang, str):
        menimbang = [menimbang]

    dasar_hukum = surat_info.get("dasar_hukum", [])
    if isinstance(dasar_hukum, str):
        dasar_hukum = [dasar_hukum]

    # Sisipkan daftar menimbang dan dasar hukum dengan style berbeda
    if len(menimbang) == 1:
        insert_numbered_paragraphs_in_tables(doc, "{{menimbang}}", menimbang, style='Normal Font')
    else:
        insert_numbered_paragraphs_in_tables(doc, "{{menimbang}}", menimbang, style='List menimbangs')

    if len(dasar_hukum) == 1:
        insert_numbered_paragraphs_in_tables(doc, "{{dasar_hukum}}", dasar_hukum, style='Normal Font')
    else:
        insert_numbered_paragraphs_in_tables(doc, "{{dasar_hukum}}", dasar_hukum, style='List numberg')

    # Isi tabel peserta
    table = doc.tables[1]
    while len(table.rows) > 1:
        table._tbl.remove(table.rows[1]._tr)

    for idx, item in enumerate(peserta, start=1):
        row_cells = table.add_row().cells
        values = [f"{idx}.", item.nama, item.nip, item.jabatan, item.satker]  # NO dengan titik
        for i in range(min(len(row_cells), len(values))):
            row_cells[i].text = values[i]

    # Tandai header tabel
    tr = table.rows[0]._tr
    trPr = tr.xpath('./w:trPr')
    if not trPr:
        trPr = docx.oxml.shared.OxmlElement('w:trPr')
        tr.append(trPr)
    trPr[0].append(docx.oxml.shared.OxmlElement('w:tblHeader'))

    doc.save(output_path)
    return output_path
