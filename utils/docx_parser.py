from docx import Document
import docx.oxml.shared
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

# --- Fungsi pengganti teks di paragraf dan tabel ---
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
    # hapus semua runs lama lalu tambahkan run baru dengan teks yang sudah diganti
    for i in range(len(paragraph.runs) - 1, -1, -1):
        paragraph._element.remove(paragraph.runs[i]._element)
    paragraph.add_run(text)

# --- Fungsi sisip paragraf bernomor ---
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

# --- Fungsi keepLines dan keepNext ---
def set_keep_with_next(paragraph, keep=True):
    pPr = paragraph._element.get_or_add_pPr()
    keepNext = OxmlElement('w:keepNext')
    if keep:
        pPr.append(keepNext)

def set_keep_lines(paragraph, keep=True):
    pPr = paragraph._element.get_or_add_pPr()
    keepLines = OxmlElement('w:keepLines')
    if keep:
        pPr.append(keepLines)

def keep_paragraphs_together(doc, start_text):
    found = False
    for p in doc.paragraphs:
        if start_text in p.text:
            found = True
        if found:
            set_keep_with_next(p)
            set_keep_lines(p)

# === FUNGSI UTAMA ===
def generate_surat_tugas(peserta, surat_info, no_template=2, template_path=None, output_path=None, header_row_count=2):
    """
    header_row_count: jumlah baris header di tabel peserta (default 2)
    """
    if template_path is None:
        if no_template == 1:
            template_path = "templates/surat_tugas_template.docx"
        elif no_template == 2:
            template_path = "templates/surat_tugas_template - Nama - NIP Combined.docx"
        else:
            raise ValueError("no_template harus bernilai 1 atau 2")

    if output_path is None:
        output_path = "output.docx"

    doc = Document(template_path)

    # Placeholder biasa
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

    # List menimbang & dasar hukum
    menimbang = surat_info.get("menimbang", [])
    if isinstance(menimbang, str):
        menimbang = [menimbang]

    dasar_hukum = surat_info.get("dasar_hukum", [])
    if isinstance(dasar_hukum, str):
        dasar_hukum = [dasar_hukum]

    if len(menimbang) == 1:
        insert_numbered_paragraphs_in_tables(doc, "{{menimbang}}", menimbang, style='Normal Font')
    else:
        insert_numbered_paragraphs_in_tables(doc, "{{menimbang}}", menimbang, style='List menimbangs')

    if len(dasar_hukum) == 1:
        insert_numbered_paragraphs_in_tables(doc, "{{dasar_hukum}}", dasar_hukum, style='Normal Font')
    else:
        insert_numbered_paragraphs_in_tables(doc, "{{dasar_hukum}}", dasar_hukum, style='List numberg')

    # === ISI TABEL PESERTA ===
    # diasumsikan tabel ke-2 adalah peserta (index 1)
    try:
        table = doc.tables[1]
    except IndexError:
        raise IndexError("Template tidak memiliki tabel kedua (index 1). Periksa template.")

    # Hapus semua baris peserta, tapi pertahankan header_row_count baris header
    if header_row_count < 1:
        header_row_count = 1
    while len(table.rows) > header_row_count:
        # selalu hapus row pada index header_row_count (yaitu row setelah header terakhir)
        table._tbl.remove(table.rows[header_row_count]._tr)

    for idx, item in enumerate(peserta, start=1):
        row_cells = table.add_row().cells
        if no_template == 1:
            # Format lama: No, Nama, NIP, Jabatan, Satker
            values = [f"{idx}.", item.nama, item.nip, item.jabatan, item.satker]
        else:
            if len(item.nip) == 18:
                formatted_nip = f"{item.nip[:8]} {item.nip[8:14]} {item.nip[14:15]} {item.nip[15:]}"
            else:
                formatted_nip = item.nip  # fallback if tidak 18 digits
            # Format baru: No, Nama\nNIP, Jabatan, Satker
            nama_nip = f"{item.nama}\nNIP. {formatted_nip}"
            values = [f"{idx}.", nama_nip, item.jabatan, item.satker]

        for i in range(min(len(row_cells), len(values))):
            # Isi teks cell
            # untuk menjaga formatting sederhana, kosongkan runs lama lalu tambah run baru
            cell_para = row_cells[i].paragraphs[0]
            # clear existing runs
            for r in cell_para.runs:
                cell_para._element.remove(r._element)
            cell_para.add_run(values[i])

            # Jika kolom pertama (No), set alignment center
            if i == 0:
                try:
                    cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception:
                    pass

    # Tandai header tabel: baris 0 .. header_row_count-1
    # hanya lakukan jika tabel saat ini memiliki minimal header_row_count baris
    if len(table.rows) >= header_row_count:
        for header_row_idx in range(header_row_count):
            tr = table.rows[header_row_idx]._tr
            # cari trPr jika ada
            trPr_list = tr.xpath('./w:trPr')
            if not trPr_list:
                # buat lalu append
                trPr_element = OxmlElement('w:trPr')
                tr.append(trPr_element)
            else:
                trPr_element = trPr_list[0]
            # append tblHeader jika belum ada
            if not trPr_element.xpath('./w:tblHeader'):
                trPr_element.append(docx.oxml.shared.OxmlElement('w:tblHeader'))

            # Set header kolom No (kolom pertama) agar center juga
            try:
                header_cell_para = table.rows[header_row_idx].cells[0].paragraphs[0]
                # clear runs then set text alignment (jika ingin mempertahankan teks header, jangan hapus runs)
                # di sini hanya atur alignment
                header_cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                pass

    # Keep paragraph setelah "Untuk"
    keep_paragraphs_together(doc, "Untuk")

    # Simpan dokumen
    doc.save(output_path)
    return output_path
