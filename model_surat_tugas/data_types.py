from typing import List, Dict

class SuratTugasData:
    def __init__(self, no: int, nama: str, nip: str, jabatan: str, satker: str, gol: str):
        self.no = no
        self.nama = nama
        self.nip = nip
        self.jabatan = jabatan
        self.satker = satker
        self.gol = gol

def parse_surat_tugas_data(data: List[Dict]) -> List[SuratTugasData]:
    surat_tugas_list = []
    for item in data:
        surat_tugas = SuratTugasData(
            no=item.get("NO"),
            nama=item.get("NAMA"),
            nip=item.get("NIP"),
            jabatan=item.get("JABATAN"),
            satker=item.get("SATKER")
        )
        surat_tugas_list.append(surat_tugas)
    return surat_tugas_list