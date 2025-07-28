from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from utils.docx_parser import generate_surat_tugas
from model_surat_tugas.data_types import SuratTugasData
import json
import os
from io import BytesIO

app = Flask(__name__)
CORS(app, resources={
    r"/api/surat_tugas": {
        "origins": "*",
        "methods": ["POST"],
        "allow_headers": ["Content-Type"]
    }
})

@app.route('/', methods=['GET'])
def index():
    return "Hello World"

@app.route('/api/surat_tugas', methods=['POST'])
def create_surat_tugas():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data provided"}), 400

        # Validasi data wajib
        required_fields = ['nomor_surat_tugas', 'nama_kegiatan', 'peserta']
        for field in required_fields:
            if field not in data:
                return jsonify({"error": f"Field {field} is required"}), 400

        # Simpan data ke file JSON internal
        save_path = os.path.join("data", "surat_tugas_last.json")
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # Ambil data peserta
        peserta_data = data.get("peserta", [])
        
        # Validasi data peserta
        if not peserta_data:
            return jsonify({"error": "Peserta data cannot be empty"}), 400

        peserta_objs = []
        for item in peserta_data:
            try:
                peserta_objs.append(SuratTugasData(**item))
            except TypeError as e:
                return jsonify({"error": f"Invalid peserta data structure: {str(e)}"}), 400

        # Ambil data lain
        surat_info = {
            "nomor_surat_tugas": data.get("nomor_surat_tugas"),
            "nama_kegiatan": data.get("nama_kegiatan"),
            "tahun_anggaran_kegiatan": data.get("tahun_anggaran_kegiatan"),
            "hari_pelaksanaan": data.get("hari_pelaksanaan"),
            "tanggal_pelaksanaan": data.get("tanggal_pelaksanaan"),
            "tempat_pelaksanaan": data.get("tempat_pelaksanaan"),
            "kota": data.get("kota"),
            "tanggal": data.get("tanggal"),
            "menimbang": data.get("menimbang", []),
            "dasar_hukum": data.get("dasar_hukum", [])
        }

        # Generate surat tugas
        output_path = generate_surat_tugas(peserta_objs, surat_info)
        
        # Baca file dan kirim sebagai response
        with open(output_path, 'rb') as f:
            file_data = BytesIO(f.read())
        
        # Hapus file setelah dibaca (opsional)
        os.remove(output_path)
        
        return send_file(
            file_data,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='surat_tugas.docx'
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    
    
if __name__ == '__main__':
    app.run(debug=True)