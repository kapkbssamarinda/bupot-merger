import os
import io
import re
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
template_dir = os.path.join(BASE_DIR, 'templates')
if not os.path.isdir(template_dir):
    template_dir = os.path.join(os.path.dirname(BASE_DIR), 'templates')

app = Flask(__name__, template_folder=template_dir)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max upload

def parse_bupot_text(text, source_name=""):
    """
    Ekstraksi data bukti potong BPPU Unifikasi format DJP dari teks PDF.
    Kolom target:
    1. Nomor
    2. Masa Pajak
    3. Sifat Pajak Penghasilan
    4. Status
    5. jenis pph
    6. kode objek pajak
    7. objek pajak
    8. dpp
    9. tarif
    10. pajak penghasilan
    11. tanggal dokumen
    12. nomor dokumen
    13. npwp
    14. nama pemotong
    15. tanggal pemotongan
    """
    # 1. Header row: NOMOR, MASA PAJAK, SIFAT, STATUS
    header_m = re.search(r'([A-Z0-9]{8,16})\s+(\d{2}-\d{4})\s+([A-Z]+)\s+([A-Z]+)', text)
    nomor = header_m.group(1) if header_m else ''
    masa_pajak = header_m.group(2) if header_m else ''
    sifat = header_m.group(3) if header_m else ''
    status = header_m.group(4) if header_m else ''

    # 2. Jenis PPh: B.2 Jenis PPh : Pasal 22
    jenis_pph_m = re.search(r'B\.2\s+Jenis\s+PPh\s*:\s*([^\n\r]+)', text)
    jenis_pph = jenis_pph_m.group(1).strip() if jenis_pph_m else ''

    # 3. Kode Objek Pajak, Objek Pajak, DPP, Tarif, Pajak Penghasilan (B.3 - B.7)
    sec_b = re.search(r'B\.3\s+B\.4\s+B\.5\s+B\.6\s+B\.7\s*([\s\S]*?)B\.8', text)
    kode_objek = ''
    objek_pajak = ''
    dpp = 0.0
    tarif = 0.0
    pph = 0.0

    if sec_b:
        b_text = sec_b.group(1).strip()
        lines = [l.strip() for l in b_text.splitlines() if l.strip()]
        if lines:
            m_row = re.search(r'^(\d{2}-\d{3}-\d{2})\s+(.*?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)$', lines[0])
            if m_row:
                kode_objek = m_row.group(1)
                objek_p1 = m_row.group(2)
                dpp_str = m_row.group(3).replace('.', '').replace(',', '.')
                tarif_str = m_row.group(4).replace(',', '.')
                pph_str = m_row.group(5).replace('.', '').replace(',', '.')
                try:
                    dpp = float(dpp_str)
                except ValueError:
                    dpp = 0.0
                try:
                    tarif = float(tarif_str)
                except ValueError:
                    tarif = 0.0
                try:
                    pph = float(pph_str)
                except ValueError:
                    pph = 0.0

                remaining_desc = ' '.join(lines[1:])
                objek_pajak = (objek_p1 + (' ' + remaining_desc if remaining_desc else '')).strip()
            else:
                objek_pajak = ' '.join(lines)

    # 4. Tanggal Dokumen: B.8 Dokumen Dasar Bukti ... Tanggal : 07 September 2026
    tgl_dok_m = re.search(r'B\.8[\s\S]*?Tanggal\s*:\s*([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})', text)
    tgl_dokumen = tgl_dok_m.group(1).strip() if tgl_dok_m else ''

    # 5. Nomor Dokumen: B.9 Nomor Dokumen : 9140424520
    no_dok_m = re.search(r'B\.9\s+Nomor\s+Dokumen\s*:\s*([^\n\r]+)', text)
    no_dokumen = no_dok_m.group(1).strip() if no_dok_m else ''

    # 6. NPWP Pemotong: C.1 NPWP / NIK : 0010611572051000
    npwp_m = re.search(r'C\.1\s+NPWP\s*/\s*NIK\s*:\s*([0-9]{15,16})', text)
    npwp = npwp_m.group(1).strip() if npwp_m else ''

    # 7. Nama Pemotong: C.3 NAMA PEMOTONG DAN/ATAU PEMUNGUT PPh : PERTAMINA PATRA NIAGA
    nama_pemotong_m = re.search(r'C\.3\s+NAMA\s+PEMOTONG[^\n:]*:\s*([^\n\r]+)', text)
    nama_pemotong = nama_pemotong_m.group(1).strip() if nama_pemotong_m else ''

    # 8. Tanggal Pemotongan: C.4 TANGGAL : 07 September 2026
    tgl_potong_m = re.search(r'C\.4\s+TANGGAL\s*:\s*([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})', text)
    tgl_pemotongan = tgl_potong_m.group(1).strip() if tgl_potong_m else ''

    return {
        'file_source': source_name,
        'Nomor': nomor,
        'Masa Pajak': masa_pajak,
        'Sifat Pajak Penghasilan': sifat,
        'Status': status,
        'jenis pph': jenis_pph,
        'kode objek pajak': kode_objek,
        'objek pajak': objek_pajak,
        'dpp': dpp,
        'tarif': tarif,
        'pajak penghasilan': pph,
        'tanggal dokumen': tgl_dokumen,
        'nomor dokumen': no_dokumen,
        'npwp': npwp,
        'nama pemotong': nama_pemotong,
        'tanggal pemotongan': tgl_pemotongan
    }

def extract_pdf_file(file_bytes_or_path, filename=""):
    try:
        with pdfplumber.open(file_bytes_or_path) as pdf:
            full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        return parse_bupot_text(full_text, source_name=filename)
    except Exception as e:
        return {
            'file_source': filename,
            'error': str(e),
            'Nomor': '',
            'Masa Pajak': '',
            'Sifat Pajak Penghasilan': '',
            'Status': 'GAGAL',
            'jenis pph': '',
            'kode objek pajak': '',
            'objek pajak': f'Gagal membaca file: {e}',
            'dpp': 0.0,
            'tarif': 0.0,
            'pajak penghasilan': 0.0,
            'tanggal dokumen': '',
            'nomor dokumen': '',
            'npwp': '',
            'nama pemotong': '',
            'tanggal pemotongan': ''
        }

@app.route('/')
@app.route('/api')
@app.route('/api/index')
@app.route('/api/index.py')
def index():
    return render_template('index.html')

class VercelPathMiddleware:
    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        path = environ.get('PATH_INFO', '')
        if path in ('/api/index.py', '/api/index', '/api', ''):
            environ['PATH_INFO'] = '/'
        elif path.startswith('/api/index.py/'):
            environ['PATH_INFO'] = path[len('/api/index.py'):]
        return self.wsgi_app(environ, start_response)

app.wsgi_app = VercelPathMiddleware(app.wsgi_app)

@app.route('/api/extract-upload', methods=['POST'])
def extract_from_upload():
    uploaded_files = request.files.getlist('files')
    if not uploaded_files or uploaded_files[0].filename == '':
        return jsonify({'success': False, 'message': 'Tidak ada file yang diunggah.'}), 400

    results = []
    for f in uploaded_files:
        if f.filename.lower().endswith('.pdf'):
            content = f.read()
            size_kb = f"{max(1, round(len(content) / 1024))} KB"
            file_stream = io.BytesIO(content)
            res = extract_pdf_file(file_stream, filename=f.filename)
            res['file_size'] = size_kb
            results.append(res)

    if not results:
        return jsonify({'success': False, 'message': 'Tidak ada file PDF valid yang diproses.'}), 400

    return jsonify({'success': True, 'data': results, 'total_files': len(results)})

@app.route('/api/export-excel', methods=['POST'])
def export_excel():
    data = request.json.get('data', [])
    if not data:
        return jsonify({'success': False, 'message': 'Data kosong, tidak dapat membuat Excel.'}), 400

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Bupot"

    # Kolom sesuai permintaan user:
    # Nomor, Masa Pajak, Sifat Pajak Penghasilan, Status, jenis pph, kode objek pajak, objek pajak,
    # dpp, tarif, pajak penghasilan, tanggal dokumen, nomor dokumen, npwp, nama pemotong, tanggal pemotongan
    headers = [
        "NOMOR",
        "MASA PAJAK",
        "SIFAT PAJAK PENGHASILAN",
        "STATUS",
        "JENIS PPH",
        "KODE OBJEK PAJAK",
        "OBJEK PAJAK",
        "DPP",
        "TARIF",
        "PAJAK PENGHASILAN",
        "TANGGAL DOKUMEN",
        "NOMOR DOKUMEN",
        "NPWP",
        "NAMA PEMOTONG",
        "TANGGAL PEMOTONGAN"
    ]

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # Tulis Header
    ws.append(headers)
    ws.row_dimensions[1].height = 26

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    # Tulis Data Baris
    row_start = 2
    for item in data:
        row_values = [
            str(item.get("Nomor", "")),
            str(item.get("Masa Pajak", "")),
            str(item.get("Sifat Pajak Penghasilan", "")),
            str(item.get("Status", "")),
            str(item.get("jenis pph", "")),
            str(item.get("kode objek pajak", "")),
            str(item.get("objek pajak", "")),
            float(item.get("dpp", 0.0) or 0.0),
            float(item.get("tarif", 0.0) or 0.0),
            float(item.get("pajak penghasilan", 0.0) or 0.0),
            str(item.get("tanggal dokumen", "")),
            str(item.get("nomor dokumen", "")),
            str(item.get("npwp", "")),
            str(item.get("nama pemotong", "")),
            str(item.get("tanggal pemotongan", ""))
        ]
        ws.append(row_values)
        current_row = ws.max_row
        ws.row_dimensions[current_row].height = 20

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.border = thin_border
            cell.font = Font(name="Calibri", size=10)

            # Formatting khusus
            if col_idx in [1, 2, 3, 4, 6, 11, 15]:  # Identitas kode & tanggal
                cell.alignment = center_align
            elif col_idx in [8, 10]:  # DPP & PPh (Numeric with thousand separator)
                cell.number_format = '#,##0'
                cell.alignment = right_align
            elif col_idx == 9:  # Tarif
                cell.number_format = '0.0"%"' if cell.value < 1 else '0.0'
                cell.alignment = right_align
            elif col_idx == 13:  # NPWP (Text format preserving leading zero)
                cell.number_format = '@'
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    # Baris Total / Summary
    last_data_row = ws.max_row
    if last_data_row >= row_start:
        total_row = last_data_row + 1
        ws.cell(row=total_row, column=1, value="TOTAL")
        ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=7)

        # Formula SUM DPP (Kolom H / 8)
        ws.cell(row=total_row, column=8, value=f"=SUM(H{row_start}:H{last_data_row})")
        # Formula SUM PPh (Kolom J / 10)
        ws.cell(row=total_row, column=10, value=f"=SUM(J{row_start}:J{last_data_row})")

        total_font = Font(name="Calibri", size=10, bold=True, color="0F172A")
        total_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        thick_top_double_bottom = Border(
            top=Side(style='thin', color='0F172A'),
            bottom=Side(style='double', color='0F172A'),
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1')
        )

        ws.row_dimensions[total_row].height = 22
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=total_row, column=col_idx)
            cell.font = total_font
            cell.fill = total_fill
            cell.border = thick_top_double_bottom
            if col_idx in [8, 10]:
                cell.number_format = '#,##0'
                cell.alignment = right_align
            elif col_idx == 1:
                cell.alignment = Alignment(horizontal="right", vertical="center")

    # Auto-fit Column Widths dengan padding
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            if cell.number_format == '#,##0' and isinstance(cell.value, (int, float)):
                val = f"{cell.value:,.0f}"
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"Rekap_Merger_Bupot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"Aplikasi Merger Bupot berjalan di http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=True)
