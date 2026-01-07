from flask import Flask, render_template, request, send_file, jsonify
import re
import requests
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from unidecode import unidecode
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

app = Flask(__name__)

EXCEL_FILE = "nrl.xlsx"
MAX_WORKERS = 15

_cached_docs = None


def get_doc_links():
    global _cached_docs
    if _cached_docs is not None:
        return _cached_docs
    
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    
    doc_links = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.hyperlink and cell.hyperlink.target:
                link = cell.hyperlink.target
                if "docs.google.com/document" in link:
                    cell_value = str(cell.value) if cell.value else ""
                    doc_links.append({"link": link, "name": cell_value})
    
    seen = set()
    unique_docs = []
    for doc in doc_links:
        if doc["link"] not in seen:
            seen.add(doc["link"])
            unique_docs.append(doc)
    
    _cached_docs = unique_docs
    return unique_docs


def read_doc_text(url, session):
    """Đọc nội dung Google Docs với retry"""
    try:
        doc_id = re.search(r"/d/([a-zA-Z0-9_-]+)", url).group(1)
        export_url = f"https://docs.google.com/document/d/{doc_id}/export?format=txt"
        
        # Retry 2 lần nếu lỗi
        for attempt in range(2):
            try:
                r = session.get(export_url, timeout=15)
                if r.status_code == 200 and "accounts.google.com" not in r.url:
                    return r.text
            except:
                if attempt == 0:
                    continue
        return None
    except:
        return None


def normalize_text(text):
    """Chuẩn hóa text: bỏ dấu, lowercase, bỏ khoảng trắng thừa"""
    text = unidecode(text.lower())
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def find_student_in_content(content, ten_sv, mssv):
    """
    Tìm sinh viên trong nội dung với nhiều cách khác nhau
    Trả về (found, stt, nrl)
    """
    # Chuẩn hóa
    content_normalized = normalize_text(content)
    ten_normalized = normalize_text(ten_sv)
    
    # Kiểm tra có cả tên VÀ MSSV không
    has_name = ten_normalized in content_normalized
    has_mssv = mssv in content
    
    if not (has_name and has_mssv):
        return False, None, None
    
    # Tìm STT và NRL
    lines = content.split('\n')
    
    # Cách 1: Tìm theo tên
    for i, line in enumerate(lines):
        line_normalized = normalize_text(line)
        
        if ten_normalized in line_normalized:
            stt = None
            nrl = None
            
            # STT: dòng trước
            if i > 0:
                prev = lines[i-1].strip()
                if re.match(r'^\d{1,4}$', prev):  # STT thường 1-4 chữ số
                    stt = int(prev)
            
            # NRL: tìm trong 5 dòng sau
            for j in range(i+1, min(i+6, len(lines))):
                check = lines[j].strip().replace(',', '.')
                # NRL: số từ 0 đến 10, có thể có phần thập phân
                match = re.match(r'^(\d+\.?\d*)$', check)
                if match:
                    val = float(match.group(1))
                    if 0 <= val <= 10:
                        nrl = val
                        break
            
            return True, stt, nrl
    
    # Cách 2: Tìm theo MSSV (backup)
    for i, line in enumerate(lines):
        if mssv in line:
            stt = None
            nrl = None
            
            # STT: 3 dòng trước (STT -> Tên -> Lớp -> MSSV)
            if i >= 3:
                stt_line = lines[i-3].strip()
                if re.match(r'^\d{1,4}$', stt_line):
                    stt = int(stt_line)
            
            # NRL: dòng sau MSSV
            if i + 1 < len(lines):
                nrl_line = lines[i+1].strip().replace(',', '.')
                match = re.match(r'^(\d+\.?\d*)$', nrl_line)
                if match:
                    val = float(match.group(1))
                    if 0 <= val <= 10:
                        nrl = val
            
            return True, stt, nrl
    
    # Tìm thấy tên + MSSV nhưng không xác định được STT/NRL
    return True, None, None


def process_doc(doc, ten_sv, mssv, session):
    """Xử lý 1 document"""
    link = doc["link"]
    doc_name = doc["name"]
    
    try:
        content = read_doc_text(link, session)
        if content is None:
            return None
        
        found, stt, nrl = find_student_in_content(content, ten_sv, mssv)
        
        if found:
            short_name = doc_name[:50] + "..." if len(doc_name) > 50 else doc_name
            return {
                "link": link,
                "doc_name": short_name or "File",
                "stt": stt if stt else "-",
                "nrl": nrl if nrl is not None else "-",
            }
        return None
    except Exception as e:
        print(f"[ERROR] {link}: {e}")
        return None


def create_excel(results, ten_sv, mssv, total_nrl):
    output_file = f"ket_qua_{mssv}.xlsx"
    
    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "Ket qua NRL"
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    ws_out.merge_cells('A1:D1')
    ws_out['A1'] = "BÁO CÁO KẾT QUẢ ĐIỂM RÈN LUYỆN"
    ws_out['A1'].font = Font(bold=True, size=14)
    ws_out['A1'].alignment = center
    
    ws_out['A3'] = "Họ và tên:"
    ws_out['B3'] = ten_sv.upper()
    ws_out['A4'] = "MSSV:"
    ws_out['B4'] = mssv
    ws_out['A5'] = "Ngày xuất:"
    ws_out['B5'] = datetime.now().strftime('%d/%m/%Y %H:%M')
    
    ws_out['A7'] = "Số file tìm thấy:"
    ws_out['B7'] = len(results)
    ws_out['A8'] = "TỔNG ĐIỂM NRL:"
    ws_out['B8'] = total_nrl
    ws_out['A8'].font = Font(bold=True, color="0000FF")
    ws_out['B8'].font = Font(bold=True, color="0000FF")
    
    table_headers = ["#", "STT", "NRL", "Tên file", "Link"]
    for col, h in enumerate(table_headers, 1):
        cell = ws_out.cell(row=10, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    
    for idx, r in enumerate(results, 1):
        row_num = 10 + idx
        ws_out.cell(row=row_num, column=1, value=idx).alignment = center
        ws_out.cell(row=row_num, column=2, value=r["stt"]).alignment = center
        ws_out.cell(row=row_num, column=3, value=r["nrl"]).alignment = center
        ws_out.cell(row=row_num, column=4, value=r["doc_name"])
        ws_out.cell(row=row_num, column=5, value=r["link"])
        
        for col in range(1, 6):
            ws_out.cell(row=row_num, column=col).border = thin_border
    
    ws_out.column_dimensions['A'].width = 5
    ws_out.column_dimensions['B'].width = 8
    ws_out.column_dimensions['C'].width = 8
    ws_out.column_dimensions['D'].width = 50
    ws_out.column_dimensions['E'].width = 60
    
    wb_out.save(output_file)
    print(f"[INFO] Đã tạo file: {output_file}")
    return output_file


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/search', methods=['POST'])
def search():
    ten_sv = request.form.get('ten_sv', '').strip()
    mssv = request.form.get('mssv', '').strip()
    
    if not ten_sv or not mssv:
        return jsonify({"error": "Vui lòng nhập cả tên VÀ MSSV để tìm chính xác"})
    
    unique_docs = get_doc_links()
    results = []
    
    # Tạo session với connection pooling
    session = requests.Session()
    adapter = requests.adapters.HTTPAdapter(
        pool_connections=MAX_WORKERS,
        pool_maxsize=MAX_WORKERS,
        max_retries=1
    )
    session.mount('https://', adapter)
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    
    print(f"[INFO] Bắt đầu quét {len(unique_docs)} files cho {ten_sv} - {mssv}")
    
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [
            executor.submit(process_doc, doc, ten_sv, mssv, session)
            for doc in unique_docs
        ]
        
        for future in as_completed(futures):
            result = future.result()
            if result:
                results.append(result)
    
    # Sắp xếp theo STT
    results.sort(key=lambda x: (x["stt"] if isinstance(x["stt"], int) else 9999))
    
    total_nrl = sum(r["nrl"] for r in results if isinstance(r["nrl"], (int, float)))
    
    print(f"[INFO] Tìm thấy {len(results)} kết quả, tổng NRL: {total_nrl}")
    
    return jsonify({
        "results": results,
        "total_nrl": total_nrl,
        "total_files": len(results),
        "ten_sv": ten_sv,
        "mssv": mssv
    })


@app.route('/download', methods=['POST'])
def download():
    """Tạo file Excel khi user bấm tải"""
    data = request.get_json()
    results = data.get('results', [])
    ten_sv = data.get('ten_sv', '')
    mssv = data.get('mssv', '')
    total_nrl = data.get('total_nrl', 0)
    
    if not results:
        return jsonify({"error": "Không có dữ liệu"}), 400
    
    excel_file = create_excel(results, ten_sv, mssv, total_nrl)
    return send_file(excel_file, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
