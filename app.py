from flask import Flask, request, send_file
from flask_cors import CORS
import openpyxl
import shutil, io, os, tempfile
from datetime import datetime

app = Flask(__name__)
CORS(app)

BASE     = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE, 'template.xlsx')


def make_excel(data):
    tax_mode = data.get('tax_mode', 'with')

    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.close()
    shutil.copy2(TEMPLATE, tmp.name)

    wb = openpyxl.load_workbook(tmp.name)
    ws = wb['報價單']

    ws['C8'].value  = data.get('date', '')
    ws['C9'].value  = data.get('client', '')
    ws['C10'].value = data.get('contact', '')

    for r in range(16, 36):
        ws[f'B{r}'].value = None
        ws[f'D{r}'].value = None
        ws[f'E{r}'].value = None
        ws[f'F{r}'].value = f'=SUM(D{r}*E{r})'

    items = data.get('items', [])
    for i, item in enumerate(items[:19]):
        r = 16 + i
        ws[f'B{r}'].value = item.get('desc', '')
        ws[f'D{r}'].value = float(item.get('price') or 0)
        ws[f'E{r}'].value = float(item.get('qty') or 0)

    blank_row = 16 + len(items)
    if blank_row <= 35:
        ws[f'B{blank_row}'].value = '以下空白'

    note = data.get('note', '')
    ws['B37'].value = f'備註:{note}' if note else '備註:'

    if tax_mode == 'without':
        ws['E37'].value = '小計'
        ws['F37'].value = '=SUM(F16:F35)'
        ws['E38'].value = None
        ws['F38'].value = None
        ws['E39'].value = '總計金額'
        ws['F39'].value = '=SUM(F37)'
    else:
        ws['E37'].value = '小計'
        ws['F37'].value = '=SUM(F16:F35)'
        ws['E38'].value = '稅金'
        ws['F38'].value = '=SUM(F37*0.05)'
        ws['E39'].value = '總計金額'
        ws['F39'].value = '=SUM(F37:F38)'

    wb.save(tmp.name)
    return tmp.name


@app.route('/')
def index():
    return send_file(os.path.join(BASE, 'index.html'))

@app.route('/ping')
def ping():
    return 'ok'


@app.route('/generate', methods=['POST'])
def generate():
    data      = request.json
    xlsx_path = make_excel(data)

    with open(xlsx_path, 'rb') as f:
        content = f.read()
    os.unlink(xlsx_path)

    client    = data.get('client', '未命名')
    today     = datetime.now().strftime('%Y%m%d')
    tax_label = '含稅' if data.get('tax_mode','with') == 'with' else '未稅'
    filename  = f'厚德報價單_{client}_{tax_label}_{today}.xlsx'

    return send_file(
        io.BytesIO(content),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/generate_pdf', methods=['POST'])
def generate_pdf():
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return ('尚未安裝 pywin32，請執行：pip install pywin32', 503)

    data      = request.json
    xlsx_path = make_excel(data)
    xlsx_path = os.path.abspath(xlsx_path)
    pdf_path  = xlsx_path.replace('.xlsx', '.pdf')

    excel = None
    try:
        # ✅ 關鍵修正：Flask 執行緒需要手動初始化 COM
        pythoncom.CoInitialize()

        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible       = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(xlsx_path)
        ws = wb.Worksheets('報價單')
        ws.ExportAsFixedFormat(
            Type                 = 0,
            Filename             = pdf_path,
            Quality              = 0,
            IncludeDocProperties = True,
            IgnorePrintAreas     = False,
        )
        wb.Close(False)

        with open(pdf_path, 'rb') as f:
            pdf_content = f.read()

        client    = data.get('client', '未命名')
        today     = datetime.now().strftime('%Y%m%d')
        tax_label = '含稅' if data.get('tax_mode','with') == 'with' else '未稅'
        filename  = f'厚德報價單_{client}_{tax_label}_{today}.pdf'

        return send_file(
            io.BytesIO(pdf_content),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return f'PDF 產生失敗：{str(e)}', 500

    finally:
        try:
            if excel:
                excel.Quit()
        except:
            pass
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        try:
            os.unlink(xlsx_path)
        except:
            pass
        try:
            os.unlink(pdf_path)
        except:
            pass


if __name__ == '__main__':
    import socket

    try:
        import win32com.client
        pdf_status = '✅ pywin32 已安裝，PDF 功能正常'
    except ImportError:
        pdf_status = '⚠️  請執行：pip install pywin32'

    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        local_ip = s.getsockname()[0]
        s.close()
    except:
        local_ip = '127.0.0.1'

    print(f"\n✅ 伺服器啟動成功！")
    print(f"\n  電腦瀏覽器: http://localhost:5000")
    print(f"  手機瀏覽器: http://{local_ip}:5000")
    print(f"\n  {pdf_status}")
    print(f"\n  手機和電腦必須連同一個 WiFi")
    print(f"  按 Ctrl+C 停止\n")
    app.run(host='0.0.0.0', port=5000, debug=False)
