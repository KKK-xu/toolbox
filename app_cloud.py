"""
🛠️ 办公自动化工具箱 v3.0 - 云部署版
适配 Render / Railway 等云平台，文件操作改为用户上传方式

安装依赖：
pip install flask openpyxl pandas PyPDF2 python-dateutil gunicorn

本地运行：
python app_cloud.py

云部署运行：
gunicorn app_cloud:app

环境变量（可选）：
PORT - 端口号（默认 5000）
"""

import os
import re
import shutil
import smtplib
import sqlite3
import tempfile
import uuid
from datetime import datetime
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

from flask import (Flask, render_template_string, request, jsonify,
                   send_file, redirect, url_for, flash)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'toolbox-secret-key-2026')

# ===================== 云端文件管理 =====================
# 用户上传的文件存到临时目录，处理完即删
UPLOAD_ROOT = os.environ.get('UPLOAD_ROOT', os.path.join(tempfile.gettempdir(), 'toolbox_uploads'))
os.makedirs(UPLOAD_ROOT, exist_ok=True)

def get_user_dir():
    """每个会话一个独立目录"""
    session_id = request.cookies.get('session_id', str(uuid.uuid4()))
    user_dir = os.path.join(UPLOAD_ROOT, session_id)
    os.makedirs(user_dir, exist_ok=True)
    return user_dir, session_id

# ===================== 数据库 =====================
DB_PATH = os.environ.get('DB_PATH', os.path.join(UPLOAD_ROOT, 'toolbox.db'))

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        action TEXT NOT NULL, detail TEXT, result TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    c.execute('''CREATE TABLE IF NOT EXISTS email_config (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, smtp_server TEXT, smtp_port INTEGER DEFAULT 465,
        sender TEXT, password TEXT, is_default INTEGER DEFAULT 0,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    c.execute('''CREATE TABLE IF NOT EXISTS invoice_results (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT, invoice_no TEXT, invoice_code TEXT,
        amount TEXT, date TEXT, seller TEXT, buyer TEXT, raw_text TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    conn.commit()
    conn.close()

def log_history(action, detail='', result=''):
    try:
        conn = get_db()
        conn.execute('INSERT INTO history (action, detail, result) VALUES (?, ?, ?)',
                     (action, detail, result))
        conn.commit()
        conn.close()
    except Exception:
        pass


# ===================== HTML 模板（内嵌，无需 templates 目录） =====================

LAYOUT = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>🛠️ 办公自动化工具箱</title>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family: 'Microsoft YaHei', sans-serif; background: #0f172a; color: #e2e8f0; min-height: 100vh; }
.nav { background: #1e293b; padding: 16px 24px; display: flex; align-items: center; gap: 24px; border-bottom: 2px solid #334155; position: sticky; top: 0; z-index: 100; }
.nav .logo { font-size: 20px; font-weight: bold; color: #60a5fa; }
.nav a { color: #94a3b8; text-decoration: none; padding: 6px 14px; border-radius: 6px; transition: all .2s; font-size: 14px; }
.nav a:hover, .nav a.active { background: #334155; color: #f1f5f9; }
.container { max-width: 960px; margin: 32px auto; padding: 0 24px; }
.card { background: #1e293b; border-radius: 12px; padding: 28px; margin-bottom: 24px; border: 1px solid #334155; }
.card h2 { color: #60a5fa; margin-bottom: 20px; font-size: 20px; }
.form-group { margin-bottom: 18px; }
.form-group label { display: block; color: #94a3b8; margin-bottom: 6px; font-size: 14px; }
.form-group input, .form-group select, .form-group textarea {
    width: 100%; padding: 10px 14px; background: #0f172a; border: 1px solid #334155;
    border-radius: 8px; color: #e2e8f0; font-size: 14px; outline: none; transition: border .2s;
}
.form-group input:focus, .form-group select:focus, .form-group textarea:focus { border-color: #60a5fa; }
.form-group textarea { min-height: 80px; resize: vertical; }
.btn { padding: 10px 24px; border: none; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: bold; transition: all .2s; }
.btn-primary { background: #3b82f6; color: white; }
.btn-primary:hover { background: #2563eb; }
.btn-success { background: #10b981; color: white; }
.btn-success:hover { background: #059669; }
.btn-danger { background: #ef4444; color: white; }
.alert { padding: 12px 18px; border-radius: 8px; margin-bottom: 16px; font-size: 14px; }
.alert-success { background: #064e3b; color: #6ee7b7; border: 1px solid #10b981; }
.alert-error { background: #450a0a; color: #fca5a5; border: 1px solid #ef4444; }
.alert-info { background: #1e3a5f; color: #93c5fd; border: 1px solid #3b82f6; }
.stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 16px; margin-bottom: 28px; }
.stat-item { background: #0f172a; border-radius: 10px; padding: 18px; text-align: center; border: 1px solid #334155; }
.stat-item .num { font-size: 28px; font-weight: bold; color: #60a5fa; }
.stat-item .label { color: #64748b; font-size: 13px; margin-top: 4px; }
.file-drop { border: 2px dashed #334155; border-radius: 10px; padding: 30px; text-align: center; color: #64748b; cursor: pointer; transition: all .2s; }
.file-drop:hover, .file-drop.dragover { border-color: #3b82f6; color: #60a5fa; background: rgba(59,130,246,0.05); }
.file-drop input[type="file"] { display: none; }
.result-table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 13px; }
.result-table th { background: #0f172a; color: #94a3b8; padding: 10px 12px; text-align: left; border-bottom: 1px solid #334155; }
.result-table td { padding: 8px 12px; border-bottom: 1px solid #1e293b; }
.result-table tr:hover td { background: #1e293b; }
.download-link { display: inline-block; margin-top: 12px; padding: 8px 20px; background: #10b981; color: white; border-radius: 8px; text-decoration: none; font-size: 14px; }
.download-link:hover { background: #059669; }
.history-item { padding: 10px 0; border-bottom: 1px solid #1e293b; display: flex; justify-content: space-between; }
.history-item .action { color: #60a5fa; font-weight: bold; }
.history-item .time { color: #64748b; font-size: 12px; }
</style>
</head>
<body>
<div class="nav">
    <div class="logo">🛠️ 办公工具箱</div>
    <a href="/" class="{{ 'active' if page=='home' else '' }}">首页</a>
    <a href="/merge-excel" class="{{ 'active' if page=='merge' else '' }}">📊 合并Excel</a>
    <a href="/extract-invoice" class="{{ 'active' if page=='invoice' else '' }}">🧾 提取发票</a>
    <a href="/send-email" class="{{ 'active' if page=='email' else '' }}">📧 发邮件</a>
    <a href="/process-excel" class="{{ 'active' if page=='process' else '' }}">📈 处理Excel</a>
    <a href="/organize-files" class="{{ 'active' if page=='organize' else '' }}">🗂️ 文件整理</a>
    <a href="/history" class="{{ 'active' if page=='history' else '' }}">📋 历史</a>
</div>
<div class="container">
{% block content %}{% endblock %}
</div>
</body>
</html>
"""

HOME_HTML = LAYOUT.replace('{% block content %}{% endblock %}', """
{% if stats %}
<div class="stats">
    <div class="stat-item"><div class="num">{{ stats.total_ops }}</div><div class="label">总操作数</div></div>
    <div class="stat-item"><div class="num">{{ stats.emails_sent }}</div><div class="label">邮件发送</div></div>
    <div class="stat-item"><div class="num">{{ stats.invoices }}</div><div class="label">发票提取</div></div>
    <div class="stat-item"><div class="num">{{ stats.renames }}</div><div class="label">文件处理</div></div>
</div>
{% endif %}
<div class="card">
    <h2>👋 欢迎使用办公自动化工具箱</h2>
    <p style="color:#94a3b8; line-height:1.8;">
        上传文件即可在线处理，无需安装任何软件。<br>
        支持：Excel合并 · PDF发票提取 · 邮件发送 · Excel筛选统计 · 文件分类整理
    </p>
</div>
{% if recent %}
<div class="card">
    <h2>📋 最近操作</h2>
    {% for r in recent %}
    <div class="history-item">
        <span class="action">{{ r.action }}</span>
        <span>{{ r.detail or '' }}</span>
        <span class="time">{{ r.created_at }}</span>
    </div>
    {% endfor %}
</div>
{% endif %}
""")

MERGE_HTML = LAYOUT.replace('{% block content %}{% endblock %}', """
<div class="card">
    <h2>📊 合并多个 Excel 文件</h2>
    <div class="form-group">
        <label>上传多个 Excel 文件</label>
        <div class="file-drop" id="dropZone" onclick="document.getElementById('files').click()">
            点击或拖拽上传 Excel 文件（支持 .xlsx .xls）
            <input type="file" id="files" name="files" multiple accept=".xlsx,.xls">
        </div>
        <div id="fileList" style="margin-top:8px; color:#94a3b8; font-size:13px;"></div>
    </div>
    <div class="form-group">
        <label>合并方式</label>
        <select name="mode">
            <option value="1">纵向合并（所有数据拼到一个 Sheet）</option>
            <option value="2">横向合并（每个文件一个 Sheet）</option>
        </select>
    </div>
    <div class="form-group">
        <label>输出文件名</label>
        <input type="text" name="output_name" value="合并结果.xlsx">
    </div>
    <button class="btn btn-primary" onclick="doMerge()">开始合并</button>
    <div id="result" style="margin-top:16px;"></div>
</div>
<script>
document.getElementById('files').onchange = function() {
    var names = Array.from(this.files).map(f => f.name).join('、');
    document.getElementById('fileList').textContent = '已选: ' + names;
};
function doMerge() {
    var fd = new FormData();
    var files = document.getElementById('files').files;
    if (!files.length) { alert('请先上传文件'); return; }
    for (var f of files) fd.append('files', f);
    fd.append('mode', document.querySelector('[name=mode]').value);
    fd.append('output_name', document.querySelector('[name=output_name]').value);
    fetch('/api/merge-excel', {method:'POST', body:fd}).then(r=>r.json()).then(d => {
        var el = document.getElementById('result');
        if (d.error) { el.innerHTML = '<div class="alert alert-error">❌ '+d.error+'</div>'; return; }
        el.innerHTML = '<div class="alert alert-success">✅ '+d.detail+'</div>' +
            '<a href="/download/'+d.download_id+'" class="download-link">📥 下载合并结果</a>';
    });
}
</script>
""")

INVOICE_HTML = LAYOUT.replace('{% block content %}{% endblock %}', """
<div class="card">
    <h2>🧾 提取 PDF 发票信息</h2>
    <div class="form-group">
        <label>上传 PDF 发票文件</label>
        <div class="file-drop" onclick="document.getElementById('files').click()">
            点击或拖拽上传 PDF 文件（支持多选）
            <input type="file" id="files" name="files" multiple accept=".pdf">
        </div>
    </div>
    <button class="btn btn-primary" onclick="doExtract()">开始提取</button>
    <div id="result" style="margin-top:16px;"></div>
</div>
<script>
function doExtract() {
    var fd = new FormData();
    var files = document.getElementById('files').files;
    if (!files.length) { alert('请先上传文件'); return; }
    for (var f of files) fd.append('files', f);
    fetch('/api/extract-invoice', {method:'POST', body:fd}).then(r=>r.json()).then(d => {
        var el = document.getElementById('result');
        if (d.error) { el.innerHTML = '<div class="alert alert-error">❌ '+d.error+'</div>'; return; }
        var html = '<div class="alert alert-success">✅ 处理 '+d.total+' 个PDF，成功 '+d.success+' 个</div>';
        if (d.results.length) {
            html += '<table class="result-table"><tr><th>文件</th><th>发票号码</th><th>金额</th><th>日期</th><th>开票方</th></tr>';
            d.results.forEach(r => {
                html += '<tr><td>'+r.filename+'</td><td>'+(r.invoice_no||'-')+'</td><td>'+(r.amount||'-')+'</td><td>'+(r.date||'-')+'</td><td>'+(r.seller||'-')+'</td></tr>';
            });
            html += '</table>';
        }
        el.innerHTML = html;
    });
}
</script>
""")

EMAIL_HTML = LAYOUT.replace('{% block content %}{% endblock %}', """
<div class="card">
    <h2>📧 发送邮件</h2>
    <div class="form-group"><label>SMTP 服务器</label><input name="smtp_server" value="smtp.qq.com"></div>
    <div class="form-group"><label>端口</label><input name="smtp_port" value="465" type="number"></div>
    <div class="form-group"><label>发件邮箱</label><input name="sender" placeholder="your@qq.com"></div>
    <div class="form-group"><label>授权码/密码</label><input name="password" type="password" placeholder="不是登录密码，是授权码"></div>
    <div class="form-group"><label>收件人（逗号分隔）</label><input name="receivers" placeholder="a@qq.com, b@163.com"></div>
    <div class="form-group"><label>主题</label><input name="subject"></div>
    <div class="form-group"><label>正文</label><textarea name="body" rows="5"></textarea></div>
    <button class="btn btn-primary" onclick="doSend()">发送邮件</button>
    <div id="result" style="margin-top:16px;"></div>
</div>
<script>
function doSend() {
    var data = {
        smtp_server: document.querySelector('[name=smtp_server]').value,
        smtp_port: document.querySelector('[name=smtp_port]').value,
        sender: document.querySelector('[name=sender]').value,
        password: document.querySelector('[name=password]').value,
        receivers: document.querySelector('[name=receivers]').value,
        subject: document.querySelector('[name=subject]').value,
        body: document.querySelector('[name=body]').value
    };
    fetch('/api/send-email', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(data)})
    .then(r=>r.json()).then(d => {
        var el = document.getElementById('result');
        if (d.error) { el.innerHTML = '<div class="alert alert-error">❌ '+d.error+'</div>'; return; }
        el.innerHTML = '<div class="alert alert-success">✅ '+d.detail+'</div>';
    });
}
</script>
""")

PROCESS_HTML = LAYOUT.replace('{% block content %}{% endblock %}', """
<div class="card">
    <h2>📈 批量处理 Excel</h2>
    <div class="form-group">
        <label>上传 Excel 文件</label>
        <div class="file-drop" onclick="document.getElementById('files').click()">
            点击上传 Excel 文件
            <input type="file" id="files" name="files" multiple accept=".xlsx,.xls">
        </div>
    </div>
    <div class="form-group">
        <label>处理方式</label>
        <select name="mode">
            <option value="3">删除重复行</option>
            <option value="4">清除空行</option>
            <option value="2">分组统计</option>
            <option value="1">条件筛选</option>
        </select>
    </div>
    <div class="form-group"><label>分组/筛选列名（统计和筛选时需要）</label><input name="column" placeholder="如：部门"></div>
    <div class="form-group"><label>统计方式（统计时需要）</label>
        <select name="agg"><option value="sum">求和</option><option value="mean">平均</option><option value="count">计数</option></select>
    </div>
    <button class="btn btn-primary" onclick="doProcess()">开始处理</button>
    <div id="result" style="margin-top:16px;"></div>
</div>
<script>
function doProcess() {
    var fd = new FormData();
    var files = document.getElementById('files').files;
    if (!files.length) { alert('请先上传文件'); return; }
    for (var f of files) fd.append('files', f);
    fd.append('mode', document.querySelector('[name=mode]').value);
    fd.append('column', document.querySelector('[name=column]').value);
    fd.append('agg', document.querySelector('[name=agg]').value);
    fetch('/api/process-excel', {method:'POST', body:fd}).then(r=>r.json()).then(d => {
        var el = document.getElementById('result');
        if (d.error) { el.innerHTML = '<div class="alert alert-error">❌ '+d.error+'</div>'; return; }
        var html = '<div class="alert alert-success">✅ 处理完成，共 '+d.rows+' 行</div>';
        html += '<a href="/download/'+d.download_id+'" class="download-link">📥 下载处理结果</a>';
        if (d.preview && d.preview.length) {
            html += '<table class="result-table"><tr>';
            d.columns.forEach(c => html += '<th>'+c+'</th>');
            html += '</tr>';
            d.preview.slice(0,10).forEach(row => {
                html += '<tr>';
                d.columns.forEach(c => html += '<td>'+(row[c]||'')+'</td>');
                html += '</tr>';
            });
            html += '</table>';
        }
        el.innerHTML = html;
    });
}
</script>
""")

ORGANIZE_HTML = LAYOUT.replace('{% block content %}{% endblock %}', """
<div class="card">
    <h2>🗂️ 文件整理（按类型归档）</h2>
    <div class="form-group">
        <label>上传需要整理的文件</label>
        <div class="file-drop" onclick="document.getElementById('files').click()">
            点击上传文件（支持任意类型，多选）
            <input type="file" id="files" name="files" multiple>
        </div>
    </div>
    <button class="btn btn-primary" onclick="doOrganize()">预览分类</button>
    <button class="btn btn-success" onclick="doOrganize(true)">直接整理并下载</button>
    <div id="result" style="margin-top:16px;"></div>
</div>
<script>
function doOrganize(download) {
    var fd = new FormData();
    var files = document.getElementById('files').files;
    if (!files.length) { alert('请先上传文件'); return; }
    for (var f of files) fd.append('files', f);
    fd.append('download', download ? '1' : '0');
    fetch('/api/organize-files', {method:'POST', body:fd}).then(r => {
        if (download) return r.blob().then(blob => {
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a'); a.href = url; a.download = '整理结果.zip'; a.click();
        });
        return r.json();
    }).then(d => {
        if (!d) return;
        var el = document.getElementById('result');
        if (d.error) { el.innerHTML = '<div class="alert alert-error">❌ '+d.error+'</div>'; return; }
        var html = '<div class="alert alert-info">📂 共 '+d.total+' 个文件，分类如下：</div>';
        for (var cat in d.groups) {
            html += '<div style="margin:8px 0;color:#94a3b8;"><b style="color:#60a5fa;">'+cat+'</b>：'+d.groups[cat].join('、')+'</div>';
        }
        el.innerHTML = html;
    });
}
</script>
""")

HISTORY_HTML = LAYOUT.replace('{% block content %}{% endblock %}', """
<div class="card">
    <h2>📋 操作历史</h2>
    {% if records %}
        {% for r in records %}
        <div class="history-item">
            <span class="action">{{ r.action }}</span>
            <span>{{ r.detail or '' }} {{ r.result or '' }}</span>
            <span class="time">{{ r.created_at }}</span>
        </div>
        {% endfor %}
    {% else %}
        <p style="color:#64748b;">暂无操作记录</p>
    {% endif %}
</div>
""")


# ===================== 文件下载管理 =====================
download_files = {}  # {id: filepath}

# ===================== 路由：首页 =====================
@app.route('/')
def index():
    conn = get_db()
    recent = conn.execute('SELECT * FROM history ORDER BY created_at DESC LIMIT 10').fetchall()
    stats = {
        'total_ops': conn.execute('SELECT COUNT(*) FROM history').fetchone()[0],
        'emails_sent': conn.execute("SELECT COUNT(*) FROM history WHERE action='发送邮件'").fetchone()[0],
        'invoices': conn.execute('SELECT COUNT(*) FROM invoice_results').fetchone()[0],
        'renames': conn.execute('SELECT COUNT(*) FROM history WHERE action IN ('批量重命名','处理Excel','文件整理')").fetchone()[0],
    }
    conn.close()
    return render_template_string(HOME_HTML, page='home', recent=recent, stats=stats)


# ===================== API：合并 Excel =====================
@app.route('/merge-excel')
def merge_excel_page():
    return render_template_string(MERGE_HTML, page='merge')

@app.route('/api/merge-excel', methods=['POST'])
def merge_excel_api():
    import pandas as pd
    files = request.files.getlist('files')
    mode = request.form.get('mode', '1')
    output_name = request.form.get('output_name', '合并结果.xlsx')

    if not files:
        return jsonify({'error': '请上传文件'}), 400

    user_dir, _ = get_user_dir()
    excel_paths = []
    for f in files:
        if f.filename.endswith(('.xlsx', '.xls')):
            path = os.path.join(user_dir, f.filename)
            f.save(path)
            excel_paths.append(path)

    if not excel_paths:
        return jsonify({'error': '没有有效的 Excel 文件'}), 400

    try:
        output_path = os.path.join(user_dir, output_name)
        if mode == "2":
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for p in excel_paths:
                    sheet_name = Path(p).stem[:31]
                    pd.read_excel(p).to_excel(writer, sheet_name=sheet_name, index=False)
            detail = f"横向合并 {len(excel_paths)} 个文件"
        else:
            all_data = []
            for p in excel_paths:
                df = pd.read_excel(p)
                df['来源文件'] = os.path.basename(p)
                all_data.append(df)
            merged = pd.concat(all_data, ignore_index=True)
            merged.to_excel(output_path, index=False)
            detail = f"纵向合并 {len(excel_paths)} 个文件，共 {len(merged)} 行"

        dl_id = str(uuid.uuid4())
        download_files[dl_id] = output_path
        log_history('合并Excel', detail)
        return jsonify({'success': True, 'detail': detail, 'download_id': dl_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ===================== API：提取发票 =====================
@app.route('/extract-invoice')
def extract_invoice_page():
    return render_template_string(INVOICE_HTML, page='invoice')

@app.route('/api/extract-invoice', methods=['POST'])
def extract_invoice_api():
    from PyPDF2 import PdfReader
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': '请上传 PDF 文件'}), 400

    user_dir, _ = get_user_dir()
    results = []
    conn = get_db()

    for f in files:
        if not f.filename.endswith('.pdf'):
            continue
        path = os.path.join(user_dir, f.filename)
        f.save(path)
        try:
            reader = PdfReader(path)
            text = "".join(page.extract_text() or "" for page in reader.pages)
            info = extract_invoice_fields(text)
            info['filename'] = f.filename
            conn.execute('''INSERT INTO invoice_results
                (filename, invoice_no, invoice_code, amount, date, seller, buyer, raw_text)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                (f.filename, info.get('发票号码',''), info.get('发票代码',''),
                 info.get('金额',''), info.get('日期',''),
                 info.get('开票方',''), info.get('购买方',''), text[:500]))
            results.append(info)
        except Exception as e:
            results.append({'filename': f.filename, 'error': str(e)})

    conn.commit(); conn.close()
    success = len([r for r in results if 'error' not in r])
    log_history('提取发票', f'处理 {len(results)} 个PDF', f'成功 {success} 个')
    return jsonify({'success': True, 'results': results, 'total': len(results), 'success': success})

def extract_invoice_fields(text):
    info = {}
    m = re.search(r'发票号码[：:\s]*(\d{8,20})', text)
    if m: info['发票号码'] = m.group(1)
    m = re.search(r'发票代码[：:\s]*(\d{10,12})', text)
    if m: info['发票代码'] = m.group(1)
    m = re.search(r'价税合计[（(]大写[)）].*?[¥￥]\s*([\d,.]+)', text)
    if m: info['金额'] = m.group(1)
    m = re.search(r'开票日期[：:\s]*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', text)
    if m: info['日期'] = m.group(1).replace(' ', '')
    m = re.search(r'销售方.*?名\s*称[：:\s]*([^\n\s]{2,30})', text)
    if m: info['开票方'] = m.group(1).strip()
    m = re.search(r'购买方.*?名\s*称[：:\s]*([^\n\s]{2,30})', text)
    if m: info['购买方'] = m.group(1).strip()
    return info


# ===================== API：发送邮件 =====================
@app.route('/send-email')
def send_email_page():
    return render_template_string(EMAIL_HTML, page='email')

@app.route('/api/send-email', methods=['POST'])
def send_email_api():
    data = request.json or request.form.to_dict()
    smtp_server = data.get('smtp_server', 'smtp.qq.com')
    smtp_port = int(data.get('smtp_port', 465))
    sender = data.get('sender', '')
    password = data.get('password', '')
    receivers = [r.strip() for r in data.get('receivers', '').split(',') if r.strip()]
    subject = data.get('subject', '')
    body = data.get('body', '')

    if not all([sender, password, receivers, subject]):
        return jsonify({'error': '请填写必要信息'}), 400

    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = ', '.join(receivers)
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
        server.login(sender, password)
        server.sendmail(sender, receivers, msg.as_string())
        server.quit()

        log_history('发送邮件', f'发给 {len(receivers)} 人：{subject}', '成功')
        return jsonify({'success': True, 'detail': f'邮件已发送给 {len(receivers)} 人'})
    except smtplib.SMTPAuthenticationError:
        return jsonify({'error': '认证失败，请检查邮箱和授权码'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ===================== API：处理 Excel =====================
@app.route('/process-excel')
def process_excel_page():
    return render_template_string(PROCESS_HTML, page='process')

@app.route('/api/process-excel', methods=['POST'])
def process_excel_api():
    import pandas as pd
    files = request.files.getlist('files')
    mode = request.form.get('mode', '3')
    column = request.form.get('column', '')
    agg = request.form.get('agg', 'sum')

    if not files:
        return jsonify({'error': '请上传文件'}), 400

    user_dir, _ = get_user_dir()
    dfs = []
    for f in files:
        if f.filename.endswith(('.xlsx', '.xls')):
            path = os.path.join(user_dir, f.filename)
            f.save(path)
            dfs.append(pd.read_excel(path))

    if not dfs:
        return jsonify({'error': '没有有效的 Excel 文件'}), 400

    try:
        merged = pd.concat(dfs, ignore_index=True)
        result_data = None

        if mode == "3":  # 去重
            result_data = merged.drop_duplicates()
        elif mode == "4":  # 清空行
            result_data = merged.dropna(how='all')
        elif mode == "2" and column:  # 分组统计
            value_col = request.form.get('value_column', '')
            if value_col:
                result_data = merged.groupby(column)[value_col].agg(agg).reset_index()
            else:
                result_data = merged.groupby(column).agg(agg).reset_index()
        elif mode == "1" and column:  # 筛选
            op = request.form.get('operator', '>')
            val = request.form.get('value', '')
            try: val = float(val)
            except: pass
            ops = {'>': lambda d: d[d[column] > val], '<': lambda d: d[d[column] < val],
                   '>=': lambda d: d[d[column] >= val], '<=': lambda d: d[d[column] <= val],
                   '==': lambda d: d[d[column] == val]}
            result_data = ops.get(op, lambda d: d)(merged)

        if result_data is not None:
            output_path = os.path.join(user_dir, '处理结果.xlsx')
            result_data.to_excel(output_path, index=False)
            preview = result_data.head(50).to_dict('records')
            columns = list(result_data.columns)
            dl_id = str(uuid.uuid4())
            download_files[dl_id] = output_path
            log_history('处理Excel', f'模式{mode}', f'输出 {len(result_data)} 行')
            return jsonify({'success': True, 'preview': preview, 'columns': columns,
                          'rows': len(result_data), 'download_id': dl_id})
        return jsonify({'error': '请设置筛选/统计参数'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ===================== API：文件整理 =====================
@app.route('/organize-files')
def organize_files_page():
    return render_template_string(ORGANIZE_HTML, page='organize')

@app.route('/api/organize-files', methods=['POST'])
def organize_files_api():
    import zipfile
    files = request.files.getlist('files')
    download = request.form.get('download', '0')

    if not files:
        return jsonify({'error': '请上传文件'}), 400

    user_dir, _ = get_user_dir()
    type_map = {
        '图片': ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp'],
        '文档': ['.doc', '.docx', '.pdf', '.txt', '.md'],
        '表格': ['.xls', '.xlsx', '.csv'],
        '演示': ['.ppt', '.pptx'],
        '视频': ['.mp4', '.avi', '.mkv', '.mov'],
        '音频': ['.mp3', '.wav', '.flac', '.aac'],
        '压缩包': ['.zip', '.rar', '.7z'],
        '代码': ['.py', '.js', '.html', '.css', '.java'],
    }

    groups = {}
    for f in files:
        ext = os.path.splitext(f.filename)[1].lower()
        cat = '其他'
        for c, exts in type_map.items():
            if ext in exts: cat = c; break
        groups.setdefault(cat, []).append(f.filename)
        # 保存到分类目录
        cat_dir = os.path.join(user_dir, cat)
        os.makedirs(cat_dir, exist_ok=True)
        f.save(os.path.join(cat_dir, f.filename))

    log_history('文件整理', f'整理 {len(files)} 个文件')

    if download == '1':
        # 打包成 ZIP 下载
        zip_path = os.path.join(user_dir, '整理结果.zip')
        with zipfile.ZipFile(zip_path, 'w') as zf:
            for cat in groups:
                cat_dir = os.path.join(user_dir, cat)
                for fname in os.listdir(cat_dir):
                    zf.write(os.path.join(cat_dir, fname), os.path.join(cat, fname))
        return send_file(zip_path, as_attachment=True, download_name='整理结果.zip')

    return jsonify({'success': True, 'groups': groups, 'total': len(files)})


# ===================== 文件下载 =====================
@app.route('/download/<dl_id>')
def download_file(dl_id):
    path = download_files.get(dl_id)
    if path and os.path.exists(path):
        return send_file(path, as_attachment=True, download_name=os.path.basename(path))
    return "文件不存在或已过期", 404


# ===================== 历史记录 =====================
@app.route('/history')
def history_page():
    conn = get_db()
    records = conn.execute('SELECT * FROM history ORDER BY created_at DESC LIMIT 100').fetchall()
    conn.close()
    return render_template_string(HISTORY_HTML, page='history', records=records)


# ===================== 启动 =====================
if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5000))
    print(f"\n🛠️  办公自动化工具箱 v3.0（云部署版）已启动！")
    print(f"📱 打开浏览器访问：http://127.0.0.1:{port}")
    print(f"☁️  云部署命令：gunicorn app_cloud:app\n")
    app.run(host='0.0.0.0', port=port, debug=True)
