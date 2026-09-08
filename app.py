# -*- coding: utf-8 -*-
import os, sys, re, json, sqlite3, smtplib, threading, time
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from io import BytesIO
from flask import Flask, render_template_string, request, jsonify, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "bipagem_nfe.db")
CONFIG_FILE = os.path.join(BASE_DIR, "config_bipagem.json")

def obter_conexao():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def inicializar_banco():
    conn = obter_conexao()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS notas (
            chave TEXT PRIMARY KEY, numero_nf TEXT, serie TEXT,
            data_bip1 TEXT, hora_bip1 TEXT, dt_completa_bip1 TEXT,
            data_bip2 TEXT, hora_bip2 TEXT, dt_completa_bip2 TEXT,
            tempo_decorrido TEXT, minutos_decorridos REAL,
            status TEXT, justificativa TEXT, qtd_bipagens INTEGER DEFAULT 1
        )
    """)
    conn.commit()
    conn.close()

def sanitizar_e_extrair_chave(texto):
    nums = re.sub(r"\D", "", texto or "")
    if len(nums) >= 44:
        chave = nums[:44]
        try:
            serie = str(int(chave[22:25]))
            numero_nf = str(int(chave[25:34]))
        except Exception:
            serie = chave[22:25]
            numero_nf = chave[25:34]
        return {"valida": True, "chave": chave, "numero_nf": numero_nf, "serie": serie}
    return {"valida": False, "chave": nums, "numero_nf": "N/D", "serie": "N/D"}

def calcular_diferenca(dt1, dt2):
    delta = dt2 - dt1
    s = max(0, int(delta.total_seconds()))
    d, r = divmod(s, 86400)
    h, r = divmod(r, 3600)
    m, seg = divmod(r, 60)
    partes = []
    if d > 0: partes.append(f"{d}d")
    if h > 0 or d > 0: partes.append(f"{h}h")
    partes.append(f"{m}m {seg}s")
    return " ".join(partes), round(s / 60.0, 2)

HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Portaria Web • Controle de NF-e</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>body{background:#f3f4f6;font-family:'Segoe UI',sans-serif;}.barra-input{font-size:1.4rem;font-family:monospace;font-weight:bold;}</style>
</head>
<body>
    <nav class="navbar navbar-dark bg-primary mb-4 shadow-sm">
        <div class="container-fluid px-4">
            <span class="navbar-brand fw-bold">🚪 PORTARIA WEB • CONTROLE DE SAÍDAS DE NF-e</span>
            <span class="text-white-50" id="clock"></span>
        </div>
    </nav>
    <div class="container-fluid px-4">
        <div class="card p-4 mb-4 border-0 shadow-sm">
            <h5 class="text-secondary fw-bold mb-3">⚡ Leitura de Código de Barras (Chave de 44 dígitos)</h5>
            <div class="input-group mb-3">
                <input type="text" id="chaveInput" class="form-control barra-input text-center" placeholder="Aponte o leitor de código de barras aqui..." autofocus autocomplete="off">
                <button class="btn btn-outline-secondary px-4 fw-bold" onclick="limpar()">Limpar</button>
            </div>
            <div id="painel" class="alert alert-light border text-center my-0 py-3 fw-bold text-success">🟢 Leitor Pronto para Bipagem</div>
        </div>
        <div class="card p-4 border-0 shadow-sm">
            <div class="d-flex justify-content-between align-items-center mb-3">
                <h5 class="text-secondary fw-bold mb-0">🔎 Histórico de Registros e Desvios</h5>
                <a href="/api/exportar" class="btn btn-success fw-bold">📊 Baixar Excel (.xlsx)</a>
            </div>
            <div class="table-responsive">
                <table class="table table-hover align-middle"><thead class="table-light"><tr class="text-center text-secondary small">
                    <th>NF</th><th>Chave de Acesso</th><th>1ª Bipagem</th><th>2ª Bipagem</th><th>Diferença</th><th>Status</th><th>Justificativa</th>
                </tr></thead><tbody id="corpo"></tbody></table>
            </div>
        </div>
    </div>
    <div class="modal fade" id="modalJust" tabindex="-1" data-bs-backdrop="static"><div class="modal-dialog modal-dialog-centered"><div class="modal-content">
        <div class="modal-header bg-danger text-white"><h5 class="modal-title fw-bold">⚠️ DUPLO REGISTRO IDENTIFICADO</h5></div>
        <div class="modal-body">
            <p id="descModal" class="mb-3"></p>
            <label class="form-label fw-bold">Informe o motivo da 2ª bipagem:</label>
            <textarea id="justInput" class="form-control" rows="3"></textarea>
        </div>
        <div class="modal-footer"><button type="button" class="btn btn-danger fw-bold w-100" onclick="salvarJust()">Gravar Justificativa</button></div>
    </div></div></div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        let modal = new bootstrap.Modal(document.getElementById('modalJust'));
        let chaveDesvio = "", ultChave = "", ultTime = 0;
        setInterval(() => document.getElementById('clock').innerText = new Date().toLocaleDateString() + ' ' + new Date().toLocaleTimeString(), 1000);
        function focar(){ const i = document.getElementById('chaveInput'); i.focus(); i.select(); }
        function limpar(){ document.getElementById('chaveInput').value = ""; focar(); }
        document.getElementById('chaveInput').addEventListener('keypress', e => { if (e.key === 'Enter') { e.preventDefault(); bipar(); } });
        async function bipar(){
            const chave = document.getElementById('chaveInput').value.trim();
            if (!chave) return;
            const now = Date.now();
            if (chave === ultChave && (now - ultTime < 2500)) return;
            ultChave = chave; ultTime = now;
            const res = await fetch('/api/bipar', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({chave}) });
            const data = await res.json();
            const p = document.getElementById('painel');
            if (!data.sucesso) { p.className = "alert alert-danger border text-center my-0 py-3 fw-bold"; p.innerText = "❌ " + data.mensagem; focar(); return; }
            if (data.tipo === "1ª Bipagem") {
                p.className = "alert alert-primary border text-center my-0 py-3 fw-bold";
                p.innerText = "✅ 1ª Bipagem Gravada: NF " + data.dados.numero_nf + " às " + data.dados.hora_bip1;
            } else {
                p.className = "alert alert-danger border text-center my-0 py-3 fw-bold";
                p.innerText = "🚨 ALERTA DE DESVIO: NF " + data.dados.numero_nf + " | Tempo: " + data.dados.tempo_decorrido;
                chaveDesvio = data.dados.chave;
                document.getElementById('descModal').innerText = "A NF " + data.dados.numero_nf + " já registrou saída. Tempo decorrido: " + data.dados.tempo_decorrido;
                document.getElementById('justInput').value = "";
                modal.show();
            }
            carregar(); focar();
        }
        async function salvarJust(){
            const just = document.getElementById('justInput').value.trim();
            if (!just) return alert("Digite o motivo.");
            await fetch('/api/justificar', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({chave: chaveDesvio, justificativa: just}) });
            modal.hide(); carregar(); focar();
        }
        async function carregar(){
            const res = await fetch('/api/historico');
            const lista = await res.json();
            const tbody = document.getElementById('corpo');
            tbody.innerHTML = "";
            lista.forEach(n => {
                const st = n.status.includes("DESVIO") ? '<span class="badge bg-danger">🚨 DESVIO</span>' : '<span class="badge bg-success">REGULAR</span>';
                tbody.innerHTML += `<tr class="text-center"><td class="fw-bold">${n.numero_nf}</td><td class="font-monospace small text-start">${n.chave}</td><td>${n.data_bip1} ${n.hora_bip1}</td><td>${n.data_bip2 ? n.data_bip2 + ' ' + n.hora_bip2 : '-'}</td><td class="fw-bold">${n.tempo_decorrido || '-'}</td><td>${st}</td><td class="text-start small">${n.justificativa || '-'}</td></tr>`;
            });
        }
        carregar(); focar();
    </script>
</body>
</html>"""

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/api/bipar", methods=["POST"])
def api_bipar():
    chave_raw = (request.get_json() or {}).get("chave", "").strip()
    d = sanitizar_e_extrair_chave(chave_raw)
    if not d["valida"]:
        return jsonify({"sucesso": False, "mensagem": f"Chave inválida ({len(d['chave'])} dígitos). Precisa ter 44."})
    
    agora = datetime.now()
    dt_str, hr_str, iso_str = agora.strftime("%d/%m/%Y"), agora.strftime("%H:%M:%S"), agora.isoformat()
    conn = obter_conexao()
    c = conn.cursor()
    c.execute("SELECT * FROM notas WHERE chave = ?", (d["chave"],))
    nota = c.fetchone()

    if not nota:
        c.execute("INSERT INTO notas (chave, numero_nf, serie, data_bip1, hora_bip1, dt_completa_bip1, status, justificativa, qtd_bipagens) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)",
                  (d["chave"], d["numero_nf"], d["serie"], dt_str, hr_str, iso_str, "REGULAR", ""))
        conn.commit()
        conn.close()
        return jsonify({"sucesso": True, "tipo": "1ª Bipagem", "dados": {"chave": d["chave"], "numero_nf": d["numero_nf"], "serie": d["serie"], "data_bip1": dt_str, "hora_bip1": hr_str}})
    else:
        try: dt1 = datetime.fromisoformat(nota["dt_completa_bip1"] or iso_str)
        except: dt1 = agora
        tempo_txt, mins = calcular_diferenca(dt1, agora)
        qtd = (nota["qtd_bipagens"] or 1) + 1
        c.execute("UPDATE notas SET data_bip2=?, hora_bip2=?, dt_completa_bip2=?, tempo_decorrido=?, minutos_decorridos=?, status=?, qtd_bipagens=? WHERE chave=?",
                  (dt_str, hr_str, iso_str, tempo_txt, mins, "DESVIO REGISTRADO", qtd, d["chave"]))
        conn.commit()
        conn.close()
        return jsonify({"sucesso": True, "tipo": "2ª Bipagem", "dados": {"chave": d["chave"], "numero_nf": d["numero_nf"], "serie": d["serie"], "data_bip1": nota["data_bip1"], "hora_bip1": nota["hora_bip1"], "data_bip2": dt_str, "hora_bip2": hr_str, "tempo_decorrido": tempo_txt}})

@app.route("/api/justificar", methods=["POST"])
def api_just():
    req = request.get_json() or {}
    conn = obter_conexao()
    conn.execute("UPDATE notas SET justificativa = ? WHERE chave = ?", (req.get("justificativa", "").strip(), req.get("chave")))
    conn.commit()
    conn.close()
    return jsonify({"sucesso": True})

@app.route("/api/historico")
def api_hist():
    conn = obter_conexao()
    rows = conn.execute("SELECT * FROM notas ORDER BY dt_completa_bip1 DESC").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route("/api/exportar")
def api_exp():
    conn = obter_conexao()
    registros = conn.execute("SELECT * FROM notas ORDER BY dt_completa_bip1 DESC").fetchall()
    conn.close()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Portaria"
    headers = ["Chave de Acesso", "Número NF", "Série", "Data 1ª Bip", "Hora 1ª Bip", "Data 2ª Bip", "Hora 2ª Bip", "Diferença", "Status", "Justificativa"]
    ws.append(headers)
    for r in registros:
        ws.append([r["chave"], r["numero_nf"], r["serie"], r["data_bip1"], r["hora_bip1"], r["data_bip2"] or "-", r["hora_bip2"] or "-", r["tempo_decorrido"] or "-", r["status"], r["justificativa"] or "-"])
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name="Relatorio_Portaria.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    inicializar_banco()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
