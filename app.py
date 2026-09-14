# -*- coding: utf-8 -*-
import os
import re
import json
import sqlite3
import smtplib
import threading
from datetime import datetime, timezone, timedelta
from io import BytesIO

from flask import Flask, render_template_string, request, jsonify, send_file
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl

# Fuso horário local (corrige diferença de hora no Render)
try:
    from zoneinfo import ZoneInfo
    LOCAL_TZ = ZoneInfo(os.getenv("APP_TZ", "America/Sao_Paulo"))
except Exception:
    LOCAL_TZ = timezone(timedelta(hours=-3))

app = Flask(__name__)
application = app  # compatível com gunicorn app:application

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "bipagem_nfe.db")
CONFIG_FILE = os.path.join(BASE_DIR, "config_bipagem.json")


# =========================
# Helpers de data/hora
# =========================
def agora_local() -> datetime:
    return datetime.now(LOCAL_TZ)


def parse_dt_banco(valor: str):
    """
    Converte ISO do banco para datetime com fuso local.
    Aceita registros antigos sem tzinfo.
    """
    if not valor:
        return None
    try:
        dt = datetime.fromisoformat(valor)
    except Exception:
        return None

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=LOCAL_TZ)
    else:
        dt = dt.astimezone(LOCAL_TZ)
    return dt


# =========================
# Banco
# =========================
def obter_conexao():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_banco():
    conn = obter_conexao()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS notas (
            chave TEXT PRIMARY KEY,
            numero_nf TEXT,
            serie TEXT,
            data_bip1 TEXT,
            hora_bip1 TEXT,
            dt_completa_bip1 TEXT,
            data_bip2 TEXT,
            hora_bip2 TEXT,
            dt_completa_bip2 TEXT,
            tempo_decorrido TEXT,
            minutos_decorridos REAL,
            status TEXT,
            justificativa TEXT,
            qtd_bipagens INTEGER DEFAULT 1,
            observacao TEXT,
            email_desvio_enviado INTEGER DEFAULT 0
        )
    """)
    conn.commit()

    # Migração para banco antigo
    c.execute("PRAGMA table_info(notas)")
    cols = [x["name"] for x in c.fetchall()]
    needed = {
        "data_bip2": "TEXT",
        "hora_bip2": "TEXT",
        "dt_completa_bip2": "TEXT",
        "tempo_decorrido": "TEXT",
        "minutos_decorridos": "REAL",
        "status": "TEXT",
        "justificativa": "TEXT",
        "qtd_bipagens": "INTEGER DEFAULT 1",
        "observacao": "TEXT",
        "email_desvio_enviado": "INTEGER DEFAULT 0",
    }
    for col, typ in needed.items():
        if col not in cols:
            try:
                c.execute(f"ALTER TABLE notas ADD COLUMN {col} {typ}")
            except Exception:
                pass

    conn.commit()
    conn.close()


# =========================
# Regras
# =========================
def sanitizar_e_extrair_chave(texto):
    nums = re.sub(r"\D", "", texto or "")
    if len(nums) >= 44:
        chave = nums[:44]
        try:
            serie = str(int(chave[22:25]))
        except Exception:
            serie = chave[22:25]
        try:
            numero_nf = str(int(chave[25:34]))
        except Exception:
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
    if d > 0:
        partes.append(f"{d}d")
    if h > 0 or d > 0:
        partes.append(f"{h}h")
    partes.append(f"{m}m {seg}s")
    return " ".join(partes), round(s / 60.0, 2)


def carregar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "destinatarios": [],
        "smtp_server": "smtp.office365.com",
        "smtp_port": 587,
        "email_remetente": "",
        "senha_app": "",
    }


def enviar_email_smtp(assunto, corpo_html, destinatarios):
    cfg = carregar_config()
    remetente = cfg.get("email_remetente", "").strip()
    senha = cfg.get("senha_app", "").strip()
    srv = cfg.get("smtp_server", "smtp.office365.com").strip()
    porta = int(cfg.get("smtp_port", 587))

    destinatarios = [d.strip() for d in (destinatarios or []) if d and d.strip()]
    if not remetente or not senha or not destinatarios:
        return False

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = assunto
        msg["From"] = remetente
        msg["To"] = ", ".join(destinatarios)
        msg.attach(MIMEText(corpo_html, "html", "utf-8"))

        with smtplib.SMTP(srv, porta, timeout=20) as server:
            server.starttls()
            server.login(remetente, senha)
            server.sendmail(remetente, destinatarios, msg.as_string())
        return True
    except Exception:
        return False


def template_email_desvio(dados):
    return f"""
    <div style="font-family:Arial,sans-serif;padding:20px;color:#333">
      <div style="background:#DC2626;color:#fff;padding:15px;border-radius:6px;text-align:center">
        <h2 style="margin:0">⚠️ ALERTA DE DESVIO: DUPLA BIPAGEM NA PORTARIA</h2>
        <p style="margin:5px 0 0 0">Nota Fiscal: {dados.get('numero_nf')} | Série: {dados.get('serie')}</p>
      </div>
      <div style="padding:20px;border:1px solid #e5e7eb;border-radius:6px;margin-top:15px">
        <p><strong>Chave:</strong> <code style="font-size:11px">{dados.get('chave')}</code></p>
        <p><strong>1ª Bipagem:</strong> {dados.get('data_bip1')} às {dados.get('hora_bip1')}</p>
        <p><strong>2ª Bipagem:</strong> <span style="color:#DC2626;font-weight:bold">{dados.get('data_bip2')} às {dados.get('hora_bip2')}</span></p>
        <p><strong>Diferença:</strong> ⏱️ <b>{dados.get('tempo_decorrido')}</b></p>
        <p><strong>Justificativa:</strong> {dados.get('justificativa', 'Pendente')}</p>
      </div>
    </div>
    """


# =========================
# HTML
# =========================
HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Portaria Web • Controle de NF-e</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body{background:#f3f4f6;font-family:'Segoe UI',sans-serif;}
    .barra-input{font-size:1.4rem;font-family:monospace;font-weight:bold;}
  </style>
</head>
<body>
  <nav class="navbar navbar-dark bg-primary mb-4 shadow-sm">
    <div class="container-fluid px-4">
      <span class="navbar-brand fw-bold">🚪 PORTARIA WEB • CONTROLE DE SAÍDAS DE NF-e</span>
      <span class="text-white-50 fw-bold" id="clock"></span>
    </div>
  </nav>

  <div class="container-fluid px-4">
    <div class="card p-4 mb-4 border-0 shadow-sm">
      <h5 class="text-secondary fw-bold mb-3">⚡ Leitura de Código de Barras (Chave de 44 dígitos)</h5>
      <div class="input-group mb-3">
        <input type="text" id="chaveInput" class="form-control barra-input text-center"
               placeholder="Aponte o leitor de código de barras aqui..." autofocus autocomplete="off">
        <button id="btnRegistrar" type="button" class="btn btn-primary px-4 fw-bold">Registrar</button>
        <button id="btnLimpar" type="button" class="btn btn-outline-secondary px-4 fw-bold">Limpar</button>
      </div>
      <div id="painel" class="alert alert-light border text-center my-0 py-3 fw-bold text-success">
        🟢 Leitor Pronto para Bipagem
      </div>
    </div>

    <div class="card p-4 border-0 shadow-sm">
      <div class="d-flex justify-content-between align-items-center mb-3">
        <h5 class="text-secondary fw-bold mb-0">🔎 Histórico de Registros e Desvios</h5>
        <a href="/api/exportar" class="btn btn-success fw-bold">📊 Baixar Excel (.xlsx)</a>
      </div>
      <div class="table-responsive">
        <table class="table table-hover align-middle">
          <thead class="table-light">
            <tr class="text-center text-secondary small">
              <th>NF</th><th>Chave de Acesso</th><th>1ª Bipagem</th><th>2ª Bipagem</th><th>Diferença</th><th>Status</th><th>Justificativa</th>
            </tr>
          </thead>
          <tbody id="corpo"></tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="modal fade" id="modalJust" tabindex="-1" data-bs-backdrop="static">
    <div class="modal-dialog modal-dialog-centered"><div class="modal-content">
      <div class="modal-header bg-danger text-white">
        <h5 class="modal-title fw-bold">⚠️ DUPLO REGISTRO IDENTIFICADO</h5>
      </div>
      <div class="modal-body">
        <p id="descModal" class="mb-3"></p>
        <label class="form-label fw-bold">Informe o motivo da 2ª bipagem (Obrigatório):</label>
        <textarea id="justInput" class="form-control" rows="3" placeholder="Digite aqui..."></textarea>
      </div>
      <div class="modal-footer">
        <button id="btnSalvarJust" type="button" class="btn btn-danger fw-bold w-100">
          Gravar Justificativa e Enviar E-mail
        </button>
      </div>
    </div></div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
  <script>
    document.addEventListener('DOMContentLoaded', function () {
      const inputElem = document.getElementById('chaveInput');
      const btnRegistrar = document.getElementById('btnRegistrar');
      const btnLimpar = document.getElementById('btnLimpar');
      const btnSalvarJust = document.getElementById('btnSalvarJust');
      const painel = document.getElementById('painel');
      const corpo = document.getElementById('corpo');
      const descModal = document.getElementById('descModal');
      const justInput = document.getElementById('justInput');

      const modalEl = document.getElementById('modalJust');
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl);

      let chaveDesvio = "";
      let ultChave = "";
      let ultTime = 0;
      let processando = false;

      // Relógio apenas visual (navegador)
      setInterval(() => {
        const d = new Date();
        document.getElementById('clock').innerText =
          d.toLocaleDateString('pt-BR') + ' ' + d.toLocaleTimeString('pt-BR');
      }, 1000);

      function focar() {
        inputElem.focus();
        inputElem.select();
      }

      function limpar() {
        inputElem.value = "";
        ultChave = "";
        focar();
      }

      async function executarRequisicao(chave) {
        processando = true;
        painel.className = "alert alert-warning border text-center my-0 py-3 fw-bold";
        painel.innerText = "⏳ Processando...";

        try {
          const res = await fetch('/api/bipar', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({chave})
          });

          if (!res.ok) throw new Error("Erro no servidor (" + res.status + ")");
          const data = await res.json();

          if (!data.sucesso) {
            painel.className = "alert alert-danger border text-center my-0 py-3 fw-bold";
            painel.innerText = "❌ " + data.mensagem;
          } else if (data.tipo === "1ª Bipagem") {
            painel.className = "alert alert-primary border text-center my-0 py-3 fw-bold";
            painel.innerText = "✅ 1ª Bipagem Gravada: NF " + data.dados.numero_nf + " às " + data.dados.hora_bip1;
            inputElem.value = "";
          } else {
            painel.className = "alert alert-danger border text-center my-0 py-3 fw-bold";
            painel.innerText = "🚨 ALERTA DE DESVIO: NF " + data.dados.numero_nf + " | Tempo: " + data.dados.tempo_decorrido;
            chaveDesvio = data.dados.chave;
            descModal.innerText = "A NF " + data.dados.numero_nf + " já registrou saída anterior.\\nTempo decorrido: " + data.dados.tempo_decorrido;
            justInput.value = "";
            btnSalvarJust.innerText = "Gravar Justificativa e Enviar E-mail";
            btnSalvarJust.disabled = false;
            modal.show();
          }
        } catch (e) {
          console.error(e);
          painel.className = "alert alert-danger border text-center my-0 py-3 fw-bold";
          painel.innerText = "❌ Falha de comunicação. Verifique sua internet.";
          ultChave = "";
        } finally {
          processando = false;
          await carregar();
          focar();
        }
      }

      function biparManual() {
        if (processando) return;
        const chave = inputElem.value.trim();
        if (!chave) {
          alert("Por favor, bip a nota fiscal ou cole a chave primeiro!");
          focar();
          return;
        }
        ultChave = "";
        executarRequisicao(chave);
      }

      function biparAutomatico() {
        if (processando) return;
        const chave = inputElem.value.trim();
        if (!chave) return;

        const now = Date.now();
        if (chave === ultChave && (now - ultTime < 2500)) return; // anti-rebote 2.5s

        ultChave = chave;
        ultTime = now;
        executarRequisicao(chave);
      }

      async function salvarJust() {
        const just = justInput.value.trim();
        if (!just) return alert("Digite o motivo obrigatório.");

        btnSalvarJust.disabled = true;
        btnSalvarJust.innerText = "Salvando e Enviando E-mail...";

        try {
          await fetch('/api/justificar', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({chave: chaveDesvio, justificativa: just})
          });
          inputElem.value = "";
        } catch (e) {
          console.log(e);
        }

        modal.hide();
        await carregar();
        focar();
      }

      async function carregar() {
        try {
          const res = await fetch('/api/historico');
          const lista = await res.json();
          corpo.innerHTML = "";

          lista.forEach(n => {
            let st = "";
            const status = n.status || "";
            if (status.includes("DESVIO")) st = '<span class="badge bg-danger">🚨 DESVIO</span>';
            else if (status.includes("PENDENTE")) st = '<span class="badge bg-warning text-dark">⚠️ PENDENTE</span>';
            else st = '<span class="badge bg-success">REGULAR</span>';

            corpo.innerHTML += `<tr class="text-center">
              <td class="fw-bold">${n.numero_nf || '-'}</td>
              <td class="font-monospace small text-start">${n.chave || '-'}</td>
              <td>${n.data_bip1 || '-'} ${n.hora_bip1 || ''}</td>
              <td>${n.data_bip2 ? n.data_bip2 + ' ' + (n.hora_bip2 || '') : '-'}</td>
              <td class="fw-bold">${n.tempo_decorrido || '-'}</td>
              <td>${st}</td>
              <td class="text-start small">${n.justificativa || '-'}</td>
            </tr>`;
          });
        } catch (e) {}
      }

      btnRegistrar.addEventListener('click', biparManual);
      btnLimpar.addEventListener('click', limpar);
      btnSalvarJust.addEventListener('click', salvarJust);

      inputElem.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          biparAutomatico();
        }
      });

      inputElem.addEventListener('input', (e) => {
        const v = e.target.value.replace(/\\D/g, '');
        if (v.length === 44 && v !== ultChave) {
          setTimeout(biparAutomatico, 200);
        }
      });

      carregar();
      focar();
    });
  </script>
</body>
</html>
"""


# =========================
# Rotas
# =========================
@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/api/bipar", methods=["POST"])
def api_bipar():
    chave_raw = (request.get_json(silent=True) or {}).get("chave", "").strip()
    d = sanitizar_e_extrair_chave(chave_raw)
    if not d["valida"]:
        return jsonify({"sucesso": False, "mensagem": f"Chave inválida ({len(d['chave'])} dígitos). Precisa ter 44."})

    agora = agora_local()
    dt_str = agora.strftime("%d/%m/%Y")
    hr_str = agora.strftime("%H:%M:%S")
    iso_str = agora.isoformat()

    conn = obter_conexao()
    c = conn.cursor()
    c.execute("SELECT * FROM notas WHERE chave = ?", (d["chave"],))
    nota = c.fetchone()

    if not nota:
        c.execute("""
            INSERT INTO notas
            (chave, numero_nf, serie, data_bip1, hora_bip1, dt_completa_bip1, status, justificativa, qtd_bipagens)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
        """, (d["chave"], d["numero_nf"], d["serie"], dt_str, hr_str, iso_str, "REGULAR", ""))
        conn.commit()
        conn.close()

        return jsonify({
            "sucesso": True,
            "tipo": "1ª Bipagem",
            "dados": {
                "chave": d["chave"],
                "numero_nf": d["numero_nf"],
                "serie": d["serie"],
                "data_bip1": dt_str,
                "hora_bip1": hr_str
            }
        })

    dt1 = parse_dt_banco(nota["dt_completa_bip1"])
    if dt1 is None:
        dt1 = agora

    tempo_txt, mins = calcular_diferenca(dt1, agora)

    try:
        qtd = int(nota["qtd_bipagens"]) if nota["qtd_bipagens"] is not None else 1
    except Exception:
        qtd = 1
    qtd += 1

    c.execute("""
        UPDATE notas
        SET data_bip2=?, hora_bip2=?, dt_completa_bip2=?,
            tempo_decorrido=?, minutos_decorridos=?, status=?, qtd_bipagens=?
        WHERE chave=?
    """, (dt_str, hr_str, iso_str, tempo_txt, mins, "DESVIO PENDENTE JUSTIFICATIVA", qtd, d["chave"]))
    conn.commit()
    conn.close()

    return jsonify({
        "sucesso": True,
        "tipo": "2ª Bipagem",
        "dados": {
            "chave": d["chave"],
            "numero_nf": d["numero_nf"],
            "serie": d["serie"],
            "data_bip1": nota["data_bip1"],
            "hora_bip1": nota["hora_bip1"],
            "data_bip2": dt_str,
            "hora_bip2": hr_str,
            "tempo_decorrido": tempo_txt
        }
    })


@app.route("/api/justificar", methods=["POST"])
def api_justificar():
    req = request.get_json(silent=True) or {}
    chave = (req.get("chave") or "").strip()
    justificativa = (req.get("justificativa") or "").strip()

    if not chave or not justificativa:
        return jsonify({"sucesso": False, "mensagem": "Chave e justificativa são obrigatórios."}), 400

    conn = obter_conexao()
    c = conn.cursor()
    c.execute("UPDATE notas SET justificativa = ?, status = 'DESVIO REGISTRADO' WHERE chave = ?", (justificativa, chave))
    c.execute("SELECT * FROM notas WHERE chave = ?", (chave,))
    row = c.fetchone()
    conn.commit()
    conn.close()

    if row:
        def disparar_email():
            cfg = carregar_config()
            destinatarios = cfg.get("destinatarios", [])
            if destinatarios:
                payload = {
                    "chave": chave,
                    "numero_nf": row["numero_nf"],
                    "serie": row["serie"],
                    "data_bip1": row["data_bip1"],
                    "hora_bip1": row["hora_bip1"],
                    "data_bip2": row["data_bip2"],
                    "hora_bip2": row["hora_bip2"],
                    "tempo_decorrido": row["tempo_decorrido"],
                    "justificativa": justificativa,
                }
                assunto = f"[ALERTA DE DESVIO - JUSTIFICADO] NF {row['numero_nf']}"
                html = template_email_desvio(payload)
                enviar_email_smtp(assunto, html, destinatarios)

        threading.Thread(target=disparar_email, daemon=True).start()

    return jsonify({"sucesso": True})


@app.route("/api/historico")
def api_historico():
    conn = obter_conexao()
    rows = conn.execute("SELECT * FROM notas ORDER BY dt_completa_bip1 DESC").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/exportar")
def api_exportar():
    conn = obter_conexao()
    registros = conn.execute("SELECT * FROM notas ORDER BY dt_completa_bip1 DESC").fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Portaria"

    headers = [
        "Chave de Acesso", "Número NF", "Série", "Data 1ª Bip", "Hora 1ª Bip",
        "Data 2ª Bip", "Hora 2ª Bip", "Diferença", "Status", "Justificativa"
    ]
    ws.append(headers)

    for r in registros:
        ws.append([
            r["chave"], r["numero_nf"], r["serie"],
            r["data_bip1"], r["hora_bip1"],
            r["data_bip2"] or "-", r["hora_bip2"] or "-",
            r["tempo_decorrido"] or "-", r["status"] or "-", r["justificativa"] or "-"
        ])

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)

    nome = f"Relatorio_Portaria_{agora_local().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        stream,
        as_attachment=True,
        download_name=nome,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# Inicializa banco no import
inicializar_banco()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
