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
            status TEXT, justificativa TEXT, qtd_bipagens INTEGER DEFAULT 1,
            observacao TEXT, email_desvio_enviado INTEGER DEFAULT 0
        )
    """)
    conn.commit()
    
    # Migração para garantir que bancos antigos não quebrem
    c.execute("PRAGMA table_info(notas)")
    cols = [x["name"] for x in c.fetchall()]
    needed = {
        "data_bip2": "TEXT", "hora_bip2": "TEXT", "dt_completa_bip2": "TEXT",
        "tempo_decorrido": "TEXT", "minutos_decorridos": "REAL",
        "status": "TEXT", "justificativa": "TEXT", "qtd_bipagens": "INTEGER DEFAULT 1",
        "observacao": "TEXT", "email_desvio_enviado": "INTEGER DEFAULT 0"
    }
    for col, typ in needed.items():
        if col not in cols:
            try: c.execute(f"ALTER TABLE notas ADD COLUMN {col} {typ}")
            except Exception: pass
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

def carregar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return {"destinatarios": [], "smtp_server": "smtp.office365.com", "smtp_port": 587, "email_remetente": "", "senha_app": ""}

def enviar_email_smtp(assunto, corpo_html, destinatarios):
    cfg = carregar_config()
    remetente = cfg.get("email_remetente", "").strip()
    senha = cfg.get("senha_app", "").strip()
    srv = cfg.get("smtp_server", "smtp.office365.com").strip()
    porta = int(cfg.get("smtp_port", 587))
    if not remetente or not senha or not destinatarios: return False
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
    except: return False

def template_email_desvio(
