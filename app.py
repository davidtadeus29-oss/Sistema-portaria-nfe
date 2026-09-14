app = Flask(__name__)
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
            <span class="text-white-50 fw-bold" id="clock"></span>
        </div>
    </nav>

    <div class="container-fluid px-4">
        <div class="card p-4 mb-4 border-0 shadow-sm">
            <h5 class="text-secondary fw-bold mb-3">⚡ Leitura de Código de Barras (Chave de 44 dígitos)</h5>
            <div class="input-group mb-3">
                <input type="text" id="chaveInput" class="form-control barra-input text-center" placeholder="Aponte o leitor de código de barras aqui..." autofocus autocomplete="off">
                <button id="btnRegistrar" type="button" class="btn btn-primary px-4 fw-bold">Registrar</button>
                <button id="btnLimpar" type="button" class="btn btn-outline-secondary px-4 fw-bold">Limpar</button>
            </div>
            <div id="painel" class="alert alert-light border text-center my-0 py-3 fw-bold text-success">🟢 Leitor Pronto para Bipagem</div>
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

    <!-- Modal Justificativa -->
    <div class="modal fade" id="modalJust" tabindex="-1" data-bs-backdrop="static">
      <div class="modal-dialog modal-dialog-centered">
        <div class="modal-content">
            <div class="modal-header bg-danger text-white">
                <h5 class="modal-title fw-bold">⚠️ DUPLO REGISTRO IDENTIFICADO</h5>
            </div>
            <div class="modal-body">
                <p id="descModal" class="mb-3"></p>
                <label class="form-label fw-bold">Informe o motivo da 2ª bipagem (Obrigatório):</label>
                <textarea id="justInput" class="form-control" rows="3" placeholder="Digite aqui..."></textarea>
            </div>
            <div class="modal-footer">
                <button id="btnSalvarJust" type="button" class="btn btn-danger fw-bold w-100">Gravar Justificativa e Enviar E-mail</button>
            </div>
        </div>
      </div>
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

        setInterval(() => {
            document.getElementById('clock').innerText =
                new Date().toLocaleDateString() + ' ' + new Date().toLocaleTimeString();
        }, 1000);

        function focar(){
            inputElem.focus();
            inputElem.select();
        }

        function limpar(){
            inputElem.value = "";
            ultChave = "";
            focar();
        }

        async function executarRequisicao(chave){
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

        function biparManual(){
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

        function biparAutomatico(){
            if (processando) return;
            const chave = inputElem.value.trim();
            if (!chave) return;

            const now = Date.now();
            if (chave === ultChave && (now - ultTime < 2500)) return;

            ultChave = chave;
            ultTime = now;
            executarRequisicao(chave);
        }

        async function salvarJust(){
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

        async function carregar(){
            try {
                const res = await fetch('/api/historico');
                const lista = await res.json();
                corpo.innerHTML = "";

                lista.forEach(n => {
                    let st = "";
                    if ((n.status || "").includes("DESVIO")) st = '<span class="badge bg-danger">🚨 DESVIO</span>';
                    else if ((n.status || "").includes("PENDENTE")) st = '<span class="badge bg-warning text-dark">⚠️ PENDENTE</span>';
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

        // Bind de eventos (sem onclick inline)
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
</html>"""
application = app
