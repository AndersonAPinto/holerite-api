from flask import Flask, request, send_file, jsonify
import tempfile
import os
import fitz  # PyMuPDF
import pandas as pd
import re

app = Flask(__name__)

def extrair_lancamentos(texto, nome):
    lancamentos = []
    capturando = False
    tipo_atual = "provento"

    for linha in texto.splitlines():
        linha = linha.strip()

        if "TOTAL DE PROVENTOS" in linha:
            tipo_atual = "desconto"
            continue

        match = re.match(r"^(\d{4})\s{2,}(.+?)\s{2,}([\d.,]{1,7})\s{2,}([\d.,]{1,10})$", linha) #r"^(\d{4})\s+(.+?)\s+(\d{1,3},\d{2})\s+(\d{1,3},\d{2})$"
        if match:
            lancamentos.append({
                "tipo": tipo_atual,
                "codigo": match.group(1),
                "descricao": match.group(2).strip(),
                "referencia": match.group(3),
                "valor": match.group(4).replace('.', '').replace(',', '.'),
                "nome": nome
            })

    return lancamentos

def extrair_resumo(texto):
    resumo = {}

    # Nome e matrícula (bloco do topo)
    match_nome = re.search(r"\b(\d{10})\s+([A-ZÇÀ-Ú ]{5,})\s+Banco", texto)
    if match_nome:
        resumo["matricula"] = match_nome.group(1).strip()
        resumo["nome"] = match_nome.group(2).strip()

    match_funcao = re.search(r"Fun[cç][aã]o\s+([\w\s]+)", texto)
    if match_funcao:
        resumo["funcao"] = match_funcao.group(1).strip()

    match_banco = re.search(r"Banco\s+(.+)", texto)
    if match_banco:
        resumo["banco"] = match_banco.group(1).strip()

    match_admissao = re.search(r"Admitido em\s+(\d{2}/\d{2}/\d{4})", texto)
    if match_admissao:
        resumo["admitido_em"] = match_admissao.group(1)

    match_proventos = re.search(r"TOTAL DE PROVENTOS\s+([\d.,]+)", texto)
    if match_proventos:
        resumo["total_proventos"] = match_proventos.group(1).replace('.', '').replace(',', '.')

    match_descontos = re.search(r"TOTAL DE DESCONTOS\s+([\d.,]+)", texto)
    if match_descontos:
        resumo["total_descontos"] = match_descontos.group(1).replace('.', '').replace(',', '.')

    match_liquido = re.search(r"L[ií]quido\s+a\s+Receber\s+([\d.,]+)", texto)
    if match_liquido:
        resumo["liquido"] = match_liquido.group(1).replace('.', '').replace(',', '.')

    return resumo

@app.route('/')
def home():
    return "API de holerites estruturada funcionando!"

@app.route('/processar-holerite', methods=['POST'])
def processar_holerite():
    if 'file' not in request.files:
        return jsonify({'error': 'Arquivo não enviado'}), 400

    file = request.files['file']
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        file.save(tmp_pdf.name)
        pdf_path = tmp_pdf.name

    try:
        doc = fitz.open(pdf_path)
        resumos = []
        detalhes = []

        for page in doc:
            text = page.get_text()
            resumo = extrair_resumo(text)
            if "nome" in resumo:
                detalhes += extrair_lancamentos(text, resumo["nome"])
            resumos.append(resumo)

        df_resumo = pd.DataFrame(resumos)
        df_lancamentos = pd.DataFrame(detalhes)

        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_resumo.to_excel(writer, sheet_name="resumo_colaborador", index=False)
            df_lancamentos.to_excel(writer, sheet_name="proventos_descontos", index=False)

        return send_file(output_path,
                         as_attachment=True,
                         download_name="holerites_estruturado.xlsx",
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.remove(pdf_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
