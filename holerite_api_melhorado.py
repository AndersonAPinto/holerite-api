from flask import Flask, request, send_file, jsonify
import tempfile
import os
import fitz
import pandas as pd
import re
from collections import Counter

app = Flask(__name__)

def parser_heuristico_melhorado(text):
    campos = {}
    linhas = text.splitlines()

    for i, linha in enumerate(linhas):
        linha = linha.strip()
        if not linha or len(linha) < 3:
            continue

        match = re.match(r"^([A-ZÀ-Úa-zçÇ\s/\-]{2,})[:\s]{1,}(.+)$", linha)
        if match:
            chave = match.group(1).strip().lower().replace(" ", "_").replace("-", "_").replace("/", "_")
            valor = match.group(2).strip()
            if chave not in campos:
                campos[chave] = valor
            continue

        if i < len(linhas) - 1:
            prox = linhas[i + 1].strip()
            if re.match(r"^[A-ZÀ-Úa-zçÇ\s/\-]{3,}$", linha) and re.match(r".{3,}", prox):
                chave = linha.lower().replace(" ", "_").replace("-", "_").replace("/", "_")
                if chave not in campos:
                    campos[chave] = prox

    return campos

def process_pdf_dinamico_melhorado(filepath):
    doc = fitz.open(filepath)
    registros = []
    field_counter = Counter()

    for i, page in enumerate(doc):
        text = page.get_text()
        campos = parser_heuristico_melhorado(text)
        registros.append(campos)
        field_counter.update(campos.keys())

    campos_frequentes = [k for k, v in field_counter.items() if v >= 1] # Reduzir exigência tornando mais flexivel.
    linhas_normalizadas = []
    for reg in registros:
        linha = {k: reg.get(k, None) for k in campos_frequentes}
        linhas_normalizadas.append(linha)

    df_final = pd.DataFrame(linhas_normalizadas)
    return df_final

@app.route('/')
def home():
    return "🚀 API de processamento de holerite online!"

@app.route('/processar-holerite/', methods=['POST'])
def processar_holerite():
    if 'file' not in request.files:
        return jsonify({'error': 'Arquivo não enviado'}), 400

    file = request.files['file']

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        file.save(tmp_pdf.name)
        pdf_path = tmp_pdf.name

    try:
        df_final = process_pdf_dinamico_melhorado(pdf_path)

        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)

        return send_file(output_path,
                         as_attachment=True,
                         download_name="planilha_holerite.xlsx",
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.remove(pdf_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
