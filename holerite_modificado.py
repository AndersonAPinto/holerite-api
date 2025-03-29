from flask import Flask, request, send_file, jsonify
import tempfile
import os
import fitz  # PyMuPDF
import pandas as pd
import re

app = Flask(__name__)

def extrair_lancamentos(texto, nome):
    lancamentos = []
    tipo_atual = "provento"

    for linha in texto.splitlines():
        linha = linha.strip()

        if "TOTAL DE PROVENTOS" in linha:
            tipo_atual = "desconto"
            continue

        match = re.match(r"^(\d{4})\s{2,}(.+?)\s{2,}([\d.,]{1,7})\s{2,}([\d.,]{1,10})$", linha)
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

def extrair_dados_pessoais(texto):
    dados = {}

    match_nome = re.search(r"(\d{10})\s+([A-ZÇÀ-Ú ]+)\s+Banco", texto)
    if match_nome:
        dados["matricula"] = match_nome.group(1)
        dados["nome"] = match_nome.group(2).strip()

    match_cnpj = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    if match_cnpj:
        dados["cnpj"] = match_cnpj.group(1)

    match_admissao = re.search(r"Admitido em:\s+(\d{2}/\d{2}/\d{4})", texto)
    if match_admissao:
        dados["admissao"] = match_admissao.group(1)

    match_banco = re.search(r"Banco:\s*(.*)", texto)
    if match_banco:
        dados["banco"] = match_banco.group(1).strip()

    match_salario = re.search(r"Sal[aá]rio Pago:\s*(\d+,\d{2})", texto)
    if match_salario:
        dados["salario_pago"] = match_salario.group(1)

    match_fgts_base = re.search(r"Base para FGTS\s+([\d.,]+)", texto)
    if match_fgts_base:
        dados["base_fgts"] = match_fgts_base.group(1)

    match_fgts_mes = re.search(r"FGTS do m[eê]s\s+([\d.,]+)", texto)
    if match_fgts_mes:
        dados["fgts_mes"] = match_fgts_mes.group(1)

    match_total_proventos = re.search(r"Total de Proventos\s+([\d.,]+)", texto)
    if match_total_proventos:
        dados["total_proventos"] = match_total_proventos.group(1)

    match_total_descontos = re.search(r"Total de Descontos\s+([\d.,]+)", texto)
    if match_total_descontos:
        dados["total_descontos"] = match_total_descontos.group(1)

    match_liquido = re.search(r"L[ií]quido a Receber =>\s+([\d.,]+)", texto)
    if match_liquido:
        dados["liquido"] = match_liquido.group(1)

    match_conta = re.search(r"Ag/Conta:\s*/\s*(\d{4,})", texto)
    if match_conta:
        dados["conta"] = match_conta.group(1)

    return dados

@app.route('/')
def home():
    return "API de holerites estruturada com dados pessoais e lançamentos!"

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
        dados_pessoais_list = []
        lancamentos_list = []

        for page in doc:
            text = page.get_text()
            dados_pessoais = extrair_dados_pessoais(text)
            if "nome" in dados_pessoais:
                lancamentos = extrair_lancamentos(text, dados_pessoais["nome"])
                lancamentos_list.extend(lancamentos)
            dados_pessoais_list.append(dados_pessoais)

        df_dados_pessoais = pd.DataFrame(dados_pessoais_list)
        df_lancamentos = pd.DataFrame(lancamentos_list)

        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_dados_pessoais.to_excel(writer, sheet_name="Dados Pessoais", index=False)
            df_lancamentos.to_excel(writer, sheet_name="Demonstrativo", index=False)

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
