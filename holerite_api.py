from flask import Flask, request, send_file, jsonify
import tempfile
import os
import fitz
import pandas as pd
import re

app = Flask(__name__)

def extract_items(table_text, tipo, ref_nome):
    items = []
    lines = table_text.strip().split('\n')
    for line in lines:
        match = re.match(r"(\d{2,5})\s+(.+?)\s+([\d.,]+)", line)
        if match:
            items.append({
                "tipo": tipo,
                "codigo": match.group(1),
                "descricao": match.group(2).strip(),
                "valor": match.group(3).replace('.', '').replace(',', '.'),
                "referencia_nome": ref_nome
            })
    return items

def infer_dynamic_fields(text):
    campos = {}
    cpf_match = re.search(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b", text)
    if cpf_match:
        campos["cpf"] = cpf_match.group()
    nome_match = re.search(r"Nome\s*\n?(.+)", text, re.IGNORECASE)
    if nome_match:
        campos["nome"] = nome_match.group(1).strip()
    elif cpf_match:
        before_cpf = text[:cpf_match.start()]
        nome_alt = re.findall(r"[A-Z][A-Z\s]+", before_cpf)
        if nome_alt:
            campos["nome"] = nome_alt[-1].strip()
    matricula_match = re.search(r"Mat[ríi]cula\s+(\d{5,})", text, re.IGNORECASE)
    if matricula_match:
        campos["matricula"] = matricula_match.group(1)
    cargo_match = re.search(r"Cargo\s+(.+)", text)
    if cargo_match:
        campos["cargo"] = cargo_match.group(1).strip()
    total_prov = re.search(r"TOTAL\s+DE\s+PROVENTOS\s+([\d.,]+)", text, re.IGNORECASE)
    if total_prov:
        campos["total_proventos"] = total_prov.group(1)
    total_desc = re.search(r"TOTAL\s+DE\s+DESCONTOS\s+([\d.,]+)", text, re.IGNORECASE)
    if total_desc:
        campos["total_descontos"] = total_desc.group(1)
    liquido = re.search(r"L[ií]quido\s+a\s+Receber\s+([\d.,]+)", text, re.IGNORECASE)
    if liquido:
        campos["liquido"] = liquido.group(1)
    ref_match = re.search(r"Refer[êe]ncia\s+(\d{2}/\d{4})", text)
    if ref_match:
        campos["referencia"] = ref_match.group(1)
    return campos

def process_pdf_adaptativo(filepath):
    doc = fitz.open(filepath)
    resumo_list = []
    detalhamento_list = []

    for i, page in enumerate(doc):
        text = page.get_text()
        campos = infer_dynamic_fields(text)
        resumo_list.append(campos)

        nome_ref = campos.get("nome", f"pagina_{i+1}")

        blocos = re.split(r"C[oó]digo\s+Descri[cç][aã]o\s+Valor", text, flags=re.IGNORECASE)
        tabelas = []

        for bloco in blocos[1:]:
            linhas = bloco.strip().split("\n")
            conteudo = []
            for linha in linhas:
                if re.match(r"\d{2,5}\s+.+?\s+[\d.,]+", linha.strip()):
                    conteudo.append(linha)
                else:
                    break
            tabelas.append("\n".join(conteudo))

        if len(tabelas) > 0:
            detalhamento_list.extend(extract_items(tabelas[0], "provento", nome_ref))
        if len(tabelas) > 1:
            detalhamento_list.extend(extract_items(tabelas[1], "desconto", nome_ref))

    df_resumo = pd.DataFrame(resumo_list)
    df_detalhamento = pd.DataFrame(detalhamento_list)
    return df_resumo, df_detalhamento

@app.route('/processar-holerite', methods=['POST'])
def processar_holerite():
    if 'file' not in request.files:
        return jsonify({'error': 'Arquivo não enviado'}), 400

    file = request.files['file']

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        file.save(tmp_pdf.name)
        pdf_path = tmp_pdf.name

    try:
        df_resumo, df_detalhamento = process_pdf_adaptativo(pdf_path)

        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_resumo.to_excel(writer, sheet_name='resumo', index=False)
            df_detalhamento.to_excel(writer, sheet_name='detalhamento', index=False)

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
