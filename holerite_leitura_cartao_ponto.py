from flask import Flask, request, send_file, jsonify
import tempfile
import os
import fitz  # PyMuPDF
import pandas as pd
import re

app = Flask(__name__)

def processar_cartao_ponto(pdf_path):
    doc = fitz.open(pdf_path)
    texto_completo = "\n".join([page.get_text() for page in doc])

    padrao_linha = re.compile(
        r"(?P<data>\d{2}/\d{2}/\d{4})\s+(?P<dia_semana>\w{3}[-\w]*)\s+"
        r"(?P<marca>\d{2}:\d{2}[rgc]?(?:\s+\d{2}:\d{2}[rgc]?)*)(.*?)"
        r"\s*(?P<eventos>(?:\d{2}:\d{2}\s+\d{3}\s+.+?-?\s*)+)"
    )

    padrao_eventos = re.compile(r"(?P<tempo>\d{2}:\d{2})\s+(?P<codigo>\d{3})\s+(?P<descricao>.+?)-")

    registros = []

    for match in padrao_linha.finditer(texto_completo):
        data = match.group("data")
        dia_semana = match.group("dia_semana")
        marca = match.group("marca").replace('\n', ' ').strip()
        eventos_raw = match.group("eventos")

        eventos = padrao_eventos.findall(eventos_raw)
        evento_dict = {}

        for tempo, codigo, descricao in eventos:
            key = descricao.strip().upper()
            if key in evento_dict:
                h, m = map(int, tempo.split(":"))
                old_h, old_m = map(int, evento_dict[key].split(":"))
                total_min = h * 60 + m + old_h * 60 + old_m
                evento_dict[key] = f"{total_min // 60:02}:{total_min % 60:02}"
            else:
                evento_dict[key] = tempo

        registro = {
            "Data": data,
            "Dia da Semana": dia_semana,
            "Marcações": marca,
        }
        registro.update(evento_dict)
        registros.append(registro)

    df = pd.DataFrame(registros)
    df[["Dia da Semana (3 letras)", "Tipo de Dia"]] = df["Dia da Semana"].str.extract(r"(\w{3})-(.*)")
    marcacoes_expandido = df["Marcações"].str.split(r'\s+', expand=True)
    marcacoes_expandido.columns = [f"Marcação {i+1}" for i in range(marcacoes_expandido.shape[1])]
    df_final = pd.concat([df.drop(columns=["Marcações", "Dia da Semana"]), marcacoes_expandido], axis=1)

    return df_final

@app.route('/')
def home():
    return "API de Cartão Ponto Online"

@app.route('/processar-cartao-ponto', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400

    file = request.files['file']
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        file.save(tmp_file.name)
        pdf_path = tmp_file.name

    try:
        df_final = processar_cartao_ponto(pdf_path)

        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        df_final.to_excel(output_path, index=False)

        return send_file(output_path,
                         as_attachment=True,
                         download_name="cartao_ponto_estruturado.xlsx",
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.remove(pdf_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
