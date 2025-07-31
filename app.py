from flask import Flask, request, render_template, send_file
from datetime import datetime
import tempfile
import os
from processamento import processar_planilha

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        if not file.filename.endswith(".xlsx"):
            return "Formato inválido. Envie um arquivo .xlsx"

        # Obtém a hora de referência, se fornecida
        hora_ref = request.form.get("hora_ref")
        hora_ref = int(hora_ref) if hora_ref and hora_ref.isdigit() else None

        # Salva o arquivo enviado em um arquivo temporário
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        file.save(temp_input.name)

        # Processa a planilha e gera o relatório formatado
        saida_path = processar_planilha(temp_input.name, hora_ref)

        # Remove o arquivo temporário de entrada
        os.unlink(temp_input.name)

        return send_file(saida_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
