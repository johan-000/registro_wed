from flask import Flask, render_template, request, redirect
import openpyxl
from datetime import datetime
import os

app = Flask(__name__)
EXCEL_PATH = "registro.xlsx"

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        nombres = request.form["nombres"]
        documento = request.form["documento"]
        motivo = request.form["motivo"]
        eps = request.form["eps"]
        arl = request.form["arl"]
        autoriza = request.form["autoriza"]
        emergencia = request.form["emergencia"]
        celular = request.form["celular"]
        observaciones = request.form["observaciones"]
        hora = datetime.now().strftime("%H:%M")
        fecha = datetime.now().strftime("%d/%m/%Y")

        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active
        fila = 8
        while ws[f"A{fila}"].value:
            fila += 1

        datos = [*fecha.split("/"), hora, nombres, documento, motivo, eps, arl, autoriza, emergencia, celular, "", "", observaciones]
        for i, valor in enumerate(datos):
            col = chr(ord("A") + i)
            ws[f"{col}{fila}"] = valor

        wb.save(EXCEL_PATH)

        return redirect("/")

    return render_template("formulario.html")

@app.route("/registros")
def mostrar_registros():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    registros = []

    fila = 8
    while ws[f"A{fila}"].value:
        fila_data = [ws[f"{chr(c)}{fila}"].value for c in range(ord("A"), ord("O")+1)]
        registros.append(fila_data)
        fila += 1

    return render_template("registros.html", registros=registros)

if __name__ == "__main__":
    app.run(debug=True)
