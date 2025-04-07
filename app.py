import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO
import os

st.set_page_config(page_title="Checklist Fiscalização", page_icon="✅")

st.title("📋 Checklist de Fiscalização")

# Nome do arquivo Excel
arquivo_excel = "checklists.xlsx"

# Perguntas do checklist
perguntas = [
    "Qual a empresa?",
    "Qual a data?",
    "Qual atividade?",
    "Qual Local?",
    "Qual OM?",
    "Isolamento e APR ok?",
    "Epi's ok?",
    "N° de funcionários OK?",
    "Atividades da OM foram executadas?",
    "Recursos Samarco disponíveis?",
    "Atividade tem início conforme programado?"
]

# Coleta de respostas via formulário Streamlit
with st.form("checklist_form"):
    st.subheader("Preencha o checklist:")
    respostas = [st.text_input(pergunta) for pergunta in perguntas]
    enviado = st.form_submit_button("Salvar no Excel")

def formatar_planilha(ws):
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 50

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    for col in ["A", "B"]:
        ws[f"{col}1"].font = header_font
        ws[f"{col}1"].fill = header_fill
        ws[f"{col}1"].alignment = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.border = thin_border

# Salvar respostas no Excel
if enviado:
    if os.path.exists(arquivo_excel):
        wb = openpyxl.load_workbook(arquivo_excel)
    else:
        wb = Workbook()
        del wb["Sheet"]

    nome_planilha = f"{respostas[0]}_{respostas[4]}"
    if nome_planilha in wb.sheetnames:
        num = 1
        while f"{nome_planilha}_{num}" in wb.sheetnames:
            num += 1
        nome_planilha = f"{nome_planilha}_{num}"

    ws = wb.create_sheet(title=nome_planilha)
    ws.append(["Pergunta", "Resposta"])
    for pergunta, resposta in zip(perguntas, respostas):
        ws.append([pergunta, resposta])

    formatar_planilha(ws)

    # Salvar em memória para download
    excel_bytes = BytesIO()
    wb.save(excel_bytes)
    excel_bytes.seek(0)

    st.success(f"Checklist salvo com sucesso na planilha '{nome_planilha}'!")

    st.download_button(
        label="📥 Baixar Excel atualizado",
        data=excel_bytes,
        file_name=arquivo_excel,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )