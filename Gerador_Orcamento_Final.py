import streamlit as st
import win32com.client as win32
from fpdf import FPDF
import os
import tkinter as tk
from tkinter import filedialog

# Função para abrir a janela de diálogo de salvamento
def selecionar_caminho_pdf():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do tkinter
    caminho_pdf = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        title="Salvar Arquivo"
    )
    return caminho_pdf

# Função para validar os campos de entrada
def validar_campos(nome, endereco, telefone, email, descricao, projeto, horas_estimadas, valor_hora, prazo):
    if not all([nome, endereco, telefone, email, descricao, projeto, horas_estimadas, valor_hora, prazo]):
        return "Preencha todos os campos!", None

    try:
        float(horas_estimadas)
        float(valor_hora)
    except ValueError:
        return "Horas estimadas e valor da hora devem ser números!", None

    return None, (nome, endereco, telefone, email, descricao, projeto, horas_estimadas, valor_hora, prazo)

# Função para gerar o conteúdo do orçamento
def gerar_conteudo_orcamento(nome, endereco, telefone, email, descricao, projeto, horas_estimadas, valor_hora, prazo):
    valor_total = float(horas_estimadas) * float(valor_hora)
    orcamento = f"""
    Nome: {nome}
    Endereço: {endereco}
    Telefone: {telefone}
    E-mail: {email}
    Descrição: {descricao}
    Nome do projeto: {projeto}
    Horas estimadas: {horas_estimadas}
    Valor da hora trabalhada: {valor_hora}
    Prazo: {prazo}
    Valor total: {valor_total}
    """
    return orcamento, valor_total

# Função para gerar o PDF
def gerar_pdf(caminho_pdf, orcamento):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, "Orçamento", 1, 1, "C")
    pdf.ln(10)
    pdf.multi_cell(0, 10, orcamento)

    try:
        pdf.output(caminho_pdf, "F")
        if os.path.exists(caminho_pdf):
            return caminho_pdf
        else:
            return None
    except Exception as e:
        st.error(f"Erro ao gerar o PDF: {str(e)}")
        return None

# Função para enviar o e-mail com o PDF anexado
def enviar_email(email_destinatario, projeto, descricao, caminho_pdf):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = email_destinatario
        mail.Subject = projeto
        mail.HTMLBody = descricao

        if os.path.isfile(caminho_pdf):
            try:
                mail.Attachments.Add(caminho_pdf)
                mail.Send()
                return True
            except Exception as e:
                st.error(f"Erro ao enviar e-mail: {str(e)}")
                return False
        else:
            st.error("Arquivo de anexo não encontrado.")
            return False
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {str(e)}")
        return False

def main():
    st.title("Gerar Orçamento")
    st.write("Preencha os campos abaixo para gerar um orçamento")

    nome = st.text_input("Nome")
    endereco = st.text_input("Endereço")
    telefone = st.text_input("Telefone")
    email_destinatario = st.text_input("E-mail")
    descricao = st.text_area("Descrição")
    projeto = st.text_input("Nome do projeto")
    horas_estimadas = st.text_input("Horas estimadas")
    valor_hora = st.text_input("Valor da hora trabalhada:")
    prazo = st.text_input("Prazo")

    if 'pdf_path' not in st.session_state:
        st.session_state.pdf_path = None

    if st.button("Gerar Orçamento"):
        erro, dados = validar_campos(nome, endereco, telefone, email_destinatario, descricao, projeto, horas_estimadas, valor_hora, prazo)
        if erro:
            st.error(erro)
        else:
            orcamento, valor_total = gerar_conteudo_orcamento(*dados)
            caminho_pdf = selecionar_caminho_pdf()
            if not caminho_pdf:
                st.error("Nenhum caminho de salvamento foi selecionado.")
            else:
                diretorio = os.path.dirname(caminho_pdf)
                if not os.path.exists(diretorio):
                    try:
                        os.makedirs(diretorio, exist_ok=True)
                    except Exception as e:
                        st.error(f"Erro ao criar diretório: {str(e)}")
                        return

                pdf_path = gerar_pdf(caminho_pdf, orcamento)
                if pdf_path:
                    st.session_state.pdf_path = pdf_path
                    st.write("Orçamento gerado:")
                    st.write(orcamento)
                    st.success(f"PDF salvo em: {pdf_path}")
                else:
                    st.error("Falha ao salvar o arquivo PDF.")

    if st.button("Enviar Orçamento por E-mail"):
        if st.session_state.pdf_path and os.path.exists(st.session_state.pdf_path):
            if enviar_email(email_destinatario, projeto, descricao, st.session_state.pdf_path):
                st.success("Orçamento enviado com sucesso!")
            else:
                st.error("Erro ao enviar e-mail")
        else:
            st.error("O arquivo PDF não foi encontrado ou não foi gerado ainda.")

if __name__ == "__main__":
    main()
