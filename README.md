Gerador de Orçamento com Streamlit e Outlook
Este projeto é uma aplicação web que permite ao usuário preencher um formulário para gerar um orçamento. O orçamento pode ser salvo como um arquivo PDF e também pode ser enviado por e-mail usando o Outlook.

Funcionalidades
Geração de Orçamento: O usuário pode preencher um formulário com informações sobre o orçamento, como nome, endereço, telefone, e-mail, descrição do projeto, horas estimadas, valor da hora trabalhada e prazo.
Salvamento em PDF: O orçamento gerado pode ser salvo como um arquivo PDF.
Envio por E-mail: O orçamento pode ser enviado por e-mail usando o Outlook, incluindo a assinatura configurada no Outlook.
Pré-requisitos
Python 3.x
Bibliotecas Python: streamlit, win32com.client, fpdf, os, tkinter, markdown
Instalação
Clone o repositório:


git clone https://github.com/seu-usuario/seu-repositorio.git
cd seu-repositorio
Instale as dependências:


pip install -r requirements.txt
Uso
Execute o script principal:


streamlit run main.py
Abra o navegador e acesse a URL fornecida pelo Streamlit (geralmente http://localhost:8501).

Preencha o formulário com as informações do orçamento.

Clique no botão "Gerar Orçamento" para gerar o orçamento e salvar como PDF.

Clique no botão "Enviar Orçamento por E-mail" para enviar o orçamento por e-mail usando o Outlook.


Contribuição
Contribuições são bem-vindas! Por favor, abra uma issue ou envie um pull request para sugerir melhorias ou corrigir problemas.

Licença
Este projeto está licenciado sob a Licença MIT. Veja o arquivo LICENSE para mais detalhes.