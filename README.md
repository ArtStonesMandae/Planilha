
# Gerador de Planilhas Mandae

Este é um app Streamlit para gerar planilhas no modelo da transportadora Mandae a partir de arquivos .csv.

## Como usar

1. Faça upload do seu arquivo CSV com os pedidos.
2. O app gera automaticamente a planilha `.xlsx` no modelo visual e estrutural exigido pela Mandae.
3. Baixe o arquivo e envie para a transportadora.

## Como publicar no Streamlit Cloud

1. Crie uma conta gratuita em: https://share.streamlit.io
2. Crie um repositório no seu GitHub com os seguintes arquivos:
   - `app_mandae.py`
   - `requirements.txt`
3. No Streamlit Cloud:
   - Clique em **"New app"**
   - Conecte com seu GitHub e selecione o repositório
   - Escolha o arquivo `app_mandae.py`
   - Clique em **"Deploy"**

Pronto! Agora você pode acessar o app direto do navegador.
