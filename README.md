# Controle de Recebimento de Materiais

Aplicativo Streamlit que lê diretamente o arquivo Excel do repositório.

## Arquivos obrigatórios na raiz

- `App_Recebimento.py`
- `Controle de Recebimento de Materiais.xlsx`
- `requirements.txt`

## Publicação no Streamlit Community Cloud

1. Crie ou abra um repositório no GitHub.
2. Envie os três arquivos obrigatórios para a raiz.
3. No Streamlit Community Cloud, crie um app apontando para o repositório.
4. Em **Main file path**, informe `App_Recebimento.py`.
5. Clique em **Deploy**.

## Atualização da base

Substitua no GitHub apenas `Controle de Recebimento de Materiais.xlsx`, preservando exatamente o nome do arquivo, a aba `Recebimento Manual` e os cabeçalhos. Após o commit, o aplicativo será atualizado. Se o cache permanecer, reinicie o app no Streamlit Cloud.

## Observação

O caminho da base é resolvido a partir da pasta do próprio código, portanto funciona no GitHub/Streamlit Cloud sem depender da pasta local do computador.
