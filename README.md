# Sistema de Geração de Recibos de Viagem
## Introdução
Essa automação fornece a criação de recibos de viagem em PDF e Word a partir de uma planilha Excel. Este documento fornece um guia detalhado sobre como usar o sistema de geração de recibos a partir de dados contidos em uma planilha Excel. O sistema permite selecionar arquivos Excel e Word, preencher modelos de documentos Word com dados específicos e convertê-los para PDF.

## Requisitos do Sistema
- Python 3.x instalado

## Bibliotecas Python necessárias:
- os
- sys
- tkinter
- openpyxl
- python-docx
- docx2pdf
- datetime

## Funcionalidades
O sistema possui as seguintes funcionalidades:

- Seleção de Arquivos: Permite ao usuário selecionar um arquivo Excel, um arquivo Word e uma pasta de destino para salvar os documentos gerados.
- Leitura e Processamento de Dados: Lê os dados da planilha Excel e preenche um modelo de documento Word com esses dados.
- Geração de Documentos: Cria documentos Word personalizados e os converte em PDFs, salvando-os na pasta de destino.

## Passo a Passo para Utilização
1. Preencher a planilha base com os dados do recibo
Lembre-se: o máximo de itens em cada recibo é igual a 5, respeitando a quantidade de colunas. As colunas podem ser deixadas em branco, e cada linha cria um novo recibo.

2. Baixar Python
Baixe Python [aqui](https://www.python.org/downloads/)


 2.1. Instalar Python
Durante a instalação, lembre-se de marcar as caixinhas de seleção necessárias. Caso seu computador não permita a instalação, entre em contato com o setor de TI ou abra um chamado para baixar o programa.

3. Atualizando Bibliotecas
Execute o programa Atualizar Bibliotecas para garantir que todas as dependências estejam atualizadas.

4. Executar Automação
Após as bibliotecas serem atualizadas, execute o script main.py. Caso apareça uma mensagem de alerta por fornecedor desconhecido, clique em "saber mais" e "executar assim mesmo".

5. Seleção dos Arquivos
Ao executar o script, você será solicitado a selecionar os seguintes arquivos:

- Arquivo Excel: Contendo os dados dos recibos preenchidos no passo 1.
- Arquivo Word: Modelo do documento base Word (não pode ser alterado).
 - Pasta de Destino: Local onde os documentos gerados serão salvos.
6. Resultado
A automação lerá os dados do arquivo Excel e preencherá o modelo de documento Word para cada linha da planilha, ignorando a primeira linha (cabeçalho). Os documentos gerados serão salvos na pasta de destino e convertidos para Word e PDF automaticamente.
## Autores

- [@IasminCQFernandes](https://www.github.com/iasmincqfernandes)
- [@Gustavo-gcr](https://github.com/Gustavo-gcr)

