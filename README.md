# 📄 Extrator de NFs em PDF

Este projeto é um **script em Python** que automatiza a extração de dados de **notas fiscais eletrônicas (NF-e)** salvas em **arquivos PDF**, gerando como resultado uma **planilha `.xlsx`** com todas as informações estruturadas.

---

## ⚙️ Funcionalidades

- 📂 Lê múltiplos arquivos PDF contendo notas fiscais.
- 🔍 Extrai automaticamente os principais dados das notas.
- 📊 Organiza as informações em colunas de uma planilha `.xlsx`.
- 📎 Gera um arquivo auxiliar com os PDFs unificados para facilitar auditorias.

---

## 🛠️ Bibliotecas Utilizadas

Para o funcionamento do script, as seguintes bibliotecas são necessárias:

```python
import pdfplumber
import pandas as pd
import os
import re
import glob
from datetime import datetime
```

> ⚠️ Certifique-se de que todas essas bibliotecas estão instaladas no seu ambiente Python.

---

## 🚀 Como Usar

### 1️⃣ Defina o caminho da pasta onde estão os arquivos PDF

No início do script, configure a variável `PASTA_RAIZ` com o caminho da pasta que contém os arquivos PDF das notas fiscais:

```python
PASTA_RAIZ = r'C:\Users\Igor\Desktop\Nova pasta\PDFs_Nao_Processados'
```

### 2️⃣ Escolha o nome do arquivo de saída

Configure a variável `NOME_ARQUIVO_SAIDA` com o nome que deseja dar à planilha gerada:

```python
NOME_ARQUIVO_SAIDA = 'dados_notas_fiscais_vfinal_corrigido.xlsx'
```

---

## 📊 Estrutura da Planilha Gerada

A planilha `.xlsx` conterá os seguintes campos organizados em colunas:

- **Nome do Arquivo de Origem**
- **Nome do Emitente**
- **Data de Emissão**
- **Número da NF**
- **Destinatário**
- **Descrição do Item**
- **Quantidade**
- **Valor Unitário**
- **Valor Total dos Itens**
- **Valor Total da NF**

---

## 🧾 Arquivo Extra de Auditoria

Além da planilha, o script também pode gerar um **arquivo PDF unificado** com todas as notas processadas, facilitando a revisão e auditoria das informações.

---

## 📌 Observações

- Este script foi desenvolvido para notas fiscais em formato PDF com estrutura legível por OCR ou com texto extraível.
- Pode ser adaptado para diferentes layouts ou campos adicionais, se necessário.

---

## 🤝 Contribuições

Contribuições são bem-vindas! Fique à vontade para abrir issues ou enviar pull requests.

---
