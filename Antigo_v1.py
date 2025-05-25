# -*- coding: utf-8 -*- # Define a codificação do arquivo como UTF-8

# --- Importação das Bibliotecas Necessárias ---
import pdfplumber
import pandas as pd
import os
import re
import glob # Importante para busca de arquivos
from datetime import datetime

# --- Configurações Iniciais ---
# >>> MODIFICADO: Defina a pasta PAI que contém as pastas dos meses <<<
PASTA_RAIZ = r'C:\Users\Igor\Desktop\Projeto Planilhas\NOTAS 2025\03969470188\00113906129\2025'
NOME_ARQUIVO_SAIDA = 'dados_notas_fiscais_COMPLETO.xlsx' # Nome para o arquivo consolidado


# --- Funções Auxiliares (MANTIDAS COMO ANTES) ---
def limpar_numero(texto_numero):
    # ...(Função robusta mantida)...
    if texto_numero is None: return None
    num_str = str(texto_numero).strip().replace('R$', '')
    num_str = num_str.replace(':', '.')
    if ',' in num_str and '.' in num_str:
         num_str = num_str.replace('.', '')
         num_str = num_str.replace(',', '.')
    elif ',' in num_str:
         num_str = num_str.replace(',', '.')
    elif '.' in num_str:
        parts = num_str.split('.')
        if len(parts) > 2: num_str = "".join(parts[:-1]) + "." + parts[-1]
    num_str_final = re.sub(r"[^-0-9.]", "", num_str)
    if num_str_final.count('.') > 1: parts = num_str_final.split('.'); num_str_final = "".join(parts[:-1]) + "." + parts[-1]
    if not num_str_final or num_str_final == '.' or num_str_final == '-': return None
    try: return float(num_str_final)
    except ValueError: return None

def extrair_texto_primeira_pagina(pdf_path):
    # ...(Função mantida)...
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages: return ""
            return pdf.pages[0].extract_text(x_tolerance=3, y_tolerance=3, layout=False) or ""
    except Exception as e: print(f"Erro pg1 {os.path.basename(pdf_path)}: {e}"); return ""

def extrair_tabelas_pdf(pdf_path):
     # ...(Função mantida)...
    tabelas_completas = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, pagina in enumerate(pdf.pages):
                tabelas_pagina = pagina.extract_tables(table_settings={ "vertical_strategy": "lines", "horizontal_strategy": "lines", "snap_tolerance": 5 })
                if not tabelas_pagina:
                     tabelas_pagina = pagina.extract_tables(table_settings={ "vertical_strategy": "text", "horizontal_strategy": "text", "snap_tolerance": 5, "join_tolerance": 3 })
                if tabelas_pagina:
                    for tabela in tabelas_pagina:
                         if isinstance(tabela, list) and all(isinstance(row, list) for row in tabela if row is not None):
                              tabelas_completas.append({'pagina': i + 1, 'tabela': tabela})
    except Exception as e: print(f"Erro tabelas {os.path.basename(pdf_path)}: {e}"); return []
    return tabelas_completas

def encontrar_valor_com_regex(texto, padrao_regex, grupo=1, flags=re.IGNORECASE | re.MULTILINE):
    # ...(Função mantida)...
    if texto is None: return None
    match = re.search(padrao_regex, texto, flags)
    if match:
        try: return match.group(grupo).strip()
        except IndexError:
             if grupo != 0:
                 try: return match.group(0).strip()
                 except IndexError: return None
             else: return None
        except AttributeError: return None
    return None

def encontrar_bloco_texto(texto, label_inicio, label_fim=None):
    # ...(Função mantida)...
    if texto is None: return None
    try:
        inicio = texto.find(label_inicio)
        if inicio == -1: return None
        inicio += len(label_inicio)
        fim = len(texto)
        if label_fim:
            fim_temp = texto.find(label_fim, inicio)
            if fim_temp != -1: fim = fim_temp
        bloco = texto[inicio:fim].strip()
        return bloco if bloco else None
    except Exception as e: print(f"Erro bloco '{label_inicio}': {e}"); return None

# --- Função Principal de Extração por NF-e (MANTIDA COMO ANTES) ---
def extrair_dados_nf(pdf_path):
    """
    Extrai dados do cabeçalho (priorizando VTotal no topo) e itens
    (tabelas multi-página com parse regex de célula mesclada).
    """
    # A lógica interna desta função permanece a mesma da versão anterior (VFinal)
    # print(f"\nProcessando arquivo: {os.path.basename(pdf_path)}") # Movido para processar_pasta_pdfs
    texto_pagina1 = extrair_texto_primeira_pagina(pdf_path)
    if not texto_pagina1:
        print(f"  AVISO: Falha ao extrair texto da página 1 de {os.path.basename(pdf_path)}.")

    # Extração do Cabeçalho
    # print("  > Extraindo dados do cabeçalho...") # Removido para diminuir verbosidade
    nome_emitente = None; data_emissao = None; numero_nf = None; nome_destinatario = None
    valor_total_nota = None
    try: # Destinatário
        dest_topo = encontrar_valor_com_regex(texto_pagina1, r'Valor\s+Total:.*?Destinatário:\s*(.*)', 1)
        if dest_topo: nome_destinatario = dest_topo
        else:
            bloco_dest = encontrar_bloco_texto(texto_pagina1, "Destinatário:", "Nº.:")
            if bloco_dest: nome_destinatario = bloco_dest.split('\n')[0].strip()
            if not nome_destinatario:
                bloco_dest_rem = encontrar_bloco_texto(texto_pagina1,"DESTINATARIO/REMETENTE", "CÁLCULO DO IMPOSTO")
                if bloco_dest_rem:
                     dest_match = encontrar_valor_com_regex(bloco_dest_rem, r'RAZÃO SOCIAL\s*\n([^\n]+)')
                     if dest_match: nome_destinatario = dest_match
    except Exception: pass
    try: # Data Emissão
        data_topo = encontrar_valor_com_regex(texto_pagina1, r'Emissão:\s*(\d{2}/\d{2}/\d{4})')
        if data_topo: data_emissao = data_topo
        else:
            data_emissao = encontrar_valor_com_regex(texto_pagina1, r'DATA DE EMISSÃO\s*(\d{2}/\d{2}/\d{4})')
            if not data_emissao: data_emissao = encontrar_valor_com_regex(texto_pagina1[:500], r'(\d{2}/\d{2}/\d{4})')
            if not data_emissao:
                 bloco_protocolo = encontrar_bloco_texto(texto_pagina1, "PROTOCOLO DE AUTORIZAÇÃO")
                 if bloco_protocolo: data_emissao = encontrar_valor_com_regex(bloco_protocolo, r'(\d{2}/\d{2}/\d{4})')
    except Exception: pass
    try: # Numero NF
        num_match = encontrar_valor_com_regex(texto_pagina1, r'Nº\.[:\s]*([\d\.]+)')
        if num_match: numero_nf = num_match.replace('.', '')
    except Exception: pass
    try: # Emitente
        bloco_emit = encontrar_bloco_texto(texto_pagina1, "Recebemos de", "os produtos")
        if bloco_emit: nome_emitente = bloco_emit.split('\n')[0].strip()
        if not nome_emitente:
            bloco_emit_alt = encontrar_bloco_texto(texto_pagina1, "IDENTIFICAÇÃO DO EMITENTE", "DANFE")
            if bloco_emit_alt: nome_emitente = bloco_emit_alt.split('\n')[0].strip()
    except Exception: pass
    try: # Valor Total NF (Prioriza Topo)
        vt_text_topo = None; vt_text_fallback = None
        pattern_vtotal_topo = r"Valor\s+Total:\s*R?\$?\s*([\d.,:]+)"
        vt_text_topo = encontrar_valor_com_regex(texto_pagina1, pattern_vtotal_topo)
        if vt_text_topo: valor_total_nota = limpar_numero(vt_text_topo)
        if valor_total_nota is None: # Fallback
            pattern_vtotal_flex = r"VALOR\s+TOTAL\s+DA\s+NOTA\s*(?:\n\s*R?\$?\s*([\d.,:]+)|R?\$?\s*([\d.,:]+))"
            match_flex = re.search(pattern_vtotal_flex, texto_pagina1, re.IGNORECASE)
            if match_flex: vt_text_fallback = match_flex.group(1) or match_flex.group(2)
            if not vt_text_fallback: vt_text_fallback = encontrar_valor_com_regex(texto_pagina1, r'VALOR TOTAL DA NOTA\s*R?\$\s*([\d.,:]+)')
            valor_total_nota = limpar_numero(vt_text_fallback)
    except Exception as e: print(f"  Erro Cabeçalho (VTotal): {e}")
    # Garantir valores padrão
    nome_emitente = nome_emitente or "Emitente não encontrado"
    data_emissao = data_emissao or "Data não encontrada"
    numero_nf = numero_nf or "NF não encontrada"
    destinatario = nome_destinatario or "Destinatário não encontrado"
    # print(f"  > Cabeçalho -> Emit: '{nome_emitente}', Data: '{data_emissao}', NF: '{numero_nf}', Dest: '{destinatario}', VTotalNota: {valor_total_nota}") # Debug

    # Extração de Itens (Mantida como na v5/v6 - com Regex na célula mesclada)
    # print("\n  > Extraindo e processando tabelas de itens...") # Debug
    dados_itens_finais = []
    tabelas_info = extrair_tabelas_pdf(pdf_path)
    if not tabelas_info: print(f"  AVISO: Nenhuma tabela encontrada em {os.path.basename(pdf_path)}.")
    else:
        regex_valores_item = re.compile(r"([A-Z]{1,3})\s*\(?([\d.,:]+)\)?\s*\(?([\d.,:]+)\)?\s*\(?([\d.,:]+)\)?")
        for info_tabela in tabelas_info:
            pagina_num = info_tabela['pagina']; tabela = info_tabela['tabela']
            if not tabela or not isinstance(tabela, list) or not tabela[0] or not isinstance(tabela[0], list): continue
            header_cells = tabela[0]
            header = [str(col).lower().replace('\n', ' ').strip() if col else "" for col in header_cells]
            has_desc_keyword = any('descri' in h or 'produto' in h for h in header)
            if has_desc_keyword and len(header) > 4:
                # print(f"    * Tabela Pág {pagina_num} identificada.") # Debug
                idx_desc = -1; idx_un_merged = -1; idx_cfop = -1
                # Mapeia colunas
                for i, h_text in enumerate(header):
                     if 'descri' in h_text: idx_desc = i; break
                if idx_desc == -1:
                     for i, h_text in enumerate(header):
                          if 'produto' in h_text and 'código' not in h_text: idx_desc = i; break
                for i, h_text in enumerate(header):
                     if 'cfop' in h_text: idx_cfop = i; break
                     if 'un' == h_text.strip() and idx_un_merged == -1: idx_un_merged = i
                if idx_cfop != -1 and idx_cfop + 1 < len(header): idx_un_merged = idx_cfop + 1
                if idx_un_merged == -1: idx_un_merged = 5 # Fallback
                if idx_desc == -1 or idx_un_merged == -1: print(f"      AVISO: Não mapeou Desc/Coluna Mesclada (Pág {pagina_num})."); continue
                # Processa Linhas
                for num_linha, linha in enumerate(tabela[1:], start=1):
                    if len(linha) > max(idx_desc, idx_un_merged) and any(c is not None and str(c).strip() != '' for c in linha):
                        def get_cell_value(row, index):
                            if index != -1 and index < len(row) and row[index] is not None: return str(row[index]).replace('\n', ' ').strip()
                            return None
                        item_desc = get_cell_value(linha, idx_desc)
                        merged_cell_text = get_cell_value(linha, idx_un_merged)
                        quantidade = None; valor_unitario = None; valor_total_item = None
                        if merged_cell_text:
                            match = regex_valores_item.search(merged_cell_text)
                            if match:
                                quantidade = limpar_numero(match.group(2))
                                valor_unitario = limpar_numero(match.group(3))
                                valor_total_item = limpar_numero(match.group(4))
                        if item_desc and quantidade is not None and valor_unitario is not None:
                            dados_itens_finais.append({
                                'Nome Emitente': nome_emitente, 'Data Emissão': data_emissao, 'Número NF': numero_nf,
                                'Destinatário': destinatario, 'Valor Total NF': valor_total_nota,
                                'Item Descrição': item_desc, 'Quantidade': quantidade,
                                'Valor Unitário': valor_unitario, 'Valor Total Item': valor_total_item
                            })
    # Fallback Final
    if not dados_itens_finais:
         # print(f"  AVISO FINAL: Nenhum item válido extraído de {os.path.basename(pdf_path)}.") # Debug
         dados_itens_finais.append({
            'Nome Emitente': nome_emitente, 'Data Emissão': data_emissao, 'Número NF': numero_nf,
            'Destinatário': destinatario, 'Valor Total NF': valor_total_nota,
            'Item Descrição': 'NENHUM ITEM EXTRAÍDO', # Simplificado
            'Quantidade': None, 'Valor Unitário': None, 'Valor Total Item': None
        })
    # print(f"  > Extração finalizada para {os.path.basename(pdf_path)}.") # Debug
    return dados_itens_finais


# --- Processamento Principal (MODIFICADO para busca recursiva) ---
def processar_pasta_pdfs(pasta_raiz, arquivo_saida): # Renomeado parâmetro
    """
    Encontra TODOS os PDFs recursivamente a partir da pasta raiz,
    extrai dados e salva em um único arquivo Excel na pasta raiz.
    """
    if not pasta_raiz or not os.path.isdir(pasta_raiz):
        print(f"ERRO: A pasta raiz '{pasta_raiz}' não foi encontrada ou é inválida.")
        return

    # --- >>> BUSCA RECURSIVA POR PDFs <<< ---
    print(f"\nBuscando arquivos PDF recursivamente em: '{pasta_raiz}'")
    padrao_busca = os.path.join(pasta_raiz, '**', '*.pdf')
    # recursive=True é essencial para o '**' funcionar
    arquivos_pdf = glob.glob(padrao_busca, recursive=True)
    # --- Fim da Busca Recursiva ---

    if not arquivos_pdf:
        print(f"Nenhum arquivo PDF encontrado em '{pasta_raiz}' ou subpastas.")
        return

    print(f"\nEncontrados {len(arquivos_pdf)} arquivos PDF para processar.")
    todos_os_dados = []

    # Processa cada PDF encontrado
    for pdf_path in arquivos_pdf:
        # Imprime o caminho relativo para melhor feedback
        print(f"\n--- Processando: {os.path.relpath(pdf_path, pasta_raiz)} ---")
        dados_da_nf_atual = extrair_dados_nf(pdf_path) # Chama a função de extração
        if dados_da_nf_atual:
            todos_os_dados.extend(dados_da_nf_atual)

    # --- Criação e Salvamento do DataFrame (Mantido) ---
    if not todos_os_dados:
        print("\nNenhum dado foi extraído de nenhum dos arquivos PDF processados.")
        return

    print(f"\nTotal de {len(todos_os_dados)} linhas de dados extraídas (incluindo fallbacks).")
    df = pd.DataFrame(todos_os_dados)
    # Filtra linhas de fallback ANTES de salvar
    df_filtrado = df[~df['Item Descrição'].str.contains("NENHUM ITEM EXTRAÍDO", na=False)]
    if df_filtrado.empty:
        print("AVISO: Apenas linhas de fallback foram geradas ou nenhum item encontrado. Nenhum arquivo Excel será salvo.")
        return # Não salva o arquivo se só tiver fallback

    df_final = df_filtrado
    print(f"Salvando {len(df_final)} linhas válidas...")
    # Ordena as colunas
    colunas_ordenadas = [
        'Nome Emitente', 'Data Emissão', 'Número NF', 'Destinatário',
        'Item Descrição', 'Quantidade', 'Valor Unitário', 'Valor Total Item',
        'Valor Total NF'
    ]
    colunas_presentes = [col for col in colunas_ordenadas if col in df_final.columns]
    df_final = df_final[colunas_presentes]

    try:
        # Salva o arquivo na PASTA RAIZ
        caminho_completo_saida = os.path.join(pasta_raiz, arquivo_saida)
        df_final.to_excel(caminho_completo_saida, index=False, engine='openpyxl')
        print(f"\nDados extraídos com sucesso e salvos em: '{caminho_completo_saida}'")
    except PermissionError: print(f"\nERRO DE PERMISSÃO: Não foi possível salvar '{caminho_completo_saida}'.")
    except Exception as e: print(f"\nErro GERAL ao salvar o arquivo Excel '{arquivo_saida}': {e}")
    #...(Fallback CSV mantido)...
    finally:
         try: # Tenta salvar CSV se Excel falhou
              if 'caminho_completo_saida' in locals() and not os.path.exists(caminho_completo_saida):
                   arquivo_saida_csv = arquivo_saida.replace('.xlsx', '.csv')
                   caminho_completo_saida_csv = os.path.join(pasta_raiz, arquivo_saida_csv)
                   df_final.to_csv(caminho_completo_saida_csv, index=False, encoding='utf-8-sig')
                   print(f"Como alternativa (Excel falhou), dados salvos em CSV: '{caminho_completo_saida_csv}'")
         except Exception as e_csv: print(f"Erro ao salvar também como CSV: {e_csv}")


# --- Execução do Script ---
if __name__ == "__main__":
    # >>> USA PASTA_RAIZ <<<
    if not PASTA_RAIZ:
         print("####################################################################")
         print("### ATENÇÃO: Defina a variável PASTA_RAIZ no início do script! ###")
         print("####################################################################")
    else:
         processar_pasta_pdfs(PASTA_RAIZ, NOME_ARQUIVO_SAIDA) # Passa a pasta raiz
    # input("\nPressione Enter para sair...")