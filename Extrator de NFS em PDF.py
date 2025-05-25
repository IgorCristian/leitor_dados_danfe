# -*- coding: utf-8 -*- # Define a codificação do arquivo como UTF-8

# --- Importação das Bibliotecas Necessárias ---
import pdfplumber
import pandas as pd
import os
import re
import glob
from datetime import datetime

# --- Configurações Iniciais ---
PASTA_RAIZ = r'C:\Users\Igor\Desktop\Teste do extrator'
NOME_ARQUIVO_SAIDA = 'dados_extraidos.xlsx' # Novo nome


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
                             tabela_limpa = [[str(cell) if cell is not None else "" for cell in row] for row in tabela if row is not None and any(cell is not None and str(cell).strip() != '' for cell in row)]
                             if tabela_limpa: tabelas_completas.append({'pagina': i + 1, 'tabela': tabela_limpa})
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


# --- Função Principal de Extração por NF-e (CORRIGIDA) ---
def extrair_dados_nf(pdf_path, nome_arquivo):
    """
    Extrai dados, com detecção de layout de itens e tratamento de erro por linha.
    Versão com correções baseadas no debug.
    """
    print(f"--- Iniciando extração para: {nome_arquivo} ---")
    texto_pagina1 = extrair_texto_primeira_pagina(pdf_path)

    # --- Extração do Cabeçalho (Mantida) ---
    nome_emitente = None; data_emissao = None; numero_nf = None; nome_destinatario = None; valor_total_nota = None
    # ... (Lógica completa de extração do cabeçalho como antes) ...
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
        if valor_total_nota is None:
            pattern_vtotal_flex = r"VALOR\s+TOTAL\s+DA\s+NOTA\s*(?:\n\s*R?\$?\s*([\d.,:]+)|R?\$?\s*([\d.,:]+))"
            match_flex = re.search(pattern_vtotal_flex, texto_pagina1, re.IGNORECASE)
            if match_flex: vt_text_fallback = match_flex.group(1) or match_flex.group(2)
            if not vt_text_fallback: vt_text_fallback = encontrar_valor_com_regex(texto_pagina1, r'VALOR TOTAL DA NOTA\s*R?\$\s*([\d.,:]+)')
            valor_total_nota = limpar_numero(vt_text_fallback)
    except Exception as e: print(f"  Erro VTotal: {e}")
    # Garantir valores padrão
    nome_emitente = nome_emitente or "Emitente N/E"; data_emissao = data_emissao or "Data N/E"
    numero_nf = numero_nf or "NF N/E"; destinatario = nome_destinatario or "Destinatário N/E"
    print(f"  DEBUG Cabeçalho Final -> Emit: '{nome_emitente}', Data: '{data_emissao}', NF: '{numero_nf}', Dest: '{destinatario}', VTotalNota: {valor_total_nota}")

    # --- Extração de Itens (TENTATIVA SEQUENCIAL CORRIGIDA) ---
    print("\n  DEBUG: Extraindo Itens por Tentativa Sequencial...")
    dados_itens_finais = []
    tabelas_info = extrair_tabelas_pdf(pdf_path)
    if not tabelas_info: print(f"  DEBUG: Nenhuma tabela encontrada em {nome_arquivo}.")
    else:
        # Regex para Layout A (UN/Qtd/VUnit/VTotal mesclados)
        regex_unqvu = re.compile(r"([A-Z]{1,3})\s*\(?([\d.,:]+)\)?\s*\(?([\d.,:]+)\)?\s*\(?([\d.,:]+)\)?")
        # --- >>> Regex CORRIGIDO para Layout D (Qtd/VUnit/VTotal mesclados) <<< ---
        # Remove parênteses opcionais em volta de cada grupo numérico
        regex_qvu = re.compile(r"\(?([\d.,:]+)\)?\s+\(?([\d.,:]+)\)?\s+\(?([\d.,:]+)\)?") # Captura 3 números separados por espaços (com ou sem parênteses GERAIS)

        for info_tabela in tabelas_info:
            pagina_num = info_tabela['pagina']; tabela = info_tabela['tabela']
            if not tabela or not isinstance(tabela, list) or len(tabela) < 2 or not isinstance(tabela[0], list): continue

            header_cells = tabela[0]
            header = [str(col).lower().replace('\n', ' ').strip() if col else "" for col in header_cells]
            # print(f"\n  DEBUG: Verificando Tabela Pag {pagina_num} - Cabeçalho: {header}") # Debug Verboso

            # --- >>> Identificação de Tabela de Itens MAIS Específica <<< ---
            # Exige Descrição E (Quantidade OU V.Unit) para considerar
            has_desc_keyword = any('descri' in h or 'produto' in h for h in header)
            has_qtd_keyword = any('qtd' in h or 'qde' in h or 'quant' in h for h in header) # 'quant' adicionado
            has_vunit_keyword = any('unit' in h or 'unitário' in h or 'unitario' in h for h in header)

            if has_desc_keyword and (has_qtd_keyword or has_vunit_keyword) and len(header) > 4:
                print(f"    DEBUG: Tabela Pag {pagina_num} parece ser de itens.")

                # --- Mapeamento Inicial de Índices (Melhor Esforço) ---
                idx_desc = -1; idx_qtd = -1; idx_vunit = -1; idx_vtotal = -1; idx_un = -1; idx_merged_qvu = -1; idx_merged_unqvu = -1
                for i, h_text in enumerate(header):
                    if ('descri' in h_text or ('produto' in h_text and 'código' not in h_text)) and idx_desc == -1: idx_desc = i
                    # --- >>> Adicionado 'quanti' <<< ---
                    if ('qtd' in h_text or 'qde' in h_text or h_text.startswith('quant')) and idx_qtd == -1: idx_qtd = i
                    if ('unit' in h_text or 'unitário' in h_text or 'unitario' in h_text) and idx_vunit == -1: idx_vunit = i
                    if ('v. tot' in h_text or 'v tot' in h_text or 'valor tot' in h_text or ('total' == h_text and i > idx_vunit if idx_vunit!=-1 else -1)) and idx_vtotal == -1:
                         if i != idx_vunit : idx_vtotal = i
                    if 'un' == h_text.strip() and idx_un == -1: idx_un = i
                # Define possíveis índices mesclados (baseado nos mapeamentos acima)
                if idx_qtd != -1 and idx_qtd == idx_vunit: idx_merged_qvu = idx_qtd # Se Qtd e VUnit caem no mesmo índice
                if idx_un != -1: idx_merged_unqvu = idx_un # Onde está a coluna 'UN'
                else: idx_merged_unqvu = 5 # Fallback

                print(f"      DEBUG Índices Mapeados -> Desc:{idx_desc}, Qtd:{idx_qtd}, VUnit:{idx_vunit}, VTotal:{idx_vtotal}, UN:{idx_un}, MergedQVU:{idx_merged_qvu}, MergedUNQVU:{idx_merged_unqvu}")

                if idx_desc == -1: print(f"      AVISO: Não mapeou Descrição."); continue

                # --- Processa as Linhas com TENTATIVA SEQUENCIAL ---
                print(f"      DEBUG: Processando {len(tabela)-1} linhas...")
                for num_linha, linha_bruta in enumerate(tabela[1:], start=1):
                    # print(f"\n      DEBUG Linha {num_linha} Bruta: {linha_bruta}") # Debug verboso linha
                    quantidade = None; valor_unitario = None; valor_total_item = None
                    item_desc = None; metodo_sucesso = "Nenhum"

                    try:
                        # Validação básica e pegar descrição
                        if len(linha_bruta) <= idx_desc or not any(c is not None and str(c).strip() != '' for c in linha_bruta):
                            continue # Pula linha curta ou vazia
                        def get_cell_value(row, index):
                            if index != -1 and index < len(row) and row[index] is not None: return str(row[index]).replace('\n', ' ').strip()
                            return None
                        item_desc = get_cell_value(linha_bruta, idx_desc)
                        if not item_desc: continue # Pula se descrição for vazia

                        # --- TENTATIVA 1: LAYOUT SEPARADO ---
                        if idx_qtd != -1 and idx_vunit != -1 and idx_qtd != idx_vunit:
                            # print(f"        DEBUG Tentando Layout SEPARADO (Qtd:{idx_qtd}, VUnit:{idx_vunit})...")
                            qtd_sep_str = get_cell_value(linha_bruta, idx_qtd)
                            vunit_sep_str = get_cell_value(linha_bruta, idx_vunit)
                            qtd_sep = limpar_numero(qtd_sep_str)
                            vunit_sep = limpar_numero(vunit_sep_str)
                            if qtd_sep is not None and vunit_sep is not None:
                                quantidade = qtd_sep
                                valor_unitario = vunit_sep
                                vtotal_sep_str = get_cell_value(linha_bruta, idx_vtotal) # Usa VTotal mapeado se existir
                                valor_total_item = limpar_numero(vtotal_sep_str)
                                metodo_sucesso = "SEPARADO"

                        # --- TENTATIVA 2: LAYOUT REGEX QVU ---
                        # Verifica se o método anterior falhou E se o índice MergedQVU foi identificado
                        if metodo_sucesso == "Nenhum" and idx_merged_qvu != -1:
                            # print(f"        DEBUG Tentando Layout REGEX_QVU (Índice:{idx_merged_qvu})...")
                            merged_text_qvu = get_cell_value(linha_bruta, idx_merged_qvu)
                            if merged_text_qvu:
                                # print(f"          Texto QVU: '{merged_text_qvu}'")
                                match_qvu = regex_qvu.search(merged_text_qvu) # Usa o regex CORRIGIDO
                                if match_qvu:
                                    # print(f"          Regex QVU Casou: {match_qvu.groups()}") # Debug
                                    qtd_qvu = limpar_numero(match_qvu.group(1))
                                    vunit_qvu = limpar_numero(match_qvu.group(2))
                                    if qtd_qvu is not None and vunit_qvu is not None:
                                        quantidade = qtd_qvu
                                        valor_unitario = vunit_qvu
                                        valor_total_item = limpar_numero(match_qvu.group(3))
                                        metodo_sucesso = "REGEX_QVU"
                                # else: print(f"          Regex QVU não casou.") # Debug

                        # --- TENTATIVA 3: LAYOUT REGEX UNQVU ---
                        # Verifica se os métodos anteriores falharam E se o índice MergedUNQVU é válido
                        if metodo_sucesso == "Nenhum" and idx_merged_unqvu != -1:
                             # print(f"        DEBUG Tentando Layout REGEX_UNQVU (Índice:{idx_merged_unqvu})...")
                             merged_text_unqvu = get_cell_value(linha_bruta, idx_merged_unqvu)
                             if merged_text_unqvu:
                                 # print(f"          Texto UNQVU: '{merged_text_unqvu}'") # Debug
                                 match_unqvu = regex_unqvu.search(merged_text_unqvu)
                                 if match_unqvu:
                                     # print(f"          Regex UNQVU Casou: {match_unqvu.groups()}") # Debug
                                     qtd_unqvu = limpar_numero(match_unqvu.group(2))
                                     vunit_unqvu = limpar_numero(match_unqvu.group(3))
                                     if qtd_unqvu is not None and vunit_unqvu is not None:
                                         quantidade = qtd_unqvu
                                         valor_unitario = vunit_unqvu
                                         valor_total_item = limpar_numero(match_unqvu.group(4))
                                         metodo_sucesso = "REGEX_UNQVU"
                                 # else: print(f"          Regex UNQVU não casou.") # Debug

                        # --- Fallback cálculo VTotalItem ---
                        if metodo_sucesso != "Nenhum" and valor_total_item is None and quantidade is not None and valor_unitario is not None:
                             try: valor_total_item = round(quantidade * valor_unitario, 2)
                             except TypeError: pass

                        # --- Adiciona o item ---
                        if item_desc and quantidade is not None and valor_unitario is not None:
                            # print(f"        --->>> ADICIONANDO ITEM Linha {num_linha} (Método: {metodo_sucesso}) <<<---") # Debug
                            dados_itens_finais.append({
                                'Arquivo Origem': nome_arquivo, 'Nome Emitente': nome_emitente, 'Data Emissão': data_emissao,
                                'Número NF': numero_nf, 'Destinatário': destinatario, 'Valor Total NF': valor_total_nota,
                                'Item Descrição': item_desc, 'Quantidade': quantidade,
                                'Valor Unitário': valor_unitario, 'Valor Total Item': valor_total_item
                            })
                        # else: # Debug item inválido
                        #     print(f"        --->>> ITEM INVÁLIDO Linha {num_linha} (Método: {metodo_sucesso}) <<<---")

                    except Exception as e_linha:
                         print(f"      ERRO ao processar linha {num_linha} Pág {pagina_num}: {e_linha}")
                         continue

    # --- Fallback Final ---
    if not dados_itens_finais:
         # print(f"\n  AVISO FINAL: Nenhum item válido extraído de {nome_arquivo}.") # Debug
         dados_itens_finais.append({ # Adiciona fallback se a lista estiver vazia
            'Arquivo Origem': nome_arquivo,'Nome Emitente': nome_emitente, 'Data Emissão': data_emissao, 'Número NF': numero_nf,
            'Destinatário': destinatario, 'Valor Total NF': valor_total_nota,
            'Item Descrição': 'NENHUM ITEM EXTRAÍDO',
            'Quantidade': None, 'Valor Unitário': None, 'Valor Total Item': None
        })
    return dados_itens_finais


# --- Processamento Principal (Itera, chama extrair_dados_nf, salva Excel) ---
# (Mantido como na versão anterior)
def processar_pasta_pdfs(pasta_raiz, arquivo_saida):
    """Processa PDFs recursivamente e salva em Excel."""
    if not pasta_raiz or not os.path.isdir(pasta_raiz): print(f"ERRO: Pasta raiz '{pasta_raiz}' inválida."); return
    padrao_busca = os.path.join(pasta_raiz, '**', '*.pdf')
    arquivos_pdf = glob.glob(padrao_busca, recursive=True)
    if not arquivos_pdf: print(f"Nenhum PDF encontrado em '{pasta_raiz}' ou subpastas."); return
    print(f"\nEncontrados {len(arquivos_pdf)} arquivos PDF para processar.")
    todos_os_dados = []
    for pdf_path in arquivos_pdf:
        nome_arquivo_pdf = os.path.basename(pdf_path)
        print(f"\n--- Processando: {os.path.relpath(pdf_path, pasta_raiz)} ---")
        try:
            dados_da_nf_atual = extrair_dados_nf(pdf_path, nome_arquivo_pdf)
            if dados_da_nf_atual: todos_os_dados.extend(dados_da_nf_atual)
        except Exception as e_pdf:
             print(f"*** ERRO FATAL AO PROCESSAR O ARQUIVO {nome_arquivo_pdf}: {e_pdf} ***")
             todos_os_dados.append({ # Adiciona linha de erro
                'Arquivo Origem': nome_arquivo_pdf, 'Nome Emitente': 'ERRO', 'Data Emissão': 'ERRO', 'Número NF': 'ERRO',
                'Destinatário': 'ERRO', 'Valor Total NF': None, 'Item Descrição': f'ERRO PDF: {e_pdf}',
                'Quantidade': None, 'Valor Unitário': None, 'Valor Total Item': None
             })

    if not todos_os_dados: print("\nNenhum dado foi extraído."); return
    print(f"\nTotal de {len(todos_os_dados)} linhas de dados geradas (incluindo erros/fallbacks).")
    df = pd.DataFrame(todos_os_dados)
    df_filtrado = df[ (~df['Item Descrição'].str.contains("NENHUM ITEM EXTRAÍDO", na=False)) & (~df['Item Descrição'].str.startswith("ERRO PDF:", na=False)) ]
    if df_filtrado.empty: print("AVISO: Nenhum item válido extraído de nenhum arquivo. Nenhum arquivo Excel será salvo."); return
    df_final = df_filtrado
    print(f"Salvando {len(df_final)} linhas válidas...")
    colunas_ordenadas = ['Arquivo Origem', 'Nome Emitente', 'Data Emissão', 'Número NF', 'Destinatário','Item Descrição', 'Quantidade', 'Valor Unitário', 'Valor Total Item','Valor Total NF']
    colunas_presentes = [col for col in colunas_ordenadas if col in df_final.columns]
    df_final = df_final[colunas_presentes]
    try:
        caminho_completo_saida = os.path.join(pasta_raiz, arquivo_saida)
        df_final.to_excel(caminho_completo_saida, index=False, engine='openpyxl')
        print(f"\nDados extraídos com sucesso e salvos em: '{caminho_completo_saida}'")
    except PermissionError: print(f"\nERRO DE PERMISSÃO: Não foi possível salvar '{caminho_completo_saida}'.")
    except Exception as e: print(f"\nErro GERAL ao salvar o arquivo Excel '{arquivo_saida}': {e}")
    #...(Fallback CSV mantido)...
    finally:
         try:
              if 'caminho_completo_saida' in locals() and not os.path.exists(caminho_completo_saida):
                   arquivo_saida_csv = arquivo_saida.replace('.xlsx', '.csv')
                   caminho_completo_saida_csv = os.path.join(pasta_raiz, arquivo_saida_csv)
                   df_final.to_csv(caminho_completo_saida_csv, index=False, encoding='utf-8-sig')
                   print(f"Como alternativa (Excel falhou), dados salvos em CSV: '{caminho_completo_saida_csv}'")
         except Exception as e_csv: print(f"Erro ao salvar também como CSV: {e_csv}")


# --- Execução do Script ---
if __name__ == "__main__":
    if not PASTA_RAIZ: print("ATENÇÃO: Defina PASTA_RAIZ!")
    else: processar_pasta_pdfs(PASTA_RAIZ, NOME_ARQUIVO_SAIDA)
    # input("\nPressione Enter para sair...")