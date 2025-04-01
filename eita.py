# -*- coding: utf-8 -*- # Define a codificação do arquivo como UTF-8

# --- Importação das Bibliotecas Necessárias ---
import pdfplumber
import pandas as pd
import os
import re
import glob
from datetime import datetime
import warnings # Para suprimir avisos

# --- Configurações Iniciais ---
PASTA_PDFS = r'C:\Users\Igor\Desktop\Projeto Planilhas\NOTAS DE 2024\01205929169\00000Isento\2024\Fevereiro'
# >>> Nome do arquivo final <<<
NOME_ARQUIVO_SAIDA = 'dados_notas_fiscais_extraidos_vfinal.xlsx'

# --- Suprimir Avisos (Opcional) ---
# warnings.filterwarnings("ignore", message="CropBox missing")

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
            # Aumenta tolerâncias para tentar juntar texto próximo na linha do topo
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

# --- Função Principal de Extração por NF-e (VERSÃO FINAL) ---
def extrair_dados_nf(pdf_path):
    """
    Extrai dados do cabeçalho (priorizando VTotal no topo) e itens
    (tabelas multi-página com parse regex de célula mesclada).
    """
    print(f"\nProcessando arquivo: {os.path.basename(pdf_path)}")
    texto_pagina1 = extrair_texto_primeira_pagina(pdf_path)
    if not texto_pagina1:
        print(f"  AVISO: Falha ao extrair texto da página 1 de {os.path.basename(pdf_path)}.")

    # --- Extração do Cabeçalho ---
    print("  > Extraindo dados do cabeçalho...")
    nome_emitente = None; data_emissao = None; numero_nf = None; nome_destinatario = None
    valor_total_nota = None # <- Variável que conterá o valor final desejado

    # (Extração de Emitente, Data, NF, Destinatário mantida como antes)
    try: # Destinatário
        # Tenta primeiro na linha do topo
        dest_topo = encontrar_valor_com_regex(texto_pagina1, r'Valor\s+Total:.*?Destinatário:\s*(.*)', 1)
        if dest_topo: nome_destinatario = dest_topo
        else: # Fallbacks
            bloco_dest = encontrar_bloco_texto(texto_pagina1, "Destinatário:", "Nº.:")
            if bloco_dest: nome_destinatario = bloco_dest.split('\n')[0].strip()
            if not nome_destinatario:
                bloco_dest_rem = encontrar_bloco_texto(texto_pagina1,"DESTINATARIO/REMETENTE", "CÁLCULO DO IMPOSTO")
                if bloco_dest_rem:
                     dest_match = encontrar_valor_com_regex(bloco_dest_rem, r'RAZÃO SOCIAL\s*\n([^\n]+)')
                     if dest_match: nome_destinatario = dest_match
    except Exception: pass
    try: # Data Emissão
        # Tenta primeiro na linha do topo
        data_topo = encontrar_valor_com_regex(texto_pagina1, r'Emissão:\s*(\d{2}/\d{2}/\d{4})')
        if data_topo: data_emissao = data_topo
        else: # Fallbacks
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

    # --- >>> LÓGICA FINAL PARA VALOR TOTAL (Prioriza Topo) <<< ---
    try:
        valor_total_nota = None # Reset
        vt_text_topo = None
        vt_text_fallback = None

        # 1. PRIORIDADE: Busca "Valor Total:" na linha do topo
        pattern_vtotal_topo = r"Valor\s+Total:\s*R?\$?\s*([\d.,:]+)"
        vt_text_topo = encontrar_valor_com_regex(texto_pagina1, pattern_vtotal_topo)
        print(f"  DEBUG VTOTAL_NF (Topo): Texto encontrado: '{vt_text_topo}'") # Debug
        if vt_text_topo:
             valor_total_nota = limpar_numero(vt_text_topo)

        # 2. FALLBACK: Se não achou no topo OU falhou na limpeza, busca "VALOR TOTAL DA NOTA"
        if valor_total_nota is None:
            print("  DEBUG VTOTAL_NF (Topo falhou): Tentando fallback 'VALOR TOTAL DA NOTA'...")
            pattern_vtotal_flex = r"VALOR\s+TOTAL\s+DA\s+NOTA\s*(?:\n\s*R?\$?\s*([\d.,:]+)|R?\$?\s*([\d.,:]+))"
            match_flex = re.search(pattern_vtotal_flex, texto_pagina1, re.IGNORECASE)
            if match_flex: vt_text_fallback = match_flex.group(1) or match_flex.group(2)
            if not vt_text_fallback: # Fallback do fallback
                 vt_text_fallback = encontrar_valor_com_regex(texto_pagina1, r'VALOR TOTAL DA NOTA\s*R?\$\s*([\d.,:]+)')
            print(f"  DEBUG VTOTAL_NF (Fallback): Texto encontrado: '{vt_text_fallback}'") # Debug
            valor_total_nota = limpar_numero(vt_text_fallback)

    except Exception as e:
        print(f"  Erro Cabeçalho (VTotal): {e}")
    # --- Fim da Lógica VTotal ---

    # --- REMOVIDA A EXTRAÇÃO DA BASE DE CÁLCULO ICMS ---

    # Garantir valores padrão
    nome_emitente = nome_emitente or "Emitente não encontrado"
    data_emissao = data_emissao or "Data não encontrada"
    numero_nf = numero_nf or "NF não encontrada"
    destinatario = nome_destinatario or "Destinatário não encontrado"
    # valor_total_nota pode ser None se ambas as tentativas falharem

    print(f"  > Cabeçalho Extraído -> Emit: '{nome_emitente}', Data: '{data_emissao}', NF: '{numero_nf}', Dest: '{destinatario}', VTotalNota: {valor_total_nota}")


    # --- Extração de Itens (Mantida como na v5 - com Regex na célula mesclada) ---
    print("\n  > Extraindo e processando tabelas de itens...")
    dados_itens_finais = []
    tabelas_info = extrair_tabelas_pdf(pdf_path)

    if not tabelas_info: print("  AVISO: Nenhuma tabela encontrada no PDF.")
    else:
        # Regex para extrair Qtd, VUnit, VTotal da célula mesclada
        regex_valores_item = re.compile(r"([A-Z]{1,3})\s*\(?([\d.,:]+)\)?\s*\(?([\d.,:]+)\)?\s*\(?([\d.,:]+)\)?")

        for info_tabela in tabelas_info:
            # ...(Lógica para identificar tabela de itens e mapear idx_desc/idx_un_merged MANTIDA)...
            pagina_num = info_tabela['pagina']; tabela = info_tabela['tabela']
            if not tabela or not isinstance(tabela, list) or not tabela[0] or not isinstance(tabela[0], list): continue
            header_cells = tabela[0]
            header = [str(col).lower().replace('\n', ' ').strip() if col else "" for col in header_cells]
            has_desc_keyword = any('descri' in h or 'produto' in h for h in header)
            if has_desc_keyword and len(header) > 4:
                # print(f"    * Tabela da Página {pagina_num} IDENTIFICADA como potencial tabela de itens.") # Debug
                idx_desc = -1; idx_un_merged = -1; idx_cfop = -1
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
                # print(f"      Índices Mapeados -> Desc:{idx_desc}, Coluna UN/Qtd/Vlr(Mesclada):{idx_un_merged}") # Debug
                if idx_desc == -1 or idx_un_merged == -1: print(f"      AVISO: Não mapeou Desc/Coluna Mesclada (Pág {pagina_num})."); continue

                # Processa Linhas
                # print(f"      Processando {len(tabela)-1} linhas de dados...") # Debug
                linhas_processadas_nesta_tabela = 0
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
                                'Destinatário': destinatario,
                                'Valor Total NF': valor_total_nota, # <- Valor correto (do topo ou fallback)
                                # REMOVIDA: 'Base Calculo ICMS': base_calculo_icms,
                                'Item Descrição': item_desc, 'Quantidade': quantidade,
                                'Valor Unitário': valor_unitario, 'Valor Total Item': valor_total_item
                            })
                            linhas_processadas_nesta_tabela += 1
                # print(f"      {linhas_processadas_nesta_tabela} linhas processadas com sucesso nesta tabela.") # Debug

    # --- Fallback Final ---
    if not dados_itens_finais:
         print("\n  AVISO FINAL: Nenhum item válido foi extraído.")
         dados_itens_finais.append({
            'Nome Emitente': nome_emitente, 'Data Emissão': data_emissao, 'Número NF': numero_nf,
            'Destinatário': destinatario, 'Valor Total NF': valor_total_nota,
            # REMOVIDA: 'Base Calculo ICMS': base_calculo_icms,
            'Item Descrição': 'NENHUM ITEM EXTRAÍDO (Método Tabela/Regex)',
            'Quantidade': None, 'Valor Unitário': None, 'Valor Total Item': None
        })

    print(f"\n  > Extração finalizada para {os.path.basename(pdf_path)}. Total de linhas de itens adicionadas: {len([d for d in dados_itens_finais if d.get('Item Descrição') != 'NENHUM ITEM EXTRAÍDO (Método Tabela/Regex)']) }")
    return dados_itens_finais


# --- Processamento Principal (Itera, chama extrair_dados_nf, salva Excel) ---
def processar_pasta_pdfs(pasta_pdfs, arquivo_saida):
    """Processa PDFs e salva em Excel."""
    # ...(Validação de pasta mantida)...
    if not pasta_pdfs or not os.path.isdir(pasta_pdfs):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        if not os.path.isdir(pasta_pdfs): print(f"ERRO: Pasta '{pasta_pdfs}' não encontrada."); return
        else: print(f"AVISO: Usando pasta do script: '{script_dir}'"); pasta_pdfs = script_dir
    padrao_busca = os.path.join(pasta_pdfs, '*.pdf')
    arquivos_pdf = glob.glob(padrao_busca)
    if not arquivos_pdf: print(f"Nenhum PDF encontrado em '{pasta_pdfs}'."); return
    print(f"\nEncontrados {len(arquivos_pdf)} arquivos PDF para processar em '{pasta_pdfs}'.")
    todos_os_dados = []
    for pdf_path in arquivos_pdf:
        dados_da_nf_atual = extrair_dados_nf(pdf_path)
        if dados_da_nf_atual: todos_os_dados.extend(dados_da_nf_atual)
    if not todos_os_dados: print("\nNenhum dado foi extraído."); return
    print(f"\nTotal de {len(todos_os_dados)} linhas de dados extraídas (incluindo fallbacks).")
    df = pd.DataFrame(todos_os_dados)
    df_filtrado = df[~df['Item Descrição'].str.contains("NENHUM ITEM EXTRAÍDO", na=False)]
    if df_filtrado.empty:
        print("AVISO: Apenas linhas de fallback foram geradas ou nenhum item encontrado. Nenhum arquivo Excel será salvo.")
        return
    df_final = df_filtrado
    print(f"Salvando {len(df_final)} linhas válidas...")
    # --- >>> COLUNAS FINAL (REMOVIDA Base Calculo ICMS) <<< ---
    colunas_ordenadas = [
        'Nome Emitente', 'Data Emissão', 'Número NF', 'Destinatário',
        'Item Descrição', 'Quantidade', 'Valor Unitário', 'Valor Total Item',
        'Valor Total NF'     # <- Contém o valor do topo (ou fallback)
    ]
    colunas_presentes = [col for col in colunas_ordenadas if col in df_final.columns]
    df_final = df_final[colunas_presentes]
    try:
        caminho_completo_saida = os.path.join(pasta_pdfs, arquivo_saida)
        df_final.to_excel(caminho_completo_saida, index=False, engine='openpyxl')
        print(f"\nDados extraídos com sucesso e salvos em: '{caminho_completo_saida}'")
    except PermissionError: print(f"\nERRO DE PERMISSÃO: Não foi possível salvar '{caminho_completo_saida}'.")
    except Exception as e: print(f"\nErro GERAL ao salvar o arquivo Excel '{arquivo_saida}': {e}")
    #...(Fallback CSV mantido)...
    finally:
         try:
              if 'caminho_completo_saida' in locals() and not os.path.exists(caminho_completo_saida):
                   arquivo_saida_csv = arquivo_saida.replace('.xlsx', '.csv')
                   caminho_completo_saida_csv = os.path.join(pasta_pdfs, arquivo_saida_csv)
                   df_final.to_csv(caminho_completo_saida_csv, index=False, encoding='utf-8-sig')
                   print(f"Como alternativa (Excel falhou), dados salvos em CSV: '{caminho_completo_saida_csv}'")
         except Exception as e_csv: print(f"Erro ao salvar também como CSV: {e_csv}")


# --- Execução do Script ---
if __name__ == "__main__":
    if not PASTA_PDFS: print("ATENÇÃO: Defina PASTA_PDFS!")
    else: processar_pasta_pdfs(PASTA_PDFS, NOME_ARQUIVO_SAIDA)
    # input("\nPressione Enter para sair...")