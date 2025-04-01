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
NOME_ARQUIVO_SAIDA = 'dados_notas_fiscais_extraidos_multi_pagina_v3.xlsx' # Novo nome

# --- Suprimir Avisos (Opcional) ---
# warnings.filterwarnings("ignore", message="CropBox missing")

# --- Funções Auxiliares ---

def limpar_numero(texto_numero):
    """Limpia string numérica robusta (ponto milhar, vírgula/ponto/dois-pontos decimal)."""
    if texto_numero is None: return None
    num_str = str(texto_numero).strip().replace('R$', '')
    # Trata ':' como '.'
    num_str = num_str.replace(':', '.')
    # Define possível separador decimal (última , ou .)
    last_sep = -1
    if ',' in num_str: last_sep = max(last_sep, num_str.rfind(','))
    if '.' in num_str: last_sep = max(last_sep, num_str.rfind('.'))

    numero_limpo = ""
    if last_sep != -1:
        potential_decimal = num_str[last_sep+1:]
        digits_after_sep = re.sub(r'\D', '', potential_decimal)
        # Considera decimal se for o último sep e tiver 1 a 3 digitos depois
        is_decimal_sep = (len(digits_after_sep) <= 3)

        if is_decimal_sep:
            integer_part_str = num_str[:last_sep]
            decimal_part_str = num_str[last_sep+1:]
            integer_part = re.sub(r'\D', '', integer_part_str)
            decimal_part = re.sub(r'\D', '', decimal_part_str)
            # Evita criar ". " se decimal for vazio
            numero_limpo = f"{integer_part}{'.' + decimal_part if decimal_part else ''}"
        else:
             numero_limpo = re.sub(r'\D', '', num_str)
    else:
        numero_limpo = re.sub(r'\D', '', num_str)

    if not numero_limpo or numero_limpo == '.': return None
    try:
        return float(numero_limpo)
    except ValueError:
        return None

def extrair_texto_primeira_pagina(pdf_path):
    """Extrai texto da primeira página (para cabeçalho)."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                 print(f"Erro: PDF {os.path.basename(pdf_path)} não contém páginas.")
                 return ""
            return pdf.pages[0].extract_text(x_tolerance=2, y_tolerance=2, layout=False) or ""
    except Exception as e:
        print(f"Erro ao ler a primeira página do PDF {os.path.basename(pdf_path)}: {e}")
        return ""

def extrair_tabelas_pdf(pdf_path):
    """Extrai TODAS as tabelas de TODAS as páginas de um PDF."""
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
    except Exception as e:
        print(f"  ERRO GERAL ao extrair tabelas do PDF {os.path.basename(pdf_path)}: {e}")
    print(tabelas_completas)    
    return tabelas_completas

def encontrar_valor_com_regex(texto, padrao_regex, grupo=1, flags=re.IGNORECASE | re.MULTILINE):
    """Busca um padrão (Regex) no texto."""
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
    """Encontra e retorna um bloco de texto entre dois labels."""
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
    except Exception as e:
        print(f"Erro ao buscar bloco entre '{label_inicio}' e '{label_fim}': {e}")
        return None


# --- Função Principal de Extração por NF-e (CORRIGIDA) ---
def extrair_dados_nf(pdf_path):
    """
    Orquestra a extração de dados, lendo cabeçalho da página 1 e
    processando tabelas de itens de TODAS as páginas.
    """
    print(f"\nProcessando arquivo: {os.path.basename(pdf_path)}")
    texto_pagina1 = extrair_texto_primeira_pagina(pdf_path)
    if not texto_pagina1:
        print(f"  AVISO: Falha ao extrair texto da primeira página de {os.path.basename(pdf_path)}. Cabeçalho pode faltar.")

    # --- Extração do Cabeçalho (LÓGICA DO SEU SCRIPT + FALLBACKS) ---
    print("  > Extraindo dados do cabeçalho...")
    nome_emitente = None; data_emissao = None; numero_nf = None; nome_destinatario = None; valor_total_nota = None

    # Destinatário
    try:
        bloco_dest = encontrar_bloco_texto(texto_pagina1, "Destinatário:", "Nº.:") # Ajuste delimitadores
        if bloco_dest: nome_destinatario = bloco_dest.split('\n')[0].strip()
        # Fallback se o primeiro falhar (procura por RAZÃO SOCIAL dentro de DESTINATARIO/REMETENTE)
        if not nome_destinatario:
            bloco_dest_rem = encontrar_bloco_texto(texto_pagina1,"DESTINATARIO/REMETENTE", "CÁLCULO DO IMPOSTO")
            if bloco_dest_rem:
                 dest_match = encontrar_valor_com_regex(bloco_dest_rem, r'RAZÃO SOCIAL\s*\n([^\n]+)')
                 if dest_match: nome_destinatario = dest_match
    except Exception as e: print(f"  Erro Cabeçalho (Dest): {e}")

    # Data de Emissão (com fallbacks)
    try:
        # 1ª Tentativa: Label explícito
        data_emissao = encontrar_valor_com_regex(texto_pagina1, r'(?:Emissão:|DATA DE EMISSÃO)\s*(\d{2}/\d{2}/\d{4})')
        # 2ª Tentativa: Qualquer data no início do texto
        if not data_emissao:
            data_emissao = encontrar_valor_com_regex(texto_pagina1[:500], r'(\d{2}/\d{2}/\d{4})') # Primeiros 500 chars
        # 3ª Tentativa: Perto do protocolo de autorização
        if not data_emissao:
             bloco_protocolo = encontrar_bloco_texto(texto_pagina1, "PROTOCOLO DE AUTORIZAÇÃO")
             if bloco_protocolo:
                  data_emissao = encontrar_valor_com_regex(bloco_protocolo, r'(\d{2}/\d{2}/\d{4})')
    except Exception as e: print(f"  Erro Cabeçalho (Data): {e}")

    # Número NF
    try:
        num_match = encontrar_valor_com_regex(texto_pagina1, r'Nº\.[:\s]*([\d\.]+)')
        if num_match: numero_nf = num_match.replace('.', '')
    except Exception as e: print(f"  Erro Cabeçalho (NF): {e}")

    # Emitente
    try:
        bloco_emit = encontrar_bloco_texto(texto_pagina1, "Recebemos de", "os produtos")
        if bloco_emit: nome_emitente = bloco_emit.split('\n')[0].strip()
        if not nome_emitente:
            bloco_emit_alt = encontrar_bloco_texto(texto_pagina1, "IDENTIFICAÇÃO DO EMITENTE", "DANFE")
            if bloco_emit_alt: nome_emitente = bloco_emit_alt.split('\n')[0].strip()
    except Exception as e: print(f"  Erro Cabeçalho (Emit): {e}")

    # Valor Total NF
    try:
        vt_text = encontrar_valor_com_regex(texto_pagina1, r'VALOR TOTAL DA NOTA\s*R?\$\s*([\d.,:]+)')
        valor_total_nota = limpar_numero(vt_text)
    except Exception as e: print(f"  Erro Cabeçalho (VTotal): {e}")

    # Garantir valores padrão
    nome_emitente = nome_emitente or "Emitente não encontrado"
    data_emissao = data_emissao or "Data não encontrada"
    numero_nf = numero_nf or "NF não encontrada"
    destinatario = nome_destinatario or "Destinatário não encontrado"

    print(f"  > Cabeçalho Extraído -> Emit: '{nome_emitente}', Data: '{data_emissao}', NF: '{numero_nf}', Dest: '{destinatario}', TotalNF: {valor_total_nota}")

    # --- Extração de Itens (Usando extrair_tabelas_pdf de TODAS as páginas) ---
    print("\n  > Extraindo e processando tabelas de itens de todas as páginas...")
    dados_itens_finais = []
    tabelas_info = extrair_tabelas_pdf(pdf_path)

    if not tabelas_info:
        print("  AVISO: Nenhuma tabela encontrada no PDF.")
    else:
        for info_tabela in tabelas_info:
            pagina_num = info_tabela['pagina']
            tabela = info_tabela['tabela']

            if not tabela or not isinstance(tabela, list) or not tabela[0] or not isinstance(tabela[0], list): continue

            header_cells = tabela[0]
            header = [str(col).lower().replace('\n', ' ').strip() if col else "" for col in header_cells]
            # print(f"\n  > Verificando Tabela da Página {pagina_num} (Cabeçalho: {header})") # Debug

            has_desc = any('descri' in h or 'produto' in h for h in header)
            has_qtd = any('qtd' in h or 'qde' in h or 'quantl' in h for h in header)
            has_vunit = any('unit' in h or 'unitário' in h or 'unitario' in h for h in header)

            # Considera itens se tiver Descrição E (Qtd OU V.Unit) E colunas suficientes
            if has_desc and (has_qtd or has_vunit) and len(header) > 4:
                print(f"    * Tabela da Página {pagina_num} IDENTIFICADA como potencial tabela de itens.")

                # --- Mapeamento de Colunas CORRIGIDO para idx_desc ---
                idx_desc, idx_qtd, idx_vunit, idx_vtotal = -1, -1, -1, -1
                # 1. Tenta achar 'descri' primeiro (mais específico)
                for i, h_text in enumerate(header):
                     if 'descri' in h_text:
                          idx_desc = i
                          break
                # 2. Se não achou, tenta 'produto' mas evita a coluna 'código produto'
                if idx_desc == -1:
                     for i, h_text in enumerate(header):
                          if 'produto' in h_text and 'código' not in h_text and 'codigo' not in h_text:
                              idx_desc = i
                              break
                # Mapeia os outros índices
                for i, h_text in enumerate(header):
                    # Não sobrescreve idx_desc se já achou
                    if ('qtd' in h_text or 'qde' in h_text or 'quantl' in h_text) and idx_qtd == -1: idx_qtd = i
                    elif ('unit' in h_text or 'unitário' in h_text or 'unitario' in h_text) and idx_vunit == -1: idx_vunit = i
                    elif ('v. tot' in h_text or 'v tot' in h_text or 'valor tot' in h_text) and idx_vtotal == -1:
                         if idx_vunit == -1 or i != idx_vunit: idx_vtotal = i
                    elif ('total' in h_text) and idx_vtotal == -1:
                         if idx_vunit == -1 or i != idx_vunit: idx_vtotal = i

                print(f"      Índices -> Desc:{idx_desc}, Qtd:{idx_qtd}, VUnit:{idx_vunit}, VTotal:{idx_vtotal if idx_vtotal != -1 else 'N/A'}")

                # Requer pelo menos Descrição e (Qtd OU VUnit) para processar
                if idx_desc == -1 or (idx_qtd == -1 and idx_vunit == -1):
                     print(f"      AVISO: Não foi possível mapear Descrição e (Qtd ou V.Unit) nesta tabela da página {pagina_num}.")
                     continue # Pula para a próxima tabela

                # --- Processa as Linhas DESTA Tabela de Itens ---
                print(f"      Processando {len(tabela)-1} linhas de dados...")
                linhas_processadas_nesta_tabela = 0
                for num_linha, linha in enumerate(tabela[1:], start=1):
                    max_idx_needed = max(idx for idx in [idx_desc, idx_qtd, idx_vunit, idx_vtotal] if idx != -1)
                    # Validação da linha (comprimento e conteúdo)
                    if len(linha) > max_idx_needed and any(c is not None and str(c).strip() != '' for c in linha):

                        def get_cell_value(row, index):
                            if index != -1 and index < len(row) and row[index] is not None:
                                return str(row[index]).replace('\n', ' ').strip()
                            return None

                        item_desc = get_cell_value(linha, idx_desc)
                        item_qtd_str = get_cell_value(linha, idx_qtd)
                        item_vunit_str = get_cell_value(linha, idx_vunit)
                        item_vtotal_str = get_cell_value(linha, idx_vtotal)

                        # Limpa os números com a função MELHORADA
                        quantidade = limpar_numero(item_qtd_str)
                        valor_unitario = limpar_numero(item_vunit_str)
                        valor_total_item = limpar_numero(item_vtotal_str)

                        # Fallback: Calcula VTotal se não puder limpar o da tabela mas tiver Qtd e VUnit
                        if valor_total_item is None and quantidade is not None and valor_unitario is not None:
                             try: valor_total_item = round(quantidade * valor_unitario, 2)
                             except TypeError: pass

                        # Adiciona se tiver descrição E (Qtd OU V.Unit válidos)
                        if item_desc and (quantidade is not None or valor_unitario is not None):
                            dados_itens_finais.append({
                                'Nome Emitente': nome_emitente, 'Data Emissão': data_emissao, 'Número NF': numero_nf,
                                'Destinatário': destinatario, 'Valor Total NF': valor_total_nota,
                                'Item Descrição': item_desc, 'Quantidade': quantidade,
                                'Valor Unitário': valor_unitario, 'Valor Total Item': valor_total_item
                            })
                            linhas_processadas_nesta_tabela += 1
                        # else: # Debug porque uma linha foi ignorada
                        #      print(f"        - Linha {num_linha} IGNORADA. Desc='{item_desc}', Qtd={quantidade}, VUnit={valor_unitario}")
                    # else: # Debug porque a linha foi pulada
                         # print(f"      - Linha {num_linha} pulada (len={len(linha)} vs max_idx={max_idx_needed} or vazia/None)")
                print(f"      {linhas_processadas_nesta_tabela} linhas processadas com sucesso nesta tabela.")
            # else:
                 # print(f"  > Tabela da Página {pagina_num} NÃO parece ser de itens.") # Debug

    # --- Fallback Final ---
    if not dados_itens_finais:
         print("\n  AVISO FINAL: Nenhum item válido foi extraído de nenhuma tabela.")
         dados_itens_finais.append({
            'Nome Emitente': nome_emitente, 'Data Emissão': data_emissao, 'Número NF': numero_nf,
            'Destinatário': destinatario, 'Valor Total NF': valor_total_nota,
            'Item Descrição': 'NENHUM ITEM EXTRAÍDO (Método Tabela)',
            'Quantidade': None, 'Valor Unitário': None, 'Valor Total Item': None
        })

    print(f"\n  > Extração finalizada para {os.path.basename(pdf_path)}. Total de linhas de itens adicionadas: {len([d for d in dados_itens_finais if d.get('Item Descrição') != 'NENHUM ITEM EXTRAÍDO (Método Tabela)']) }")
    return dados_itens_finais


# --- Processamento Principal (Itera, chama extrair_dados_nf, salva Excel) ---
def processar_pasta_pdfs(pasta_pdfs, arquivo_saida):
    """Processa PDFs e salva em Excel."""
    # ...(código mantido como na resposta anterior)...
    if not pasta_pdfs or not os.path.isdir(pasta_pdfs):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        if not os.path.isdir(pasta_pdfs):
             print(f"ERRO: A pasta '{pasta_pdfs}' não foi encontrada.")
             if os.path.isdir(script_dir):
                  print(f"AVISO: Tentando usar a pasta do script como alternativa: '{script_dir}'")
                  pasta_pdfs = script_dir
             else:
                  print("ERRO: Pasta do script também não encontrada. Verifique o caminho.")
                  return
        else:
             print(f"AVISO: Nenhuma pasta especificada, usando a pasta do script: '{script_dir}'")
             pasta_pdfs = script_dir

    padrao_busca = os.path.join(pasta_pdfs, '*.pdf')
    arquivos_pdf = glob.glob(padrao_busca)
    if not arquivos_pdf: print(f"Nenhum arquivo PDF encontrado em '{pasta_pdfs}'."); return

    print(f"\nEncontrados {len(arquivos_pdf)} arquivos PDF para processar em '{pasta_pdfs}'.")
    todos_os_dados = []
    for pdf_path in arquivos_pdf:
        dados_da_nf_atual = extrair_dados_nf(pdf_path)
        if dados_da_nf_atual: todos_os_dados.extend(dados_da_nf_atual)

    if not todos_os_dados: print("\nNenhum dado foi extraído."); return

    print(f"\nTotal de {len(todos_os_dados)} linhas de dados extraídas (incluindo fallbacks).")
    df = pd.DataFrame(todos_os_dados)
    df_filtrado = df[df['Item Descrição'] != 'NENHUM ITEM EXTRAÍDO (Método Tabela)']
    if df_filtrado.empty:
         if df.empty: print("Nenhum dado gerado."); return
         print("AVISO: Apenas linhas de fallback foram geradas. Verifique os logs. Nenhum arquivo Excel será salvo.")
         return # Não salva se só tiver fallback
    else:
         df_final = df_filtrado

    print(f"Salvando {len(df_final)} linhas válidas...")
    colunas_ordenadas = [
        'Nome Emitente', 'Data Emissão', 'Número NF', 'Destinatário',
        'Item Descrição', 'Quantidade', 'Valor Unitário', 'Valor Total Item',
        'Valor Total NF'
    ]
    colunas_presentes = [col for col in colunas_ordenadas if col in df_final.columns]
    df_final = df_final[colunas_presentes]

    try:
        caminho_completo_saida = os.path.join(pasta_pdfs, arquivo_saida)
        df_final.to_excel(caminho_completo_saida, index=False, engine='openpyxl')
        print(f"\nDados extraídos com sucesso e salvos em: '{caminho_completo_saida}'")
    except PermissionError:
         print(f"\nERRO DE PERMISSÃO: Não foi possível salvar '{caminho_completo_saida}'. Verifique se o arquivo está aberto.")
    except Exception as e:
        print(f"\nErro GERAL ao salvar o arquivo Excel '{arquivo_saida}': {e}")
        try:
            # ... (Fallback CSV mantido) ...
            arquivo_saida_csv = arquivo_saida.replace('.xlsx', '.csv')
            caminho_completo_saida_csv = os.path.join(pasta_pdfs, arquivo_saida_csv)
            df_final.to_csv(caminho_completo_saida_csv, index=False, encoding='utf-8-sig')
            print(f"Como alternativa, os dados foram salvos em formato CSV: '{caminho_completo_saida_csv}'")
        except Exception as e_csv:
            print(f"Erro ao salvar também como CSV: {e_csv}")


# --- Execução do Script ---
if __name__ == "__main__":
    if not PASTA_PDFS: print("ATENÇÃO: Defina PASTA_PDFS!")
    else: processar_pasta_pdfs(PASTA_PDFS, NOME_ARQUIVO_SAIDA)
    # input("\nPressione Enter para sair...")