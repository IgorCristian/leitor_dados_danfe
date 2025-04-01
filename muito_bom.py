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
NOME_ARQUIVO_SAIDA = 'dados_notas_fiscais_extraidos' # Novo nome

# --- Suprimir Avisos (Opcional) ---
# warnings.filterwarnings("ignore", message="CropBox missing")

# --- Funções Auxiliares ---

# Função limpar_numero (mantida a versão robusta anterior)
def limpar_numero(texto_numero):
    if texto_numero is None: return None
    num_str = str(texto_numero).strip().replace('R$', '')
    num_str = num_str.replace(':', '.') # Trata : como .
    # Remove pontos de milhar ANTES de tratar a vírgula decimal
    if ',' in num_str and '.' in num_str:
         # Se tem ambos, assume ponto é milhar e vírgula é decimal
         num_str = num_str.replace('.', '')
         num_str = num_str.replace(',', '.')
    elif ',' in num_str:
         # Se só tem vírgula, assume como decimal
         num_str = num_str.replace(',', '.')
    # Se só tem ponto, o último pode ser decimal
    elif '.' in num_str:
        parts = num_str.split('.')
        if len(parts) > 2: # Mais de um ponto -> primeiros são milhar
            num_str = "".join(parts[:-1]) + "." + parts[-1]
        # Se só um ponto, já está ok (ex: "1000." ou "0.60")

    # Remove caracteres não numéricos finais, exceto o ponto decimal
    num_str_final = re.sub(r"[^-0-9.]", "", num_str)
    # Garante apenas um ponto decimal
    if num_str_final.count('.') > 1:
         parts = num_str_final.split('.')
         num_str_final = "".join(parts[:-1]) + "." + parts[-1]

    if not num_str_final or num_str_final == '.' or num_str_final == '-': return None
    try:
        return float(num_str_final)
    except ValueError:
        # print(f"Debug limpar_numero: Erro final ao converter '{num_str_final}' de '{texto_numero}'")
        return None


# Outras funções auxiliares (extrair_texto_primeira_pagina, extrair_tabelas_pdf, encontrar_valor_com_regex, encontrar_bloco_texto)
# MANTENHA-AS EXATAMENTE COMO NA VERSÃO ANTERIOR
def extrair_texto_primeira_pagina(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages: return ""
            return pdf.pages[0].extract_text(x_tolerance=2, y_tolerance=2, layout=False) or ""
    except Exception as e: print(f"Erro pg1 {os.path.basename(pdf_path)}: {e}"); return ""

def extrair_tabelas_pdf(pdf_path):
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
    except Exception as e: print(f"Erro tabelas {os.path.basename(pdf_path)}: {e}"); return [] # Retorna lista vazia em erro
    return tabelas_completas

def encontrar_valor_com_regex(texto, padrao_regex, grupo=1, flags=re.IGNORECASE | re.MULTILINE):
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


# --- Função Principal de Extração por NF-e (COM PARSE MANUAL DA CÉLULA UN/QTD/VUNIT/VTOT) ---
def extrair_dados_nf(pdf_path):
    """
    Extrai dados, lendo cabeçalho da pág 1 e usando Regex para
    extrair Qtd/VUnit/VTotal de célula mesclada na tabela de itens.
    """
    print(f"\nProcessando arquivo: {os.path.basename(pdf_path)}")
    texto_pagina1 = extrair_texto_primeira_pagina(pdf_path)
    # ... (Extração do Cabeçalho mantida como na v4 - com fallbacks) ...
    print("  > Extraindo dados do cabeçalho...")
    nome_emitente = None; data_emissao = None; numero_nf = None; nome_destinatario = None; valor_total_nota = None
    try: # Destinatário
        bloco_dest = encontrar_bloco_texto(texto_pagina1, "Destinatário:", "Nº.:")
        if bloco_dest: nome_destinatario = bloco_dest.split('\n')[0].strip()
        if not nome_destinatario:
            bloco_dest_rem = encontrar_bloco_texto(texto_pagina1,"DESTINATARIO/REMETENTE", "CÁLCULO DO IMPOSTO")
            if bloco_dest_rem:
                 dest_match = encontrar_valor_com_regex(bloco_dest_rem, r'RAZÃO SOCIAL\s*\n([^\n]+)')
                 if dest_match: nome_destinatario = dest_match
    except Exception: pass
    try: # Data Emissão
        data_emissao = encontrar_valor_com_regex(texto_pagina1, r'(?:Emissão:|DATA DE EMISSÃO)\s*(\d{2}/\d{2}/\d{4})')
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
    try: # Valor Total NF
        vt_text = encontrar_valor_com_regex(texto_pagina1, r'VALOR TOTAL DA NOTA\s*R?\$\s*([\d.,:]+)')
        valor_total_nota = limpar_numero(vt_text)
    except Exception: pass
    # Valores Padrão
    nome_emitente = nome_emitente or "Emitente não encontrado"
    data_emissao = data_emissao or "Data não encontrada"
    numero_nf = numero_nf or "NF não encontrada"
    destinatario = nome_destinatario or "Destinatário não encontrado"
    print(f"  > Cabeçalho Extraído -> Emit: '{nome_emitente}', Data: '{data_emissao}', NF: '{numero_nf}', Dest: '{destinatario}', TotalNF: {valor_total_nota}")


    # --- Extração de Itens (MODIFICADA) ---
    print("\n  > Extraindo e processando tabelas de itens (com parse de célula mesclada)...")
    dados_itens_finais = []
    tabelas_info = extrair_tabelas_pdf(pdf_path)

    if not tabelas_info: print("  AVISO: Nenhuma tabela encontrada no PDF.")
    else:
        # --- >>> Regex para extrair Qtd, VUnit, VTotal da célula mesclada <<< ---
        # Captura: Grupo 1=UN, Grupo 2=Qtd, Grupo 3=VUnit, Grupo 4=VTotal
        # Permite espaços e parênteses opcionais entre os números
        # Aceita .,: como separadores nos números
        regex_valores_item = re.compile(r"([A-Z]{1,3})\s*\(?([\d.,:]+)\)?\s*\(?([\d.,:]+)\)?\s*\(?([\d.,:]+)\)?")

        for info_tabela in tabelas_info:
            pagina_num = info_tabela['pagina']
            tabela = info_tabela['tabela']
            if not tabela or not isinstance(tabela, list) or not tabela[0] or not isinstance(tabela[0], list): continue

            header_cells = tabela[0]
            header = [str(col).lower().replace('\n', ' ').strip() if col else "" for col in header_cells]
            # print(f"\n  > Verificando Tabela da Página {pagina_num} (Cabeçalho: {header})") # Debug

            # Identifica se parece tabela de itens (precisa ter descrição)
            has_desc_keyword = any('descri' in h or 'produto' in h for h in header)
            if has_desc_keyword and len(header) > 4: # Um pouco mais flexível na identificação
                print(f"    * Tabela da Página {pagina_num} IDENTIFICADA como potencial tabela de itens.")

                # --- Mapeamento de Colunas (Só precisamos da Descrição e da célula mesclada) ---
                idx_desc = -1
                idx_un_merged = -1 # Índice da coluna que começa com 'UN'

                # Mapeia Descrição (prioriza 'descri')
                for i, h_text in enumerate(header):
                     if 'descri' in h_text: idx_desc = i; break
                if idx_desc == -1:
                     for i, h_text in enumerate(header):
                          if 'produto' in h_text and 'código' not in h_text and 'codigo' not in h_text: idx_desc = i; break

                # Mapeia a coluna que contém a Unidade (que deve ser a mesclada)
                # Assume que é a coluna logo após 'CFOP' ou que contém 'UN' no header
                idx_cfop = -1
                for i, h_text in enumerate(header):
                     if 'cfop' in h_text: idx_cfop = i; break
                     if 'un' == h_text.strip() and idx_un_merged == -1: idx_un_merged = i # Tenta achar 'un' exato

                # Se achou CFOP e a coluna seguinte existe, usa ela como preferencial para 'UN'
                if idx_cfop != -1 and idx_cfop + 1 < len(header):
                     idx_un_merged = idx_cfop + 1
                # Se não achou 'un' exato nem CFOP+1, pode indicar erro de layout/parse
                if idx_un_merged == -1:
                     print(f"      AVISO: Não foi possível encontrar o índice da coluna 'UN' (mesclada?) nesta tabela (Pág {pagina_num}). Tentando índice 5 como fallback.")
                     idx_un_merged = 5 # Fallback para o índice 5 baseado no seu exemplo

                print(f"      Índices Mapeados -> Desc:{idx_desc}, Coluna UN/Qtd/Vlr(Mesclada):{idx_un_merged}")

                # Requer pelo menos Descrição e o índice da coluna mesclada
                if idx_desc == -1 or idx_un_merged == -1:
                     print(f"      AVISO: Não mapeou Descrição ou Coluna Mesclada (Pág {pagina_num}). Pulando processamento.")
                     continue

                # --- Processa as Linhas usando Regex na célula mesclada ---
                print(f"      Processando {len(tabela)-1} linhas de dados...")
                linhas_processadas_nesta_tabela = 0
                for num_linha, linha in enumerate(tabela[1:], start=1):
                    # Validação básica da linha
                    if len(linha) > max(idx_desc, idx_un_merged) and any(c is not None and str(c).strip() != '' for c in linha):

                        def get_cell_value(row, index): # Função auxiliar
                            if index != -1 and index < len(row) and row[index] is not None:
                                return str(row[index]).replace('\n', ' ').strip()
                            return None

                        item_desc = get_cell_value(linha, idx_desc)
                        merged_cell_text = get_cell_value(linha, idx_un_merged)

                        quantidade = None
                        valor_unitario = None
                        valor_total_item = None
                        qtd_str, vunit_str, vtotal_str = None, None, None # Strings originais

                        if merged_cell_text:
                            # Aplica o Regex para extrair os valores da célula mesclada
                            match = regex_valores_item.search(merged_cell_text)
                            if match:
                                # Pega os grupos capturados (ignora o grupo 1 = UN)
                                qtd_str = match.group(2)
                                vunit_str = match.group(3)
                                vtotal_str = match.group(4) # Pega o valor total do item aqui

                                # Limpa os valores extraídos pelo Regex
                                quantidade = limpar_numero(qtd_str)
                                valor_unitario = limpar_numero(vunit_str)
                                valor_total_item = limpar_numero(vtotal_str) # Usa o VTotal extraído
                                # print(f"        Linha {num_linha}: Regex Match! Qtd='{qtd_str}'->{quantidade}, VUnit='{vunit_str}'->{valor_unitario}, VTot='{vtotal_str}'->{valor_total_item}") # Debug
                            else:
                                # Regex não casou, talvez o formato seja inesperado
                                print(f"        AVISO Linha {num_linha}: Regex não encontrou padrão Qtd/VUnit/VTotal em '{merged_cell_text}'")
                                # Tenta pegar pelo menos Qtd/VUnit da célula seguinte se houver? (Mais complexo, evitar por ora)
                        else:
                            print(f"        AVISO Linha {num_linha}: Célula mesclada (índice {idx_un_merged}) está vazia.")


                        # Adiciona se tiver Descrição, Quantidade E V.Unit válidos (após regex e limpeza)
                        if item_desc and quantidade is not None and valor_unitario is not None:
                            dados_itens_finais.append({
                                'Nome Emitente': nome_emitente, 'Data Emissão': data_emissao, 'Número NF': numero_nf,
                                'Destinatário': destinatario, 'Valor Total NF': valor_total_nota,
                                'Item Descrição': item_desc, 'Quantidade': quantidade,
                                'Valor Unitário': valor_unitario, 'Valor Total Item': valor_total_item # Pode ser None se regex/limpeza falhar
                            })
                            linhas_processadas_nesta_tabela += 1
                        # else: # Debug porque não adicionou
                        #      reason = []
                        #      if not item_desc: reason.append("Descrição inválida")
                        #      if quantidade is None: reason.append(f"Qtd inválida (str='{qtd_str}')")
                        #      if valor_unitario is None: reason.append(f"V.Unit inválido (str='{vunit_str}')")
                        #      print(f"        - Item linha {num_linha} IGNORADO. Motivo(s): {', '.join(reason)}")

                print(f"      {linhas_processadas_nesta_tabela} linhas processadas com sucesso nesta tabela.")

    # --- Fallback Final ---
    if not dados_itens_finais:
         print("\n  AVISO FINAL: Nenhum item válido foi extraído de nenhuma tabela.")
         dados_itens_finais.append({
            'Nome Emitente': nome_emitente, 'Data Emissão': data_emissao, 'Número NF': numero_nf,
            'Destinatário': destinatario, 'Valor Total NF': valor_total_nota,
            'Item Descrição': 'NENHUM ITEM EXTRAÍDO (Método Tabela/Regex)',
            'Quantidade': None, 'Valor Unitário': None, 'Valor Total Item': None
        })

    print(f"\n  > Extração finalizada para {os.path.basename(pdf_path)}. Total de linhas de itens adicionadas: {len([d for d in dados_itens_finais if d.get('Item Descrição') != 'NENHUM ITEM EXTRAÍDO (Método Tabela/Regex)']) }")
    return dados_itens_finais


# --- Processamento Principal (Itera, chama extrair_dados_nf, salva Excel) ---
# (Mantido como na versão anterior)
def processar_pasta_pdfs(pasta_pdfs, arquivo_saida):
    """Processa PDFs e salva em Excel."""
    # ...(código mantido)...
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
    df_filtrado = df[~df['Item Descrição'].str.contains("NENHUM ITEM EXTRAÍDO", na=False)] # Filtra fallback
    if df_filtrado.empty:
        print("AVISO: Apenas linhas de fallback foram geradas ou nenhum item encontrado. Nenhum arquivo Excel será salvo.")
        return
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
    except PermissionError: print(f"\nERRO DE PERMISSÃO: Não foi possível salvar '{caminho_completo_saida}'. Verifique se o arquivo está aberto.")
    except Exception as e: print(f"\nErro GERAL ao salvar o arquivo Excel '{arquivo_saida}': {e}")
    #...(Fallback CSV mantido)...


# --- Execução do Script ---
if __name__ == "__main__":
    if not PASTA_PDFS: print("ATENÇÃO: Defina PASTA_PDFS!")
    else: processar_pasta_pdfs(PASTA_PDFS, NOME_ARQUIVO_SAIDA)
    # input("\nPressione Enter para sair...")