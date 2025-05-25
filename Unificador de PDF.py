import os
import fitz  # PyMuPDF
import time  # Para medir o tempo de execução
import shutil # Para copiar arquivos (Opcional, explicado abaixo)

# --- Configuração ---

# 1. Pasta Onde Estão Seus PDFs Individuais:
#    !! IMPORTANTE: Coloque aqui o caminho CORRETO da pasta !!
pasta_origem_pdfs = r'C:\Users\Igor\Desktop\Nova pasta\2024' 

# 2. Pasta Para Onde Copiar os PDFs (Opcional):
#    Se você quiser trabalhar com CÓPIAS para não modificar os originais,
#    defina uma pasta de destino. Se não quiser copiar, deixe como None.
#    !! CERTIFIQUE-SE QUE ESTA PASTA EXISTA OU SERÁ CRIADA !!
pasta_destino_copias = 'pasta_copias_nfs' # Exemplo: 'copias_para_unir' ou None
#    Se definido, os PDFs da pasta_origem_pdfs serão copiados para cá antes de unir.
#    Se for None, o script usará os arquivos diretamente da pasta_origem_pdfs.

# 3. Nome do Arquivo PDF Unificado que Será Criado:
nome_arquivo_final_unificado = 'TODAS_AS_NOTAS_UNIFICADAS.pdf' 

# --------------------

def copiar_pdfs(origem, destino):
    """Copia arquivos .pdf da pasta de origem para a de destino."""
    print(f"Iniciando cópia de PDFs de '{origem}' para '{destino}'...")
    if not os.path.exists(destino):
        print(f"Criando pasta de destino: '{destino}'")
        os.makedirs(destino)
        
    arquivos_copiados = 0
    erros_copia = []
    
    for item in os.listdir(origem):
        caminho_origem_item = os.path.join(origem, item)
        if os.path.isfile(caminho_origem_item) and item.lower().endswith('.pdf'):
            caminho_destino_item = os.path.join(destino, item)
            try:
                shutil.copy2(caminho_origem_item, caminho_destino_item) # copy2 preserva metadados
                arquivos_copiados += 1
            except Exception as e:
                print(f"  Erro ao copiar '{item}': {e}")
                erros_copia.append(item)
                
    print(f"Cópia concluída. {arquivos_copiados} arquivos PDF copiados.")
    if erros_copia:
        print("Arquivos com erro durante a cópia:", erros_copia)
    print("-" * 30)
    return arquivos_copiados > 0 # Retorna True se algo foi copiado

def unir_pdfs(pasta_contendo_pdfs, arquivo_saida):
    """Une todos os PDFs de uma pasta em um único arquivo de saída."""
    arquivos_pdf_para_unir = []
    print(f"Buscando arquivos PDF para unir em: '{os.path.abspath(pasta_contendo_pdfs)}'...")

    # Busca PDFs na pasta especificada (não busca em subpastas aqui, ajuste se necessário)
    for filename in os.listdir(pasta_contendo_pdfs):
        if filename.lower().endswith('.pdf'):
            caminho_completo = os.path.join(pasta_contendo_pdfs, filename)
            arquivos_pdf_para_unir.append(caminho_completo)

    if not arquivos_pdf_para_unir:
        print("Nenhum arquivo PDF encontrado na pasta para unificação.")
        return

    # Ordena os arquivos (alfabeticamente pelo nome) para ordem consistente
    arquivos_pdf_para_unir.sort()
    num_arquivos = len(arquivos_pdf_para_unir)
    print(f"Encontrados {num_arquivos} arquivos PDF para unir.")

    # Cria o documento PDF de saída (vazio inicialmente)
    pdf_final = fitz.open() 
    erros_uniao = []
    start_time = time.time()

    print("Iniciando processo de unificação...")
    for i, caminho_pdf in enumerate(arquivos_pdf_para_unir):
        nome_base = os.path.basename(caminho_pdf)
        # Imprime progresso a cada 100 arquivos ou no último
        if (i + 1) % 100 == 0 or (i + 1) == num_arquivos: 
             print(f"  Processando arquivo {i+1}/{num_arquivos}: {nome_base}")
        try:
            # Abre o PDF de entrada atual
            with fitz.open(caminho_pdf) as pdf_entrada:
                # Insere (anexa) todas as páginas do PDF de entrada no PDF final
                pdf_final.insert_pdf(pdf_entrada)
        except Exception as e:
            print(f"  ERRO ao processar o arquivo '{nome_base}': {e}. Pulando este arquivo.")
            erros_uniao.append(nome_base)
            continue # Pula para o próximo arquivo

    if len(pdf_final) == 0:
         print("Nenhuma página foi adicionada ao arquivo final (talvez todos os PDFs deram erro?). Saindo.")
         pdf_final.close()
         return

    print("\nFinalizando e salvando o arquivo unificado...")
    try:
        # Salva o PDF final unificado
        pdf_final.save(arquivo_saida, garbage=4, deflate=True) # garbage=4 otimiza, deflate=True comprime
        end_time = time.time()
        print("-" * 30)
        print(f"SUCESSO! Arquivo unificado salvo como: '{arquivo_saida}'")
        print(f"Total de páginas no arquivo final: {len(pdf_final)}")
        print(f"Tempo total de unificação: {end_time - start_time:.2f} segundos")
        if erros_uniao:
            print("\nAtenção: Os seguintes arquivos não puderam ser processados e foram pulados:")
            for erro_f in erros_uniao:
                print(f"  - {erro_f}")
        print("-" * 30)

    except Exception as e:
        print(f"ERRO ao salvar o arquivo final '{arquivo_saida}': {e}")
    finally:
        # Fecha o documento final
        pdf_final.close()

# --- Execução Principal ---
if __name__ == "__main__":
    
    pasta_a_unir = pasta_origem_pdfs # Por padrão, usa a pasta original

    # Bloco Opcional de Cópia:
    if pasta_destino_copias:
        # Verifica se a pasta de origem e destino são a mesma para evitar problemas
        if os.path.abspath(pasta_origem_pdfs) == os.path.abspath(pasta_destino_copias):
            print("Erro: A pasta de origem e a pasta de destino das cópias não podem ser a mesma.")
        else:
            if copiar_pdfs(pasta_origem_pdfs, pasta_destino_copias):
                # Se a cópia foi bem sucedida, muda a pasta a ser unida para a pasta de cópias
                pasta_a_unir = pasta_destino_copias
            else:
                print("A cópia dos arquivos falhou ou nenhum arquivo foi copiado. Verifique os erros.")
                # Decide se quer parar ou continuar com a original (aqui vamos parar)
                exit() # Sai do script se a cópia era desejada e falhou
    else:
        print("INFO: Nenhuma pasta de destino para cópias foi definida. Os arquivos originais serão usados para a unificação.")

    # Unifica os PDFs da pasta definida (original ou de cópias)
    unir_pdfs(pasta_a_unir, nome_arquivo_final_unificado)

    print("\nProcesso de unificação de PDFs concluído.")