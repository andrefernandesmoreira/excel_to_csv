import streamlit as st
import io
import zipfile
import csv
from openpyxl import load_workbook

# --- Configurações da página (DEVE SER A PRIMEIRA COISA APÓS OS IMPORTS) ---
st.set_page_config(
    page_title="Conversor de Excel para CSV - Andre", # Título que aparece na aba do navegador
    page_icon="📄", # Um emoji como ícone da aba
    layout="centered", # 'centered' ou 'wide' - define a largura da página
    initial_sidebar_state="collapsed" # 'auto', 'expanded', 'collapsed' - se houver sidebar
)

st.title("🗂️ Conversor de Excel para CSV") # Título principal da sua aplicação na página
st.markdown("""
    Esta ferramenta permite converter facilmente seus arquivos **Excel (.xlsx)**
    para o formato **CSV (UTF-8)**,
    mantendo a fidelidade de exportação do Excel.
""")
st.info("💡 **Dica:** Você pode selecionar e carregar múltiplos arquivos de uma vez para conversão em lote!")

# Widget para upload de arquivos
uploaded_files = st.file_uploader(
    "Selecione um ou mais arquivos Excel para converter:",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

if uploaded_files:
    # Cria um placeholder vazio para mensagens de progresso/status
    processing_message = st.empty()
    processing_message.info("⏳ Processando seus arquivos... Por favor, aguarde.")

    # ZIP em memória para armazenar os CSVs convertidos
    zip_buffer = io.BytesIO()
    processed_count = 0
    errors_occurred = False

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
        for file in uploaded_files:
            try:
                # Atualiza a mensagem de progresso para o arquivo atual
                processing_message.info(f"⚙️ Convertendo: **{file.name}**...")
                
                # Carregar planilha inteira com openpyxl
                wb = load_workbook(file, data_only=True)
                ws = wb.active  # pega a aba ativa

                # --- Lógica para determinar a largura máxima da planilha ---
                # Isso garante que todas as linhas terão o mesmo número de "colunas" (delimitadores)
                max_column_in_sheet = 0
                for row_idx, row_data in enumerate(ws.iter_rows()):
                    current_row_max_col = 0
                    for cell_idx, cell in enumerate(row_data):
                        # Considera uma célula "não vazia" se tiver valor ou não for vazia após strip
                        if cell.value is not None and str(cell.value).strip() != '':
                            current_row_max_col = max(current_row_max_col, cell.column)
                    
                    max_column_in_sheet = max(max_column_in_sheet, current_row_max_col)
                
                # Garante que pelo menos uma coluna seja escrita, mesmo em planilhas vazias
                if max_column_in_sheet == 0:
                    max_column_in_sheet = 1

                # Buffer em memória para o CSV
                csv_buffer = io.StringIO()
                
                # Configura o escritor CSV com ponto e vírgula como delimitador
                writer = csv.writer(
                    csv_buffer,
                    delimiter=";",  # Usa ponto e vírgula como delimitador
                    lineterminator="\r\n", # Padrão de quebra de linha do Windows (mais compatível)
                    quoting=csv.QUOTE_MINIMAL # Cita apenas strings que contêm delimitadores, aspas, etc.
                )

                # Escreve linha a linha no buffer CSV
                for row_data in ws.iter_rows(values_only=True):
                    row_list = list(row_data)

                    # Garante que a linha tenha o número máximo de colunas detectado na planilha.
                    # Preenche com strings vazias para simular os delimitadores extras do Excel em colunas vazias.
                    if len(row_list) < max_column_in_sheet:
                        row_list.extend([''] * (max_column_in_sheet - len(row_list)))
                    
                    # Converte cada valor para string (None vira "") e limita a linha à largura máxima
                    final_row = ["" if v is None else str(v) for v in row_list[:max_column_in_sheet]]
                    
                    writer.writerow(final_row)

                # Converte o conteúdo do buffer CSV para bytes UTF-8 com BOM (Byte Order Mark)
                csv_bytes = csv_buffer.getvalue().encode("utf-8-sig")

                # Define o nome do arquivo CSV dentro do ZIP
                base_name = file.name.rsplit(".", 1)[0] # Remove a extensão original
                csv_name = f"{base_name}.csv" # Adiciona a extensão .csv

                # Grava o CSV convertido no arquivo ZIP em memória
                zipf.writestr(csv_name, csv_bytes)
                processed_count += 1 # Incrementa o contador de arquivos processados

            except Exception as e:
                # Exibe mensagem de erro para o arquivo específico que falhou
                st.error(f"❌ Erro ao processar o arquivo **'{file.name}'**: {e}")
                errors_occurred = True # Sinaliza que houve algum erro

    # Limpa a mensagem de processamento após a conclusão de todos os arquivos
    processing_message.empty()

    # Mensagens de sucesso/aviso final
    if processed_count > 0:
        if not errors_occurred:
            st.success(f"🎉 Conversão concluída com sucesso! **{processed_count} arquivo(s)** convertido(s).")
        else:
            st.warning(f"⚠️ Conversão concluída, mas com **erros** em alguns arquivos. **{processed_count} arquivo(s)** convertido(s) com sucesso.")

        # Botão de download do arquivo ZIP
        st.download_button(
            label="📥 Baixar todos os CSVs (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="arquivos_convertidos.zip", # Nome do arquivo ZIP final para download
            mime="application/zip",
            help="Clique para baixar um arquivo ZIP contendo todos os seus CSVs convertidos."
        )
    else:
        # Mensagem se nenhum arquivo válido foi processado (ex: upload cancelado, ou todos deram erro)
        st.info("Nenhum arquivo Excel válido foi processado para conversão.")

st.markdown("---") # Linha divisória
st.markdown("👨‍💻 Desenvolvido por [Andre Fernandes Moreira].") # Link para o seu GitHub