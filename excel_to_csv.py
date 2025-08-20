import streamlit as st
import io
import zipfile
import csv
from openpyxl import load_workbook

st.title("Conversor Excel → CSV")

uploaded_files = st.file_uploader(
    "Selecione um ou mais arquivos Excel",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

if uploaded_files:
    # ZIP em memória
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
        for file in uploaded_files:
            try:
                # Carregar planilha inteira com openpyxl
                wb = load_workbook(file, data_only=True)
                ws = wb.active  # pega a aba ativa

                # --- Lógica para determinar a largura máxima da planilha ---
                # Isso garante que todas as linhas terão o mesmo número de "colunas" (delimitadores)
                max_column_in_sheet = 0
                for row_idx, row_data in enumerate(ws.iter_rows()):
                    current_row_max_col = 0
                    for cell_idx, cell in enumerate(row_data):
                        if cell.value is not None and str(cell.value).strip() != '':
                            current_row_max_col = max(current_row_max_col, cell.column)
                    
                    max_column_in_sheet = max(max_column_in_sheet, current_row_max_col)
                
                if max_column_in_sheet == 0:
                    max_column_in_sheet = 1 # Garante que pelo menos uma coluna seja escrita

                # Buffer para CSV
                csv_buffer = io.StringIO()
                
                # *** CORREÇÃO AQUI: ALTERADO DELIMITADOR PARA PONTO E VÍRGULA ***
                writer = csv.writer(
                    csv_buffer,
                    delimiter=";",  # Agora usa ponto e vírgula como delimitador
                    lineterminator="\r\n", # Padrão do Windows
                    quoting=csv.QUOTE_MINIMAL # Mantido, pois é o mais próximo do Excel para não citar desnecessariamente
                )

                # Escreve linha a linha, garantindo o número correto de colunas e preenchimento
                for row_data in ws.iter_rows(values_only=True):
                    row_list = list(row_data)

                    # Garante que a linha tenha pelo menos o número máximo de colunas.
                    # Preenche com strings vazias para simular os delimitadores extras do Excel.
                    if len(row_list) < max_column_in_sheet:
                        row_list.extend([''] * (max_column_in_sheet - len(row_list)))
                    
                    # Converte valores para string (None vira "") e limita ao max_column_in_sheet
                    final_row = ["" if v is None else str(v) for v in row_list[:max_column_in_sheet]]
                    
                    writer.writerow(final_row)

                # Converte para bytes UTF-8 com BOM (Byte Order Mark)
                csv_bytes = csv_buffer.getvalue().encode("utf-8-sig")

                # Nome do arquivo
                base_name = file.name.rsplit(".", 1)[0]
                csv_name = f"{base_name}.csv"

                # Grava no ZIP
                zipf.writestr(csv_name, csv_bytes)

            except Exception as e:
                st.error(f"Erro ao processar {file.name}: {e}")

    st.success("Conversão concluída em CSV UTF-8!")
    st.download_button(
        label="📥 Baixar CSVs (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="csv_utf8.zip",
        mime="application/zip"
    )