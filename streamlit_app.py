import pandas as pd
import streamlit as st
from google_drive import authenticate, list_files

def main():
    st.title("Editor de Arquivos Excel - Google Drive")

    service = authenticate()

    # Listar arquivos do Google Drive
    files = list_files(service)
    
    # Exibir arquivos disponíveis
    st.write("Arquivos encontrados no Google Drive:")
    for file in files:
        st.write(f"- {file['name']} (ID: {file['id']})")

    # Selecionar arquivo
    selected_file_id = st.selectbox("Escolha um arquivo", [file['name'] for file in files])

    # Buscar o arquivo selecionado
    selected_file = next(file for file in files if file['name'] == selected_file_id)
    file_id = selected_file['id']
    
    # Baixar o arquivo Excel
    file = service.files().get_media(fileId=file_id).execute()
    file_path = f"temp_{selected_file_id}.xlsx"
    with open(file_path, 'wb') as f:
        f.write(file)
    
    # Carregar e exibir o conteúdo do Excel com Pandas
    df = pd.read_excel(file_path)
    st.write("Conteúdo do arquivo Excel:", df)

    # Permitir edição da tabela
    edited_df = st.experimental_data_editor(df)

    # Permitir o envio do arquivo de volta para o Google Drive
    if st.button("Salvar alterações"):
        edited_df.to_excel(file_path, index=False)
        # Fazer upload do arquivo editado para o Google Drive
        file_metadata = {'name': selected_file_id}
        media = MediaIoBaseDownload(io.open(file_path, 'rb'), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        service.files().update(fileId=file_id, media_body=media).execute()
        st.success("Alterações salvas no Google Drive!")

if __name__ == "__main__":
    main()
