import hmac
import streamlit as st
import pickle
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
from st_aggrid import AgGrid, GridOptionsBuilder
import io

st.set_page_config(
    page_title="Gestor financeiro",
    page_icon="💲",
    layout="centered",
    initial_sidebar_state="auto",
    menu_items={
    }
)

st.set_option('client.showErrorDetails', True)

def check_password():
    """Returns `True` if the user had the correct password."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if hmac.compare_digest(st.session_state["password"], st.secrets["password"]):
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store the password.
        else:
            st.session_state["password_correct"] = False

    # Return True if the password is validated.
    if st.session_state.get("password_correct", False):
        return True

    # Show input for password.
    st.text_input(
        "Password", type="password", on_change=password_entered, key="password"
    )
    if "password_correct" in st.session_state:
        st.error("😕 Password incorrect")
    return False


if not check_password():
    st.stop()  # Do not continue if check_password is not True.

# Escopos necessários para acessar o Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/drive.file']

def authenticate():
    """Autenticação com o Google Drive usando as credenciais do Streamlit secrets"""
    google_secrets = st.secrets["google"]
    credentials_dict = {
        "type": "service_account",
        "project_id": google_secrets["project_id"],
        "private_key_id": google_secrets["private_key_id"],
        "private_key": google_secrets["private_key"],
        "client_email": google_secrets["client_email"],
        "client_id": google_secrets["client_id"],  # Caso necessário
        "auth_uri": google_secrets["auth_uri"],
        "token_uri": google_secrets["token_uri"],
        "auth_provider_x509_cert_url": google_secrets["auth_provider_x509_cert_url"],
        "client_x509_cert_url": google_secrets["client_x509_cert_url"],
        "universe_domain": google_secrets["universe_domain"]
    }
    credentials = service_account.Credentials.from_service_account_info(credentials_dict, scopes=SCOPES)
    service = build('drive', 'v3', credentials=credentials)
    test_authentication(service)
    return service

def test_authentication(service):
    """Teste simples para verificar se a autenticação foi bem-sucedida"""
    try:
        results = service.files().list(pageSize=1).execute()
        if 'files' in results and len(results['files']) > 0:
            st.success("Autenticação bem-sucedida!")
        else:
            st.error("Nenhum arquivo encontrado, mas autenticação bem-sucedida.")
    except Exception as e:
        st.error(f"Erro de autenticação: {e}")

def list_files(service, folder_id=None):
    """Lista arquivos e pastas. Se `folder_id` for passado, lista o conteúdo dessa pasta."""
    query = f"'{folder_id}' in parents" if folder_id else "trashed = false"
    results = service.files().list(q=query, pageSize=10, fields="files(id, name, mimeType)").execute()
    items = results.get('files', [])
    return items

def main():
    st.title("Gestor financeiro")

    # Adicionando a barra lateral
    with st.sidebar:
        st.header("Opções")
        st.text("Escolha uma das opções abaixo para navegar")
    
        # Botão de autenticação
        #if st.button("Reautenticar"):
            #service = authenticate()
        
        # Exibir mensagem de status
        #st.text("Status da autenticação:")
        #st.text("Autenticação: Bem-sucedida")

    # Autenticação no Google Drive
    service = authenticate()

    # Listar arquivos e pastas na raiz
    items = list_files(service)

    # Separar pastas e arquivos
    folders = [item for item in items if item['mimeType'] == 'application/vnd.google-apps.folder']
    files = [item for item in items if item['mimeType'] != 'application/vnd.google-apps.folder']

    # Mostrar pastas na barra lateral
    selected_folder_name = st.sidebar.selectbox("Escolha uma pasta", [folder['name'] for folder in folders] if folders else ["Sem pastas"])
    selected_folder = next((folder for folder in folders if folder['name'] == selected_folder_name), None)
    
    # Mostrar arquivos na barra lateral
    if selected_folder:
        selected_folder_id = selected_folder['id']
        folder_files = list_files(service, folder_id=selected_folder_id)
        selected_file_name = st.sidebar.selectbox("Escolha um arquivo dentro da pasta", [file['name'] for file in folder_files])
    else:
        selected_file_name = st.sidebar.selectbox("Escolha um arquivo na raiz", [file['name'] for file in files])

    # Buscar o arquivo selecionado
    if selected_file_name:
        if selected_folder:
            selected_file = next(file for file in folder_files if file['name'] == selected_file_name)
        else:
            selected_file = next(file for file in files if file['name'] == selected_file_name)

        file_id = selected_file['id']

        # Baixar o arquivo Excel
        file = service.files().get_media(fileId=file_id).execute()
        file_path = f"temp_{selected_file_name}.xlsx"
        with open(file_path, 'wb') as f:
            f.write(file)

        # Carregar e exibir o conteúdo do Excel com Pandas
        df = pd.read_excel(file_path)
        st.write("Conteúdo do arquivo Excel:", df)

        # Permitir edição da tabela usando st-aggrid
        gb = GridOptionsBuilder.from_dataframe(df)
        gb.configure_pagination()  # Ativa paginação
        gb.configure_default_column(editable=True)  # Permite edição
        grid_options = gb.build()

        # Exibir a tabela editável com AgGrid
        edited_df = AgGrid(df, gridOptions=grid_options, editable=True, fit_columns_on_grid_load=True)

        # Permitir o envio do arquivo editado para o Google Drive
        if st.button("Salvar alterações"):
            edited_df['data'].to_excel(file_path, index=False)  # Salva as alterações no Excel
            # Fazer upload do arquivo editado para o Google Drive
            media = MediaIoBaseDownload(io.open(file_path, 'rb'), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            service.files().update(fileId=file_id, media_body=media).execute()
            st.success("Alterações salvas no Google Drive!")

if __name__ == "__main__":
    main()
