import streamlit as st
import pandas as pd
import io
from streamlit_option_menu import option_menu

# Configuração da página
st.set_page_config(page_title="DataMergeApp", page_icon=":file_folder:", layout="wide")

# Função para combinar arquivos
def combinar_arquivos():
    st.subheader("Combinar Arquivos")
    st.write("Faça upload de múltiplos arquivos XLSX ou CSV para combiná-los em um único arquivo XLSX.")
    st.info("Certifique-se de que os arquivos tenham a mesma quantidade de colunas com as mesmas nomenclaturas.", icon=":material/info:")
    
    # Upload dos arquivos
    uploaded_files = st.file_uploader("Escolha arquivos XLSX ou CSV", accept_multiple_files=True, type=['xlsx', 'csv'])
    
    if uploaded_files:
        # Dicionário para armazenar DataFrames de cada aba
        dfs = {}
        
        # Checkbox para seleção de todas as colunas (inicializado como True por padrão)
        select_all_columns = st.checkbox("Selecionar todas as colunas", value=True)
        
        # Checkbox para seleção de todas as abas (inicializado como True por padrão)
        select_all_sheets = st.checkbox("Selecionar todas as abas", value=True)
        
        for file in uploaded_files:
            # Verifica se o arquivo é XLSX ou CSV
            if file.name.endswith('.xlsx'):
                # Carrega o arquivo Excel
                xls = pd.ExcelFile(file)
                
                # Lista de abas disponíveis no arquivo
                sheet_names = xls.sheet_names
                
                # Seleciona todas as abas ou abas específicas
                if select_all_sheets:
                    selected_sheets = sheet_names
                else:
                    selected_sheets = st.multiselect(f"Selecione as abas do arquivo '{file.name}'", sheet_names, default=sheet_names)
                
                for sheet in selected_sheets:
                    # Lê o DataFrame da aba selecionada
                    df = pd.read_excel(file, sheet_name=sheet)
                    
                    # Seleciona todas as colunas ou colunas específicas
                    if select_all_columns:
                        dfs_key = f"{file.name} - {sheet} (Todas as Colunas)"
                        dfs[dfs_key] = df.copy()
                    else:
                        selected_columns = st.multiselect(f"Selecione as colunas de '{file.name}' - '{sheet}'", df.columns.tolist(), df.columns.tolist())
                        df_selected = df[selected_columns]
                        dfs_key = f"{file.name} - {sheet} ({', '.join(selected_columns)})"
                        dfs[dfs_key] = df_selected.copy()
                    
            elif file.name.endswith('.csv'):
                # Carrega o arquivo CSV
                df = pd.read_csv(file)
                
                # Seleciona todas as colunas ou colunas específicas
                if select_all_columns:
                    dfs_key = f"{file.name} (Todas as Colunas)"
                    dfs[dfs_key] = df.copy()
                else:
                    selected_columns = st.multiselect(f"Selecione as colunas de '{file.name}'", df.columns.tolist(), df.columns.tolist())
                    df_selected = df[selected_columns]
                    dfs_key = f"{file.name} ({', '.join(selected_columns)})"
                    dfs[dfs_key] = df_selected.copy()
        
        # Cria um novo arquivo Excel combinado
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for dfs_key, df in dfs.items():
                df.to_excel(writer, index=False, sheet_name=dfs_key[:31])  # Limita o tamanho do nome da aba
            
        output.seek(0)
        st.download_button(
            label="Baixar Arquivo Combinado",
            data=output,
            file_name="dados_combinados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Menu de navegação usando option_menu
cols1, cols2, cols3 = st.columns([1, 2, 1])
with cols2:
    selected_page = option_menu(
        menu_title=None,
        options=["Combinar Arquivos", "Separar Arquivos"],
        icons=["files", "file-earmark-break"],
        menu_icon="cast",
        default_index=0,
        orientation="horizontal"
    )

# Lógica de seleção da página
if selected_page == "Combinar Arquivos":
    combinar_arquivos()

elif selected_page == "Separar Arquivos":
    st.subheader("Separar Arquivos")
    st.write("Implementação para separar arquivos ainda será adicionada.")
