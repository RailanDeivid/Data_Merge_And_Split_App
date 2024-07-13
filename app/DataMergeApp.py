import streamlit as st
import pandas as pd
import io
from streamlit_option_menu import option_menu
import zipfile

# Configuração da página
st.set_page_config(page_title="DataMergeApp", page_icon=":file_folder:", layout="wide")

# Função para combinar arquivos
def combinar_arquivos():
    st.subheader("Combinar Arquivos")
    st.write("Faça upload de múltiplos arquivos XLSX ou CSV para combiná-los em um único arquivo XLSX.")
    st.info("""Certifique-se de que os arquivos tenham a mesma quantidade de colunas com as mesmas nomenclaturas. Caso sejam arquivos de Excel com várias abas a serem combinadas, verifique se os nomes das abas são consistentes em todos os arquivos selecionados.""", icon="ℹ️")

    
    # Upload dos arquivos
    uploaded_files = st.file_uploader("Escolha arquivos XLSX ou CSV", accept_multiple_files=True, type=['xlsx', 'csv'])
    
    if uploaded_files:
        # Checkbox para inserir ou não a coluna de origem, aparece apenas se o arquivo for XLSX
        if any(file.name.endswith('.xlsx') for file in uploaded_files):
            add_origin_column = st.checkbox("Adicionar coluna com o nome da origem", value=False)
            origin_column_name = None
            create_column = False
            
            if add_origin_column:
                origin_column_name = st.text_input("Nome da coluna de origem", value="Aba Origem")
                create_column = st.button("Criar Coluna")
        
        # Opções para separador e encoding
        if any(file.name.endswith('.csv') for file in uploaded_files):
            cols = st.columns(5)
            with cols[0]:
                sep = st.selectbox("Selecione o separador CSV:", [',', ';'])
            with cols[1]:
                encoding = st.selectbox("Selecione o encoding do arquivo CSV:", ['utf-8', 'latin1', 'iso-8859-1'])
        
        # Lista para armazenar DataFrames de cada arquivo
        dfs = []
        
        for file in uploaded_files:
            # Verifica se o arquivo é XLSX ou CSV
            if file.name.endswith('.xlsx'):
                # Carrega o arquivo Excel
                xls = pd.ExcelFile(file)
                
                # Lista de abas disponíveis no arquivo
                sheet_names = xls.sheet_names
                
                # Checkbox para seleção de todas as abas (inicializado como True por padrão)
                select_all_sheets = st.checkbox(f"Selecionar todas as abas de '{file.name}'", value=True)
                
                if select_all_sheets:
                    selected_sheets = sheet_names
                else:
                    selected_sheets = st.multiselect(f"Selecione as abas de '{file.name}'", sheet_names, default=sheet_names)
                
                
                for sheet in selected_sheets:
                    # Lê o DataFrame da aba selecionada
                    df = pd.read_excel(file, sheet_name=sheet)
                    
                    # Adiciona uma coluna com o nome da aba origem se a opção estiver selecionada e o botão for clicado
                    if add_origin_column and create_column:
                        df[origin_column_name] = sheet
                    
                    dfs.append(df)
  
            elif file.name.endswith('.csv'):
                    try:  
                        # Lê o DataFrame do arquivo CSV com as opções selecionadas
                        df = pd.read_csv(file, sep=sep, encoding=encoding)
                        nome_sheet = file.name.rfind('.')
                        df['Arquivo Origem'] = file.name[:nome_sheet]
                        dfs.append(df)
                    except UnicodeDecodeError:
                        st.error(f"Erro ao ler o arquivo '{file.name}'. Verifique o encoding selecionado.")
                   
        if dfs:
            # Concatena todos os DataFrames em um único DataFrame
            combined_data = pd.concat(dfs, ignore_index=True)
            
            # Mostra os dados combinados
            st.subheader("Dados Combinados:")
            st.write(combined_data.head(10))
            
            # Download do arquivo combinado
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                combined_data.to_excel(writer, index=False, sheet_name='DadosCombinados')
            
            output.seek(0)
            # Botão de download com o nome personalizado do arquivo
            st.download_button(
                label="Baixar Dados Combinados",
                data=output,
                file_name="DadosCombinados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    pass

# Dicionário para armazenar os DataFrames filtrados
filtered_data_dict = {}

# Função para separar arquivos
def separar_arquivos():
    global filtered_data_dict  # Acesso ao dicionário global

    st.subheader("Separar Arquivos")
    st.write("Selecione um arquivo XLSX para separá-lo com base em uma coluna ou aba específica ou um arquivo CSV")

    # Upload do arquivo
    uploaded_file = st.file_uploader("Escolha um arquivo XLSX ou CSV", type=['xlsx', 'csv'])

    if uploaded_file:
        if uploaded_file.name.endswith('.xlsx'):
            # Se o arquivo for XLSX, mostra opções para selecionar aba e coluna
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names

            cols = st.columns(5)
            with cols[0]:
                selected_sheet = st.selectbox("Selecione a aba:", sheet_names)

            # Lê o DataFrame da aba selecionada
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        
            # Seleção da coluna para separação
            col_options = df.columns.tolist()
            with cols[1]:
                selected_column = st.selectbox("Selecione a coluna para separar os dados:", col_options)

        elif uploaded_file.name.endswith('.csv'):
            try:
                # Se o arquivo for CSV, mostra opções para selecionar separador, encoding e coluna
                cols = st.columns(5)
                with cols[0]:
                    sep = st.selectbox("Selecione o separador CSV:", [',', ';'])
                with cols[1]:
                    encoding = st.selectbox("Selecione o encoding do arquivo CSV:", ['utf-8', 'latin1', 'iso-8859-1'])

                # Lê o DataFrame do arquivo CSV com as opções selecionadas
                df = pd.read_csv(uploaded_file, sep=sep, encoding=encoding)

                # Seleção da coluna para separação
                col_options = df.columns.tolist()
                with cols[2]:
                    selected_column = st.selectbox("Selecione a coluna para separar os dados:", col_options)

            except UnicodeDecodeError:
                st.error(f"Erro ao ler o arquivo '{uploaded_file.name}'. Verifique o encoding selecionado.")
                return

        # Mostra os dados originais
        st.subheader("Dados Originais:")
        st.write(df.head(10))

        # Obtém os valores únicos da coluna selecionada
        unique_values = df[selected_column].unique()

        # Cria um botão para separar os dados com base na escolha do usuário
        if st.button("Separar Dados"):
            for value in unique_values:
                # Filtra o DataFrame com base no valor único da coluna selecionada
                filtered_data = df[df[selected_column] == value]

                # Armazena o DataFrame filtrado no dicionário
                filtered_data_dict[value] = filtered_data

                # Mostra os dados filtrados
                st.subheader(f"Dados separados para '{value}':")
                st.write(filtered_data.head(10))

                # Download do arquivo separado
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    filtered_data.to_excel(writer, index=False, sheet_name=f'Dados_{value}')
                output.seek(0)
                st.download_button(
                    label=f"Baixar Dados para '{value}'",
                    data=output,
                    file_name=f"{value}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # Botão para baixar todos os arquivos separados
            st.markdown("---")
            if st.button("Baixar Todos os Arquivos Separados"):
                zip_output = io.BytesIO()
                with zipfile.ZipFile(zip_output, 'w') as zipf:
                    for value, data in filtered_data_dict.items():
                        filename = f"Dados_{value}.xlsx"  # Nome do arquivo no zip
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            data.to_excel(writer, index=False, sheet_name=f'Dados_{value}')
                        output.seek(0)
                        zipf.writestr(filename, output.read())
                
                zip_output.seek(0)
                st.download_button(
                    label="Baixar Todos os Arquivos",
                    data=zip_output,
                    file_name="dados_separados.zip",
                    mime="application/zip"
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
    separar_arquivos()