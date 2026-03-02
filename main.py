import streamlit as st
import pandas as pd
import os

# 1. CONFIGURAÇÃO DA PÁGINA (Deve ser a primeira coisa!)
st.set_page_config(page_title="Quantitativo de Materiais", page_icon="☀️", layout="centered")

# --- DEFINIÇÕES DE ARQUIVO ---
ARQUIVO_EXCEL = "seu_arquivo_com_macros.xlsm"
ABA_ALVO = "LISTA_MÃE"

# --- FUNÇÕES DE APOIO ---
def carregar_dados():
    try:
        return pd.read_excel(ARQUIVO_EXCEL, sheet_name=ABA_ALVO, engine='openpyxl')
    except Exception as e:
        # Se o arquivo não existir ou der erro, cria um dataframe vazio para não travar o app
        return pd.DataFrame(columns=["Nome", "Preço"])

def salvar_dados(df):
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=ABA_ALVO, index=False)

# --- ESTILIZAÇÃO CSS (Global) ---
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Afacad:wght@400;500&family=Comfortaa:wght@300;500;600&display=swap');
    
    .stApp {{
        background-color: #ffffff;
        background-image: url('data:image/svg+xml;charset=utf8,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20viewBox%3D%220%200%201024%201024%22%20width%3D%221024%22%20height%3D%221024%22%20preserveAspectRatio%3D%22none%22%3E%20%3Cpath%20d%3D%22M256%20192v128M192%20256h128M768%20704v128M704%20768h128%22%20stroke%3D%22rgba(0,0,0,0.15)%22%20stroke-width%3D%223%22%20fill%3D%22none%22%2F%3E%3C%2Fsvg%3E');
        background-size: 80px;
    }}
    
    .main .block-container {{
        background-color: #FFFFFF;
        border: 2px solid rgba(10,10,10,0.56);
        padding: 2.5rem;
        box-shadow: 15px 15px 0px rgba(0,0,0,0.1);
        font-family: 'Afacad', sans-serif;
        color: #474747;
    }}

    h1, h2 {{
        font-family: 'Comfortaa', sans-serif !important;
        text-transform: uppercase;
        letter-spacing: 0.15rem;
        text-align: center;
    }}

    .section-title {{
        font-family: 'Comfortaa', sans-serif;
        font-size: 0.75rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.1rem;
        color: #3d3d3d;
        margin: 1.5rem 0 0.8rem 0;
        border-bottom: 1px solid rgba(0,0,0,0.1);
        padding-bottom: 5px;
    }}

    .stButton>button {{
        width: 100%;
        background-color: #3d3d3d !important;
        color: #FFFFFF !important;
        border-radius: 0px !important;
        text-transform: uppercase;
        letter-spacing: 0.1rem;
    }}
    </style>
    """, unsafe_allow_html=True)

# --- CRIAÇÃO DAS ABAS ---
tab1, tab2, tab3 = st.tabs(["🏠 Home", "📋 Lista Mãe", "📊 Cálculos"])

# --- ABA 1: HOME ---
with tab1:
    st.write("# Quantitativo Solar☀️ ")
    st.write("### Sistema Integrado de Engenharia Solar")
    st.info("Utilize as abas acima para gerenciar materiais ou realizar cálculos de quantitativo.")

# --- ABA 2: LISTA MÃE ---
with tab2:
    st.write("## 📋 Gerenciar Lista Mãe")
    
    # Adição
    with st.expander("➕ Adicionar Novo Material"):
        with st.form("add_form"):
            # Dividindo em 5 colunas para caber todos os campos na mesma linha
            col1, col2, col3, col4, col5 = st.columns(5)
            categoria_novo = col1.text_input("Categoria")
            familia_novo = col2.text_input("Família") # Mantive o acento para padronizar
            nome_novo = col3.text_input("Nome")
            preco_novo = col4.number_input("Preço", min_value=0.0, step=0.01)
            referencia_nova = col5.text_input("Referência")
            
            if st.form_submit_button("Confirmar Cadastro"):
                if nome_novo: # Pequena trava de segurança: só cadastra se tiver nome
                    df = carregar_dados()
                    
                    # Criando o novo item com as colunas batendo exatamente com o Excel
                    novo_item = pd.DataFrame({
                        "Categoria": [categoria_novo],
                        "Família": [familia_novo],
                        "Nome": [nome_novo], 
                        "Preço": [preco_novo], 
                        "Referência": [referencia_nova]
                    })
                    
                    df = pd.concat([df, novo_item], ignore_index=True)
                    salvar_dados(df)
                    st.success(f"Item '{nome_novo}' adicionado com sucesso!")
                    st.rerun() # Atualiza a tabela logo abaixo
                else:
                    st.error("Por favor, preencha pelo menos o Nome do material.")

    # Remoção
    with st.expander("🗑️ Remover Material"):
        df_del = carregar_dados()
        if not df_del.empty:
            # Dropdown para escolher o item pelo nome
            item_para_remover = st.selectbox("Selecione para excluir:", df_del["Nome"].unique(), key="del_select")
            
            if st.button("Apagar Definitivamente", type="primary"):
                df_del = df_del[df_del["Nome"] != item_para_remover]
                salvar_dados(df_del)
                st.warning(f"{item_para_remover} removido da LISTA_MÃE.")
                st.rerun()
        else:
            st.info("A lista está vazia no momento.")

    # Visualização dinâmica
    st.write("---")
    st.write("### Visualizando apenas a LISTA_MÃE")
    
    # Aqui ele lê o arquivo atualizado
    dados_atuais = carregar_dados()
    st.dataframe(dados_atuais, use_container_width=True)
# --- ABA 3: CÁLCULOS ---
with tab3:
    st.write("<h1>Quantitativo</h1>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.write('<div class="section-title">Infra e Aterramento</div>', unsafe_allow_html=True)
        n_hastes = st.number_input("Nº de Hastes", value=3, key="hastes")
        secc_at = st.number_input("Cabo Cobre Nu (mm²)", value=35, key="secc_at")
        d_at = st.number_input("Dist. Haste até BEP (m)", value=10.0, key="d_at")
        n_arranjos = st.number_input("Nº de Arranjos (Matrizes)", value=2, key="arranjos")

    with col2:
        st.write('<div class="section-title">Circuitos CC</div>', unsafe_allow_html=True)
        n_strings = st.number_input("Qtd Total de Strings", value=2, min_value=1, key="strings")
        n_sb = st.number_input("Qtd Stringbox", value=1, key="sb")
        d_sb = st.number_input("Dist. SB até Inversor (m)", value=5.0, key="d_sb")
        d_descida = st.number_input("Descida / Prumada (m)", value=15.0, key="descida")

    st.write('<div style="font-size:0.65rem; color:#888; font-weight:600; margin-top:10px">DISTÂNCIAS UNITÁRIAS (M):</div>', unsafe_allow_html=True)
    cols_strings = st.columns(4)
    dists_st = []
    for i in range(int(n_strings)):
        with cols_strings[i % 4]:
            val = st.number_input(f"ST {i+1}", value=15.0, key=f"st_calc_{i}")
            dists_st.append(val)

    spda = st.selectbox("SPDA no local?", ["Não (Equipot. 6mm²)", "Sim (Equipot. 16mm²)"], key="spda_box")

    if st.button("Gerar Lista Completa"):
        somaDistStrings = sum(dists_st)
        cabo_nu = d_at + (2.4 * (max(0, n_hastes - 1)))
        secc_eqp = 16 if "Sim" in spda else 6
        term_m6 = (n_arranjos * 2) + n_sb + 1
        cabo_eqp = (2.5 * term_m6) + d_descida + d_at
        d_sb_real = d_sb if n_sb > 0 else 0
        cabo_cc_parcial = (somaDistStrings + (n_strings * (d_descida + d_sb_real)))

        materiais = [
            ["Cabo de Cobre Nu | " + str(int(secc_at)) + "mm²", f"{cabo_nu:.2f}", "m"],
            ["Haste Aterramento 5/8 (2400mm)", n_hastes, "un"],
            ["Conector Split-Bolt " + str(int(secc_at)) + "mm²", n_hastes * 2, "un"],
            ["Cabo 1kV " + str(secc_eqp) + "mm² HEPR Verde", f"{cabo_eqp:.2f}", "m"],
            ["Terminal de Compressão M6 - " + str(secc_eqp) + "mm²", term_m6, "un"],
            ["Parafuso Aço Inox M6 x 25mm", term_m6, "un"],
            ["Porca de Aço Inox M6", term_m6, "un"],
            ["Arruela Lisa Aço Inox M6", term_m6 * 2, "un"],
            ["Cabo Solar 1,8KVCC Preto (6mm²)", f"{cabo_cc_parcial:.2f}", "m"],
            ["Cabo Solar 1,8KVCC Vermelho (6mm²)", f"{cabo_cc_parcial:.2f}", "m"],
            ["Caixa de Inspeção PVC 300x300", n_hastes, "un"]
        ]

        df_res = pd.DataFrame(materiais, columns=["Material", "Qtd", "Un"])
        st.write("---")
        st.table(df_res)

        csv = df_res.to_csv(index=False).encode('utf-8-sig')
        st.download_button(label="📥 Exportar para Excel", data=csv, file_name='quantitativo_solar.csv', mime='text/csv')