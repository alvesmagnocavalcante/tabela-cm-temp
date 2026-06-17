import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# ==========================================
# 1. CONFIGURAÇÕES INICIAIS DA PÁGINA
# ==========================================
# Define o título da aba do navegador, o ícone e força o layout a usar toda a largura da tela
st.set_page_config(page_title="Auditoria de Tarifas Carmel", page_icon="🏨", layout="wide")

# Estilização CSS personalizada para melhorar o design visual do painel
st.markdown("""
    <style>
    .main .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    h1 { color: #1E3A8A; font-weight: 700; }
    h3 { color: #2C3E50; margin-top: 1.5rem; }
    .stAlert p { margin-bottom: 0; font-weight: 500; }
    </style>
""", unsafe_allow_html=True)

# Títulos exibidos no topo da aplicação
st.title("🏨 Painel de Auditoria de Tarifas")
st.markdown("Carregue as planilhas para identificar divergências de forma visual e categorizada.")

# Inicializa o "cofre de memória" (session_state) do Streamlit. 
# Isso impede que o painel perca os dados cruzados caso o usuário clique em algum filtro na tela.
if 'df_master' not in st.session_state:
    st.session_state.df_master = None

# ==========================================
# 2. INTERFACE DE UPLOAD DE ARQUIVOS
# ==========================================
# Cria uma caixa com borda contendo duas colunas para o upload das planilhas
with st.container(border=True):
    col1, col2 = st.columns(2)
    # Upload da matriz Base (ordem lógica de negócio: primeiro a base oficial)
    file_carmel = col1.file_uploader("📂 Tabela Base (MODELO TARIFA CARMEL)", type=["xls", "xlsx", "csv"])
    # Upload da matriz de distribuição (IO)
    file_io = col2.file_uploader("📂 Tabela de Envio (MODELO IO)", type=["xls", "xlsx", "csv"])

# ==========================================
# 3. FUNÇÕES DE SUPORTE E REGRAS DE NEGÓCIO
# ==========================================

def load_file(file):
    """
    Lê o arquivo de forma inteligente.
    Se for CSV, tenta ler com vírgula (padrão internacional). Se falhar, 
    usa ponto e vírgula (padrão Brasil). Retira o cabeçalho automático (header=None)
    para evitar que quebras de linha estéticas no Excel quebrem o código.
    """
    if file.name.lower().endswith('.csv'):
        try:
            df = pd.read_csv(file, header=None)
            if len(df.columns) <= 1:
                file.seek(0)
                df = pd.read_csv(file, sep=';', header=None)
            return df
        except:
            file.seek(0)
            return pd.read_csv(file, sep=';', header=None)
    return pd.read_excel(file, header=None)

def normalize_category(cat):
    """
    Padroniza os nomes das categorias de quartos.
    Remove prefixos visuais como 'NORDESTINA' para facilitar o match (cruzamento).
    Ex: 'NORDESTINA DO MAR' vira apenas 'MAR'.
    """
    c = str(cat).upper().strip()
    c = c.replace('NORDESTINA ', '')
    if c == 'DO MAR': c = 'MAR'
    return c

def normalize_capacity(cap):
    """
    Tradução de nomenclaturas de capacidade (acomodação).
    Garante que as duas tabelas falem a mesma língua antes de comparar.
    """
    c = str(cap).upper().strip().replace(' ', '')
    if c == 'DBL': return 'DPL'
    if c == 'STP': return 'SEX'
    if c.startswith('CR'): return c.replace('CR', 'CHD') # Transforma CR1, CR2 em CHD1, CHD2
    return c

def export_to_excel(df):
    """
    Gera um arquivo Excel na memória RAM (sem salvar no disco do servidor)
    pronto para o usuário fazer o download ao final da auditoria.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatório de Auditoria')
    return output.getvalue()

# ==========================================
# 4. MOTOR PRINCIPAL DE CRUZAMENTO DE DADOS
# ==========================================
# O processamento só começa se o usuário clicar no botão e tiver anexado os dois arquivos
if st.button("🔍 Iniciar Cruzamento de Dados", type="primary", use_container_width=True):
    if file_carmel and file_io:
        with st.spinner("Mapeando matrizes de dados..."): # Mostra barra de carregamento
            try:
                # Carrega as tabelas cruas na memória
                df_carmel_raw = load_file(file_carmel)
                df_io_raw = load_file(file_io)

                # -----------------------------------------------------
                # PASSO 4.1: EXTRAÇÃO DE DADOS DA BASE CARMEL
                # -----------------------------------------------------
                cat_row_idx = None
                
                # Procura a linha exata onde estão os cabeçalhos das categorias (ex: 'ÁRVORE')
                for i in range(min(15, len(df_carmel_raw))):
                    row_vals = [str(x).upper() for x in df_carmel_raw.iloc[i].values]
                    if any('ÁRVORE' in val for val in row_vals):
                        cat_row_idx = i
                        break

                prices_carmel = {} # Dicionário para armazenar as tarifas da Base
                
                if cat_row_idx is not None:
                    col_to_cat = {}
                    # Mapeia qual coluna pertence a qual categoria
                    for col_idx, val in enumerate(df_carmel_raw.iloc[cat_row_idx].values):
                        val_str = str(val).upper().strip()
                        if val_str not in ['NAN', 'NONE', ''] and 'SITE' not in val_str and 'PREÇO' not in val_str:
                            col_to_cat[col_idx] = normalize_category(val_str)
                            
                    # Varre as linhas abaixo do cabeçalho caçando os valores financeiros
                    for i in range(cat_row_idx + 1, len(df_carmel_raw)):
                        cap = None
                        # Procura a capacidade (SGL, DBL, etc) nas primeiras 5 colunas
                        for col_idx in range(5):
                            val = str(df_carmel_raw.iloc[i, col_idx]).upper().strip()
                            if val in ['SGL', 'DBL', 'DPL', 'TPL', 'QDP', 'QTP', 'STP', 'SEX'] or val.startswith('CR') or val.startswith('CHD'):
                                cap = normalize_capacity(val)
                                break
                        
                        # Se encontrou a capacidade, guarda o preço de acordo com a categoria
                        if cap:
                            for col_idx, cat in col_to_cat.items():
                                if col_idx < len(df_carmel_raw.columns):
                                    val = df_carmel_raw.iloc[i, col_idx]
                                    if pd.notna(val):
                                        # Limpa 'R$' e converte o texto do preço em número inteiro
                                        v_str = str(val).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                        try: prices_carmel[(cat, cap)] = int(float(v_str))
                                        except ValueError: pass

                # -----------------------------------------------------
                # PASSO 4.2: EXTRAÇÃO DE DADOS DO MODELO IO
                # -----------------------------------------------------
                header_row_idx = None
                
                # Localiza a linha de cabeçalho do IO procurando por padrões como '-SGL'
                for i in range(min(15, len(df_io_raw))):
                    row_vals = [str(x).upper() for x in df_io_raw.iloc[i].values]
                    if any('-SGL' in val for val in row_vals) or any('ÁRVORE' in val for val in row_vals):
                        header_row_idx = i
                        break
                        
                prices_io = {}
                io_expected_keys = set() # Guarda as colunas que o IO realmente exige (Escopo Oficial)
                
                if header_row_idx is not None:
                    col_names = df_io_raw.iloc[header_row_idx].values
                    
                    # Constrói o escopo separando o cabeçalho (ex: 'FRUTO-DPL' -> Categoria: FRUTO, Acomodação: DPL)
                    for col_name in col_names:
                        c_name = str(col_name).upper().strip()
                        if '-' in c_name:
                            parts = c_name.rsplit('-', 1)
                            if len(parts) == 2:
                                cat_io_exp = normalize_category(parts[0])
                                cap_io_exp = normalize_capacity(parts[1])
                                io_expected_keys.add((cat_io_exp, cap_io_exp))
                    
                    # Acha a linha que contém o texto "MELHOR PREÇO" ou "MELHOR TARIFA"
                    row_prices = None
                    for i in range(header_row_idx + 1, len(df_io_raw)):
                        row_vals = [str(x).upper() for x in df_io_raw.iloc[i].values[:4]]
                        if any('MELHOR PREÇO' in v or 'MELHOR TARIFA' in v for v in row_vals):
                            row_prices = df_io_raw.iloc[i].values
                            break
                    
                    # Salva os preços do IO no dicionário
                    if row_prices is not None:
                        for col_idx, col_name in enumerate(col_names):
                            c_name = str(col_name).upper().strip()
                            if '-' in c_name:
                                parts = c_name.rsplit('-', 1)
                                if len(parts) == 2:
                                    cat_io = normalize_category(parts[0])
                                    cap_io = normalize_capacity(parts[1])
                                    
                                    if col_idx < len(row_prices):
                                        val = str(row_prices[col_idx]).strip()
                                        if val.upper() not in ['NAN', 'NONE', '']:
                                            v_str = val.replace('R$', '').replace('.', '').replace(',', '.').strip()
                                            try: prices_io[(cat_io, cap_io)] = int(float(v_str))
                                            except ValueError: pass

                # -----------------------------------------------------
                # PASSO 4.3: COMPARAÇÃO (A AUDITORIA EM SI)
                # -----------------------------------------------------
                comparison = []
                # Une todas as combinações (Categoria + Acomodação) achadas nas duas planilhas
                all_keys = set(prices_carmel.keys()).union(set(prices_io.keys()))

                for k in sorted(all_keys):
                    cat, cap = k
                    val_carmel = prices_carmel.get(k, None)
                    val_io = prices_io.get(k, None)
                    
                    # REGRA DE OURO: Ignora tarifas extras da Base que não são pedidas pelo IO (evita falso erro de 'CHD4')
                    if val_io is None and k not in io_expected_keys:
                        continue
                        
                    # Motor de regras lógicas de avaliação
                    if val_carmel is None:
                        status = "Faltando na Base"
                        cor_status = "🔴 Ausente na Base"
                    elif val_io is None:
                        status = "Faltando no IO"
                        cor_status = "🟡 Ausente no IO"
                    elif val_carmel != val_io:
                        diff = abs(val_carmel - val_io)
                        # Margem de tolerância: Diferenças de arredondamento de até 1 real são aceitas
                        if diff <= 1:
                            status = "Correto (Arredondado)"
                            cor_status = "⚪ OK (Ajuste R$1)"
                        else:
                            status = "Inconsistência de Valor"
                            cor_status = "🔴 Erro de Valor"
                    else:
                        status = "Correto"
                        cor_status = "🟢 OK"
                        
                    # Adiciona a linha ao relatório final
                    comparison.append({
                        "Categoria": cat,
                        "Acomodação": cap,
                        "Valor Base (Carmel)": val_carmel,
                        "Valor Envio (IO)": val_io,
                        "Status": status,
                        "Análise": cor_status
                    })

                # Verifica se houve erro geral de formatação
                if len(comparison) == 0:
                    st.error("❌ Não conseguimos cruzar os dados. As nomenclaturas das categorias estão muito diferentes do padrão.")
                else:
                    st.success("✅ Cruzamento rigoroso realizado com sucesso! Matrizes mapeadas.")
                    # Salva os resultados no cofre da sessão
                    st.session_state.df_master = pd.DataFrame(comparison)
                
            except Exception as e:
                st.error(f"Erro inesperado no processamento: {e}")
    else:
        st.warning("⚠️ Insira as duas planilhas para gerar a auditoria.")

# ==========================================
# 5. RENDERIZAÇÃO DOS DASHBOARDS E TABELAS
# ==========================================
# Este bloco só aparece se o cruzamento de dados já tiver ocorrido com sucesso
if st.session_state.df_master is not None:
    df_master = st.session_state.df_master
    
    st.divider()
    st.markdown("### 📊 Visão Geral da Auditoria")
    
    # Filtra as volumetrias para os Cards
    erros = df_master[df_master['Status'] == 'Inconsistência de Valor']
    ausentes = df_master[df_master['Status'].str.contains('Faltando')]
    corretos = df_master[df_master['Status'].str.contains('Correto')]
    
    # Cards de Métricas Superiores
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Valores Incorretos", len(erros))
    c2.metric("Itens Ausentes", len(ausentes))
    c3.metric("Itens Corretos", len(corretos))
    c4.metric("Total Auditado", len(df_master))

    st.markdown("---")
    
    # Prepara o espaço para os dois gráficos
    col_graf1, col_graf2 = st.columns(2)
    
    # Define as cores oficiais de cada status para manter consistência nos gráficos
    color_map = {
        "Correto": "#10B981",              # Verde
        "Correto (Arredondado)": "#6EE7B7", # Verde Claro
        "Inconsistência de Valor": "#EF4444", # Vermelho
        "Faltando na Base": "#F97316",      # Laranja
        "Faltando no IO": "#EAB308"         # Amarelo
    }
    
    with col_graf1:
        # Gráfico de Rosca: Mostra a saúde geral da operação
        status_counts = df_master['Status'].value_counts().reset_index()
        status_counts.columns = ['Status', 'Quantidade']
        
        fig_pie = px.pie(status_counts, values='Quantidade', names='Status', hole=0.45,
                         color='Status', color_discrete_map=color_map,
                         title="Distribuição de Resultados")
        fig_pie.update_traces(textposition='inside', textinfo='percent+label', showlegend=False)
        fig_pie.update_layout(margin=dict(t=40, b=0, l=0, r=0))
        st.plotly_chart(fig_pie, use_container_width=True)
        
    with col_graf2:
        # Gráfico Sunburst: Mostra a hierarquia (Onde os erros estão acontecendo?)
        df_problemas = df_master[~df_master['Status'].str.startswith('Correto')]
        
        if not df_problemas.empty:
            fig_sun = px.sunburst(
                df_problemas, 
                path=['Status', 'Categoria'], # Centro do gráfico é o Status, as bordas são as Categorias
                color='Status',
                color_discrete_map=color_map,
                title="Distribuição de Falhas"
            )
            
            fig_sun.update_traces(
                textinfo="label+value",
                insidetextorientation='radial'
            )
            fig_sun.update_layout(margin=dict(t=40, b=10, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_sun, use_container_width=True)
        else:
            # Se a matriz estiver perfeita, exibe mensagem de sucesso em vez de gráfico vazio
            st.info("🎉 Excelente! Nenhum erro ou ausência encontrada nas categorias.")

    st.markdown("---")
    
    # -----------------------------------------------------
    # PASSO 5.1: TABELAS DE DETALHAMENTO (EXPANDERS)
    # -----------------------------------------------------
    # Seletor para o usuário escolher o que quer ver nas tabelas abaixo
    opcao_visao = st.radio(
        "Filtre a visão da Tabela abaixo:",
        ["Mostrar Apenas Erros e Ausências", "Mostrar Apenas Corretos", "Mostrar Tudo"],
        horizontal=True
    )

    # Aplica o filtro selecionado no DataFrame de exibição
    if opcao_visao == "Mostrar Apenas Erros e Ausências":
        df_exibicao = df_master[~df_master['Status'].str.contains('Correto')]
    elif opcao_visao == "Mostrar Apenas Corretos":
        df_exibicao = df_master[df_master['Status'].str.contains('Correto')]
    else:
        df_exibicao = df_master

    if not df_exibicao.empty:
        categorias = df_exibicao['Categoria'].unique()
        st.markdown(f"### 📋 Detalhamento ({opcao_visao})")
        
        # Agrupa os resultados por categoria usando sanfonas expansíveis (Expander)
        for cat in categorias:
            with st.expander(f"🏢 CATEGORIA: {cat}", expanded=True):
                # Seleciona e ordena as colunas que importam na tela
                df_cat = df_exibicao[df_exibicao['Categoria'] == cat][['Acomodação', 'Valor Base (Carmel)', 'Valor Envio (IO)', 'Análise']]
                
                # Renderiza a tabela formatando as colunas financeiras
                st.dataframe(
                    df_cat,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Valor Base (Carmel)": st.column_config.NumberColumn(format="R$ %.2f"),
                        "Valor Envio (IO)": st.column_config.NumberColumn(format="R$ %.2f")
                    }
                )
    else:
        st.info("👍 Tudo limpo no painel baseado no seu filtro atual!")

    # -----------------------------------------------------
    # PASSO 5.2: EXPORTAÇÃO FINAL
    # -----------------------------------------------------
    st.divider()
    st.markdown("### 💾 Exportar Resultados")
    
    # Chama a função lá de cima para transformar o dataframe exibido em .xlsx
    excel_file = export_to_excel(df_exibicao)
    
    # Botão de download nativo do Streamlit
    st.download_button(
        label=f"⬇️ Baixar Planilha (.xlsx)",
        data=excel_file,
        file_name="Auditoria_Tarifas_Carmel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )