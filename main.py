import re
import unicodedata
from datetime import datetime
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Auditoria de Tarifas Carmel",
    page_icon="🏨",
    layout="wide",
)

st.markdown(
    """
    <style>
    .main .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    h1 { color: #1E3A8A; font-weight: 700; }
    h3 { color: #2C3E50; margin-top: 1.5rem; }
    .stAlert p { margin-bottom: 0; font-weight: 500; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🏨 Conferência de Tarifas")
st.markdown(
    "Envie a tabela oficial e os relatórios que deseja conferir. "
    "O sistema encontra os períodos e compara os preços automaticamente."
)

if "df_master" not in st.session_state:
    st.session_state.df_master = None


CAPACITIES = {"SGL", "DBL", "DPL", "TPL", "QDP", "QTP", "STP", "SEX"}


def clean_text(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def remove_accents(value):
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", str(value))
        if not unicodedata.combining(char)
    )


def normalize_category(category):
    value = remove_accents(clean_text(category).upper())
    value = value.replace("NORDESTINA ", "")
    aliases = {
        "DO MAR": "MAR",
        "PASS LOFT": "PASSARO LOFT",
        "PASSARO LOFT": "PASSARO LOFT",
    }
    return aliases.get(value, value)


def normalize_capacity(capacity):
    value = remove_accents(clean_text(capacity).upper()).replace(" ", "")
    aliases = {"DBL": "DPL", "STP": "SEX"}
    value = aliases.get(value, value)
    if value.startswith("CR"):
        value = value.replace("CR", "CHD", 1)
    return value


def is_capacity(value):
    normalized = normalize_capacity(value)
    return (
        normalized in {normalize_capacity(item) for item in CAPACITIES}
        or normalized.startswith("CHD")
    )


def parse_money(value):
    if pd.isna(value) or clean_text(value).upper() in {"", "NAN", "NONE"}:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value).replace("R$", "").replace(" ", "")
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def extract_dates(value):
    dates = []
    for day, month, year in re.findall(r"\b(\d{2})/(\d{2})/(\d{2,4})\b", str(value)):
        if len(year) == 2:
            year = f"20{year}"
        try:
            dates.append(datetime(int(year), int(month), int(day)).date())
        except ValueError:
            continue
    return dates


def infer_period_index(rate_label, periods):
    """Seleciona o período base que contém a primeira data da linha IO."""
    io_dates = extract_dates(rate_label)
    if not io_dates:
        return None
    target = io_dates[0]
    candidates = []

    for index, period in enumerate(periods):
        dates = extract_dates(period["label"])
        intervals = list(zip(dates[::2], dates[1::2]))
        matching_spans = [
            (end - start).days
            for start, end in intervals
            if start <= target <= end
        ]
        if matching_spans:
            candidates.append((min(matching_spans), index))

    return min(candidates)[1] if candidates else None


def reset_file(file):
    if hasattr(file, "seek"):
        file.seek(0)


def list_sheets(file):
    if file.name.lower().endswith(".csv"):
        return ["CSV"]
    reset_file(file)
    if file.name.lower().endswith(".xls"):
        excel = pd.ExcelFile(file, engine="calamine")
    else:
        try:
            excel = pd.ExcelFile(file)
        except Exception:
            # O Calamine também serve como fallback para arquivos OOXML atípicos.
            reset_file(file)
            excel = pd.ExcelFile(file, engine="calamine")
    sheets = excel.sheet_names
    reset_file(file)
    return sheets


def load_file(file, sheet_name=None):
    reset_file(file)
    if file.name.lower().endswith(".csv"):
        try:
            df = pd.read_csv(file, header=None, sep=None, engine="python")
        finally:
            reset_file(file)
        return df

    try:
        if file.name.lower().endswith(".xls"):
            df = pd.read_excel(
                file, sheet_name=sheet_name, header=None, engine="calamine"
            )
        else:
            try:
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
            except Exception:
                reset_file(file)
                df = pd.read_excel(
                    file, sheet_name=sheet_name, header=None, engine="calamine"
                )
    finally:
        reset_file(file)
    return df


def find_header_row(df):
    for row_idx in range(min(30, len(df))):
        values = [normalize_category(value) for value in df.iloc[row_idx].values]
        if any("ARVORE" in value for value in values):
            return row_idx
    return None


def find_base_groups(df):
    """Identifica blocos horizontais como SITE, OMNIBEES VENDA e NET."""
    header_row = find_header_row(df)
    if header_row is None:
        return []

    candidates = []
    for period_col in range(max(0, len(df.columns) - 1)):
        capacity_col = period_col + 1
        matches = 0
        for row_idx in range(header_row + 1, len(df)):
            if is_capacity(df.iat[row_idx, capacity_col]):
                matches += 1
                if matches >= 2:
                    break
        if matches >= 2:
            candidates.append(period_col)

    # Colunas de período adjacentes não podem representar grupos diferentes.
    period_columns = []
    for col in candidates:
        if not period_columns or col > period_columns[-1] + 1:
            period_columns.append(col)

    groups = []
    for index, period_col in enumerate(period_columns):
        next_period_col = (
            period_columns[index + 1]
            if index + 1 < len(period_columns)
            else len(df.columns)
        )
        category_columns = []
        for col in range(period_col + 2, next_period_col):
            category = normalize_category(df.iat[header_row, col])
            if category and category not in {"NAN", "NONE"}:
                category_columns.append((col, category))

        if not category_columns:
            continue

        name = clean_text(df.iat[header_row, period_col]) or f"Bloco {index + 1}"
        groups.append(
            {
                "name": name,
                "header_row": header_row,
                "period_col": period_col,
                "capacity_col": period_col + 1,
                "category_columns": category_columns,
            }
        )
    return groups


def extract_base_periods(df, group):
    periods = []
    period_col = group["period_col"]
    capacity_col = group["capacity_col"]
    header_row = group["header_row"]

    for row_idx in range(header_row + 1, len(df)):
        period = clean_text(df.iat[row_idx, period_col])
        if period and is_capacity(df.iat[row_idx, capacity_col]):
            periods.append({"label": period, "start_row": row_idx})
    return periods


def extract_base_prices(df, group, period):
    prices = {}
    row_idx = period["start_row"]
    start_row = row_idx
    period_col = group["period_col"]
    capacity_col = group["capacity_col"]

    while row_idx < len(df) and is_capacity(df.iat[row_idx, capacity_col]):
        if row_idx > start_row and clean_text(df.iat[row_idx, period_col]):
            break
        capacity = normalize_capacity(df.iat[row_idx, capacity_col])
        for col_idx, category in group["category_columns"]:
            value = parse_money(df.iat[row_idx, col_idx])
            if value is not None:
                prices[(category, capacity)] = value
        row_idx += 1
    return prices


def find_io_layout(df):
    for row_idx in range(min(40, len(df))):
        columns = []
        for col_idx, value in enumerate(df.iloc[row_idx].values):
            name = clean_text(value).upper()
            if "-" not in name:
                continue
            category, capacity = name.rsplit("-", 1)
            if category and is_capacity(capacity):
                columns.append(
                    (
                        col_idx,
                        normalize_category(category),
                        normalize_capacity(capacity),
                    )
                )
        if columns:
            return row_idx, columns
    return None, []


def extract_io_rate_rows(df):
    rows = []
    current_columns = []
    first_columns = []
    context = ""

    for row_idx in range(len(df)):
        detected_columns = []
        for col_idx, value in enumerate(df.iloc[row_idx].values):
            name = clean_text(value).upper()
            if "-" not in name:
                continue
            category, capacity = name.rsplit("-", 1)
            if category and is_capacity(capacity):
                detected_columns.append(
                    (
                        col_idx,
                        normalize_category(category),
                        normalize_capacity(capacity),
                    )
                )

        if detected_columns:
            current_columns = detected_columns
            if not first_columns:
                first_columns = detected_columns

            context_values = []
            if row_idx >= 1:
                context_values = [
                    clean_text(value)
                    for value in df.iloc[row_idx - 1].values[:4]
                    if clean_text(value)
                ]
            context = " | ".join(context_values)
            continue

        if not current_columns:
            continue

        first_cells = " | ".join(
            clean_text(value).upper() for value in df.iloc[row_idx].values[:5]
        )
        normalized = remove_accents(first_cells)
        if "MELHOR PRECO" in normalized or "MELHOR TARIFA" in normalized:
            rate_name = clean_text(df.iat[row_idx, 0]) or f"Linha {row_idx + 1}"
            label_parts = [part for part in [context, rate_name] if part]
            label = " | ".join(label_parts) + f" (linha {row_idx + 1})"
            rows.append(
                {
                    "label": label,
                    "row_idx": row_idx,
                    "columns": current_columns,
                }
            )
    return rows, first_columns


def extract_io_prices(df, rate_row, columns):
    prices = {}
    expected = set()
    for col_idx, category, capacity in columns:
        key = (category, capacity)
        expected.add(key)
        value = parse_money(df.iat[rate_row["row_idx"], col_idx])
        if value is not None:
            prices[key] = value
    return prices, expected


def extract_fluctuation(rate_label):
    parts = [clean_text(part) for part in rate_label.split("|")]
    return parts[3] if len(parts) >= 4 else ""


def compare_prices(
    base_prices,
    io_prices,
    expected,
    source,
    sheet,
    period,
    rate_label,
):
    comparison = []
    all_keys = set(base_prices).union(io_prices)

    for category, capacity in sorted(all_keys):
        key = (category, capacity)
        base_value = base_prices.get(key)
        io_value = io_prices.get(key)

        # Itens extras da base que não fazem parte do relatório IO não são falhas.
        if io_value is None and key not in expected:
            continue

        difference_value = None
        difference_pct = None
        tolerance = (
            min(max(5.0, base_value * 0.001), 15.0)
            if base_value is not None and base_value > 0
            else None
        )
        if base_value is None:
            status = "Faltando na Base"
            analysis = "🔴 Ausente na Base"
        elif io_value is None:
            status = "Faltando no IO"
            analysis = "🟡 Ausente no IO"
        elif io_value <= 0 and base_value > 0:
            status = "Valor Zerado no IO"
            analysis = "🔴 Valor zerado no IO"
            difference_value = io_value - base_value
            difference_pct = -100.0
        elif abs(base_value - io_value) <= 0.01:
            status = "Correto"
            analysis = "🟢 OK"
            difference_value = 0.0
            difference_pct = 0.0
        else:
            difference_value = io_value - base_value
            difference_pct = (
                difference_value / base_value * 100 if base_value else 0
            )
            if abs(difference_value) <= tolerance:
                status = "Correto (Diferença tolerada)"
                analysis = (
                    f"🟡 Diferença de R$ {abs(difference_value):.2f} "
                    f"dentro do limite de R$ {tolerance:.2f}"
                )
            else:
                direction = "acima" if difference_value > 0 else "abaixo"
                status = "Inconsistência de Valor"
                analysis = (
                    f"🔴 {abs(difference_pct):.2f}% {direction} da tabela oficial"
                )

        comparison.append(
            {
                "Arquivo IO": source,
                "Aba IO": sheet,
                "Linha IO": rate_label,
                "Flutuação": extract_fluctuation(rate_label),
                "Período Base": period,
                "Categoria": category,
                "Acomodação": capacity,
                "Valor Base (Carmel)": base_value,
                "Valor Envio (IO)": io_value,
                "Tolerância (R$)": tolerance,
                "Diferença (R$)": (
                    round(difference_value, 2)
                    if difference_value is not None
                    else None
                ),
                "Diferença (%)": (
                    round(difference_pct, 2)
                    if difference_pct is not None
                    else None
                ),
                "Status": status,
                "Análise": analysis,
            }
        )
    return comparison


def export_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Relatório de Auditoria")
    return output.getvalue()


with st.container(border=True):
    col1, col2 = st.columns(2)
    file_base = col1.file_uploader(
        "📘 1. Tabela oficial de tarifas",
        type=["xls", "xlsx", "csv"],
        key="base_file",
        help="Planilha que contém os preços oficiais do hotel.",
    )
    files_io = col2.file_uploader(
        "📄 2. Relatórios para conferência",
        type=["xls", "xlsx", "csv"],
        accept_multiple_files=True,
        key="io_files",
        help="Você pode selecionar vários relatórios de uma só vez.",
    )


base_config = None
if file_base:
    try:
        base_sheets = list_sheets(file_base)
        base_sheet = st.selectbox("Qual aba contém os preços oficiais?", base_sheets)
        base_df = load_file(file_base, base_sheet)
        base_groups = find_base_groups(base_df)

        if not base_groups:
            st.error("Não encontrei preços organizados na aba selecionada.")
        else:
            group_names = [group["name"] for group in base_groups]
            group_name = st.selectbox("Qual tipo de preço deseja conferir?", group_names)
            base_group = base_groups[group_names.index(group_name)]
            base_periods = extract_base_periods(base_df, base_group)
            if not base_periods:
                st.error("Não encontrei períodos de venda nesse conjunto de preços.")
            else:
                base_config = {
                    "df": base_df,
                    "sheet": base_sheet,
                    "group": base_group,
                    "periods": base_periods,
                }
    except Exception as error:
        st.error(f"Não foi possível abrir a tabela oficial: {error}")


io_configs = []
if files_io and base_config:
    st.markdown("### Relatórios encontrados")
    with st.expander("Como a conferência funciona"):
        st.write(
            "O sistema compara diretamente cada preço do relatório com o preço "
            "correspondente na tabela oficial. Valores diferentes, zerados ou "
            "ausentes são destacados para conferência."
        )
        st.write(
            "Uma diferença é aceita quando fica dentro do maior valor entre "
            "R$ 5 e 0,10% do preço oficial, com limite máximo de R$ 15. "
            "Diferenças aceitas continuam visíveis em amarelo."
        )
    for file_index, file_io in enumerate(files_io):
        with st.expander(f"📄 {file_io.name}", expanded=True):
            try:
                io_sheets = list_sheets(file_io)
                io_sheet = st.selectbox(
                    "Aba que contém os preços",
                    io_sheets,
                    key=f"io_sheet_{file_index}_{file_io.name}",
                )
                io_df = load_file(file_io, io_sheet)
                rate_rows, io_columns = extract_io_rate_rows(io_df)
                if not io_columns:
                    st.error("Não reconheci as colunas de quarto e ocupação.")
                    continue
                if not rate_rows:
                    st.error("Não encontrei linhas de preços para conferir.")
                    continue

                available_rows = [
                    row
                    for row in rate_rows
                    if "MELHOR PRECO DISPONIVEL"
                    in remove_accents(row["label"].upper())
                ]
                if not available_rows:
                    st.error("Não encontrei preços na modalidade 'Melhor Preço Disponível'.")
                    continue

                st.success(
                    f"Pronto para conferir: {len(available_rows)} período(s) encontrado(s)."
                )
                io_configs.append(
                    {
                        "name": file_io.name,
                        "sheet": io_sheet,
                        "df": io_df,
                        "rate_rows": available_rows,
                    }
                )
            except Exception as error:
                st.error(
                    f"Não foi possível abrir este arquivo: {error}. "
                    "Se for um .xls antigo ou danificado, abra-o no Excel e "
                    "salve uma nova cópia em .xlsx."
                )


if st.button(
    "🔍 Conferir tarifas",
    type="primary",
    width="stretch",
    disabled=not (base_config and io_configs),
):
    with st.spinner("Conferindo os preços..."):
        results = []
        for config in io_configs:
            for rate_row in config["rate_rows"]:
                period_index = infer_period_index(
                    rate_row["label"], base_config["periods"]
                )
                if period_index is None:
                    st.warning(
                        f"Um período do relatório não foi encontrado na tabela oficial: "
                        f"{rate_row['label']}"
                    )
                    continue
                period = base_config["periods"][period_index]
                base_prices = extract_base_prices(
                    base_config["df"], base_config["group"], period
                )
                io_prices, expected = extract_io_prices(
                    config["df"], rate_row, rate_row["columns"]
                )
                results.extend(
                    compare_prices(
                        base_prices,
                        io_prices,
                        expected,
                        config["name"],
                        config["sheet"],
                        period["label"],
                        rate_row["label"],
                    )
                )

        if results:
            st.session_state.df_master = pd.DataFrame(results)
            st.success(
                f"Conferência concluída: {len(io_configs)} relatório(s) processado(s)."
            )
        else:
            st.session_state.df_master = None
            st.error("Não encontrei preços que pudessem ser comparados.")


if st.session_state.df_master is not None:
    df_master = st.session_state.df_master.copy()
    if "Tolerância (R$)" not in df_master.columns:
        df_master["Tolerância (R$)"] = df_master["Valor Base (Carmel)"].map(
            lambda value: (
                min(max(5.0, value * 0.001), 15.0)
                if pd.notna(value) and value > 0
                else None
            )
        )
    status_names = {
        "Correto": "OK",
        "Correto (Diferença tolerada)": "Diferença tolerada",
        "Inconsistência de Valor": "Preço diferente",
        "Valor Zerado no IO": "Preço zerado",
        "Faltando na Base": "Não encontrado na base",
        "Faltando no IO": "Não encontrado no IO",
    }

    def friendly_io_period(label):
        parts = [clean_text(part) for part in str(label).split("|")]
        if len(parts) < 3:
            return clean_text(label)
        return parts[1] if parts[1] == parts[2] else f"{parts[1]} a {parts[2]}"

    df_master["Data IO"] = df_master["Linha IO"].map(friendly_io_period)
    df_master["Resultado"] = df_master["Status"].map(status_names).fillna(
        df_master["Status"]
    )
    df_master["Problema"] = ~df_master["Status"].str.startswith("Correto")
    df_master["Tolerado"] = df_master["Status"].eq(
        "Correto (Diferença tolerada)"
    )
    problems = df_master[df_master["Problema"]]
    correct = df_master[~df_master["Problema"]]
    tolerated = df_master[df_master["Tolerado"]]
    exact = df_master[df_master["Status"].eq("Correto")]

    def visual_result(value):
        if value == "OK":
            return "✅ Correto"
        if value == "Diferença tolerada":
            return "🟡 Diferença tolerada"
        return f"🔴 {value}"

    def color_result_rows(row, status_column="Situação"):
        value = str(row.get(status_column, ""))
        if "✅" in value or value in {"OK", "Correto"}:
            color = "background-color: rgba(22, 163, 74, 0.18)"
        elif "🟡" in value or "tolerada" in value.lower():
            color = "background-color: rgba(234, 179, 8, 0.20)"
        elif "⚠️" in value or "🔴" in value or "problema" in value.lower():
            color = "background-color: rgba(220, 38, 38, 0.18)"
        else:
            color = ""
        return [color] * len(row)

    st.divider()
    st.markdown("### Resultado geral")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Relatórios conferidos", df_master["Arquivo IO"].nunique())
    c2.metric("Períodos analisados", df_master["Linha IO"].nunique())
    c3.metric("Preços aprovados", len(correct))
    c4.metric("Preços com problema", len(problems))

    if problems.empty:
        st.success("Tudo certo. Nenhuma divergência foi encontrada.")
    else:
        combinations = problems[
            ["Arquivo IO", "Categoria", "Acomodação", "Resultado"]
        ].drop_duplicates()
        st.warning(
            f"Foram encontradas {len(problems)} ocorrências em "
            f"{len(combinations)} tarifa(s). Veja abaixo o que precisa de atenção."
        )

    st.markdown("#### Visão dos resultados")
    chart_col1, chart_col2 = st.columns(2, gap="large")
    result_chart = pd.DataFrame(
        {
            "Resultado": ["Iguais", "Dentro da tolerância", "Com problema"],
            "Quantidade": [len(exact), len(tolerated), len(problems)],
        }
    )
    with chart_col1:
        result_figure = px.pie(
            result_chart,
            names="Resultado",
            values="Quantidade",
            hole=0.66,
            color="Resultado",
            color_discrete_map={
                "Iguais": "#22C55E",
                "Dentro da tolerância": "#FACC15",
                "Com problema": "#EF4444",
            },
        )
        result_figure.update_traces(
            textinfo="percent+value",
            textposition="inside",
            marker=dict(line=dict(color="rgba(255,255,255,0.75)", width=2)),
            hovertemplate="<b>%{label}</b><br>%{value} preços<br>%{percent}<extra></extra>",
        )
        result_figure.add_annotation(
            text=f"<b>{len(df_master)}</b><br>preços",
            x=0.5,
            y=0.5,
            showarrow=False,
            font=dict(size=20),
        )
        result_figure.update_layout(
            title=dict(text="Resumo da conferência", x=0.03, xanchor="left"),
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.05,
                xanchor="center",
                x=0.5,
                title_text="",
            ),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=55, b=55, l=20, r=20),
            height=390,
        )
        st.plotly_chart(result_figure, width="stretch", config={"displayModeBar": False})

    with chart_col2:
        if problems.empty:
            approved_chart = pd.DataFrame(
                {
                    "Resultado": ["Iguais", "Dentro da tolerância"],
                    "Quantidade": [len(exact), len(tolerated)],
                }
            )
            approved_figure = px.bar(
                approved_chart,
                x="Resultado",
                y="Quantidade",
                text="Quantidade",
                color="Resultado",
                color_discrete_map={
                    "Iguais": "#22C55E",
                    "Dentro da tolerância": "#FACC15",
                },
            )
            approved_figure.update_layout(
                title=dict(text="Preços aprovados", x=0.03, xanchor="left"),
                showlegend=False,
                xaxis_title="",
                yaxis_title="Quantidade",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                margin=dict(t=55, b=45, l=20, r=20),
                height=390,
            )
            approved_figure.update_traces(
                textposition="outside",
                marker_line_width=0,
                hovertemplate="<b>%{x}</b><br>%{y} preços<extra></extra>",
            )
            st.plotly_chart(
                approved_figure, width="stretch", config={"displayModeBar": False}
            )
        else:
            problem_chart = (
                problems["Resultado"]
                .value_counts()
                .rename_axis("Problema")
                .reset_index(name="Quantidade")
                .sort_values("Quantidade")
            )
            problem_figure = px.bar(
                problem_chart,
                x="Quantidade",
                y="Problema",
                orientation="h",
                text="Quantidade",
                color="Problema",
                color_discrete_sequence=["#EF4444", "#F97316", "#DC2626"],
            )
            problem_figure.update_traces(
                textposition="outside",
                cliponaxis=False,
                marker_line_width=0,
                hovertemplate="<b>%{y}</b><br>%{x} ocorrências<extra></extra>",
            )
            problem_figure.update_layout(
                title=dict(text="Problemas encontrados", x=0.03, xanchor="left"),
                showlegend=False,
                xaxis_title="Quantidade",
                yaxis_title="",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                margin=dict(t=55, b=45, l=20, r=40),
                height=390,
            )
            st.plotly_chart(
                problem_figure, width="stretch", config={"displayModeBar": False}
            )

    st.markdown("#### Tabela principal")
    st.caption(
        "Veja os preços mais importantes da comparação. Use os filtros para "
        "mostrar somente o que precisa de atenção ou consultar todos os resultados."
    )
    filter_col1, filter_col2 = st.columns([1, 1])
    default_view = "Somente problemas" if not problems.empty else "Todos"
    with filter_col1:
        main_filter = st.radio(
            "O que deseja visualizar?",
            ["Somente problemas", "Todos", "Somente corretos"],
            index=["Somente problemas", "Todos", "Somente corretos"].index(
                default_view
            ),
            horizontal=True,
            key="main_result_filter",
        )
    with filter_col2:
        report_options = ["Todos os relatórios"] + sorted(
            df_master["Arquivo IO"].unique()
        )
        main_report = st.selectbox(
            "Qual relatório?",
            report_options,
            key="main_report_filter",
        )

    if main_filter == "Somente problemas":
        main_source = problems.copy()
    elif main_filter == "Somente corretos":
        main_source = correct.copy()
    else:
        main_source = df_master.copy()
    if main_report != "Todos os relatórios":
        main_source = main_source[main_source["Arquivo IO"] == main_report]

    main_view = main_source[
        [
            "Resultado",
            "Data IO",
            "Flutuação",
            "Categoria",
            "Acomodação",
            "Valor Base (Carmel)",
            "Valor Envio (IO)",
            "Tolerância (R$)",
            "Diferença (R$)",
            "Diferença (%)",
        ]
    ].rename(
        columns={
            "Resultado": "Situação",
            "Data IO": "Data",
            "Flutuação": "Faixa de preço",
            "Categoria": "Quarto",
            "Acomodação": "Ocupação",
            "Valor Base (Carmel)": "Preço oficial",
            "Valor Envio (IO)": "Preço encontrado",
            "Tolerância (R$)": "Limite aceito",
        }
    )
    main_view["Situação"] = main_view["Situação"].map(visual_result)
    main_styled = main_view.style.apply(
        color_result_rows, axis=1, status_column="Situação"
    )
    st.dataframe(
        main_styled,
        width="stretch",
        hide_index=True,
        column_config={
            "Preço oficial": st.column_config.NumberColumn(format="R$ %.2f"),
            "Preço encontrado": st.column_config.NumberColumn(format="R$ %.2f"),
            "Limite aceito": st.column_config.NumberColumn(format="R$ %.2f"),
            "Diferença (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
            "Diferença (%)": st.column_config.NumberColumn(format="%.2f%%"),
        },
    )
    st.caption(f"{len(main_view)} preço(s) exibido(s).")

    additional_view = df_master[
        [
            "Arquivo IO",
            "Data IO",
            "Flutuação",
            "Período Base",
            "Categoria",
            "Acomodação",
            "Valor Base (Carmel)",
            "Valor Envio (IO)",
            "Tolerância (R$)",
            "Diferença (R$)",
            "Diferença (%)",
            "Resultado",
        ]
    ].rename(
        columns={
            "Arquivo IO": "Relatório",
            "Data IO": "Data",
            "Flutuação": "Faixa de preço",
            "Período Base": "Período da tabela oficial",
            "Categoria": "Quarto",
            "Acomodação": "Ocupação",
            "Valor Base (Carmel)": "Preço da tabela oficial",
            "Valor Envio (IO)": "Preço encontrado",
            "Tolerância (R$)": "Limite aceito",
            "Diferença (%)": "Diferença (%)",
            "Resultado": "Situação",
        }
    )
    additional_view["Situação"] = additional_view["Situação"].map(visual_result)

    st.divider()
    st.download_button(
        label="⬇️ Baixar relatório completo (.xlsx)",
        data=export_to_excel(additional_view),
        file_name="Auditoria_Tarifas_Carmel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
