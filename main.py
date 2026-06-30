import re
import unicodedata
from datetime import datetime
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st


CAPACITIES = {"SGL", "DPL", "TPL", "QDP", "QTP", "SEX"}


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


def normalize_category(value):
    value = remove_accents(clean_text(value).upper()).replace("NORDESTINA ", "")
    return {
        "DO MAR": "MAR",
        "PASS LOFT": "PASSARO LOFT",
    }.get(value, value)


def normalize_capacity(value):
    value = remove_accents(clean_text(value).upper()).replace(" ", "")
    value = {"DBL": "DPL", "STP": "SEX"}.get(value, value)
    return value.replace("CR", "CHD", 1) if value.startswith("CR") else value


def is_capacity(value):
    value = normalize_capacity(value)
    return value in CAPACITIES or value.startswith("CHD")


def parse_money(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = clean_text(value).replace("R$", "").replace(" ", "")
    if not text or text.upper() in {"NAN", "NONE"}:
        return None
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def extract_dates(value):
    text = str(value)
    found = []
    occupied = []

    ranges = re.compile(
        r"\b(\d{1,2})(?:/(\d{1,2}))?\s*(?:a|até|ate|-)\s*"
        r"(\d{1,2})/(\d{1,2})/(\d{2,4})\b",
        re.IGNORECASE,
    )
    for match in ranges.finditer(text):
        start_day, start_month, end_day, end_month, year = match.groups()
        year = f"20{year}" if len(year) == 2 else year
        try:
            start = datetime(
                int(year), int(start_month or end_month), int(start_day)
            ).date()
            end = datetime(int(year), int(end_month), int(end_day)).date()
        except ValueError:
            continue
        found.extend([(match.start(), 0, start), (match.start(), 1, end)])
        occupied.append(match.span())

    def overlaps(span):
        return any(span[0] < end and span[1] > start for start, end in occupied)

    for match in re.finditer(r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b", text):
        if overlaps(match.span()):
            continue
        day, month, year = match.groups()
        year = f"20{year}" if len(year) == 2 else year
        try:
            found.append(
                (match.start(), 0, datetime(int(year), int(month), int(day)).date())
            )
        except ValueError:
            pass

    for match in re.finditer(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", text):
        year, month, day = match.groups()
        try:
            found.append(
                (match.start(), 0, datetime(int(year), int(month), int(day)).date())
            )
        except ValueError:
            pass
    return [date for _, _, date in sorted(found)]


def extract_years(value):
    return {int(year) for year in re.findall(r"\b(20\d{2})\b", str(value))}


def rate_years(rates):
    return {date.year for rate in rates for date in extract_dates(rate["label"])}


def infer_period_index(label, periods):
    dates = extract_dates(label)
    if not dates:
        return None
    target = dates[0]
    matches = []
    for index, period in enumerate(periods):
        period_dates = extract_dates(period["label"])
        intervals = list(zip(period_dates[::2], period_dates[1::2]))
        if len(period_dates) % 2:
            intervals.append((period_dates[-1], period_dates[-1]))
        for start, end in intervals:
            if start <= target <= end:
                matches.append(((end - start).days, index))
    return min(matches)[1] if matches else None


def reset_file(file):
    if hasattr(file, "seek"):
        file.seek(0)


def list_sheets(file):
    if file.name.lower().endswith(".csv"):
        return ["CSV"]
    reset_file(file)
    engine = "calamine" if file.name.lower().endswith(".xls") else None
    try:
        sheets = pd.ExcelFile(file, engine=engine).sheet_names
    except Exception:
        reset_file(file)
        sheets = pd.ExcelFile(file, engine="calamine").sheet_names
    reset_file(file)
    return sheets


def load_file(file, sheet_name=None):
    reset_file(file)
    try:
        if file.name.lower().endswith(".csv"):
            return pd.read_csv(file, header=None, sep=None, engine="python")
        engine = "calamine" if file.name.lower().endswith(".xls") else None
        try:
            return pd.read_excel(
                file, sheet_name=sheet_name, header=None, engine=engine
            )
        except Exception:
            reset_file(file)
            return pd.read_excel(
                file, sheet_name=sheet_name, header=None, engine="calamine"
            )
    finally:
        reset_file(file)


def find_header_row(df):
    for row in range(min(30, len(df))):
        if any("ARVORE" in normalize_category(value) for value in df.iloc[row]):
            return row
    return None


def find_base_groups(df):
    header = find_header_row(df)
    if header is None:
        return []

    candidates = []
    for column in range(len(df.columns) - 1):
        matches = sum(
            is_capacity(df.iat[row, column + 1])
            for row in range(header + 1, len(df))
        )
        if matches >= 2:
            candidates.append(column)

    period_columns = []
    for column in candidates:
        if not period_columns or column > period_columns[-1] + 1:
            period_columns.append(column)

    groups = []
    for index, period_column in enumerate(period_columns):
        end = (
            period_columns[index + 1]
            if index + 1 < len(period_columns)
            else len(df.columns)
        )
        categories = [
            (column, normalize_category(df.iat[header, column]))
            for column in range(period_column + 2, end)
            if clean_text(df.iat[header, column])
        ]
        if categories:
            groups.append(
                {
                    "name": clean_text(df.iat[header, period_column])
                    or f"Bloco {index + 1}",
                    "header": header,
                    "period_column": period_column,
                    "capacity_column": period_column + 1,
                    "categories": categories,
                }
            )
    return groups


def extract_base_periods(df, group):
    return [
        {"label": clean_text(df.iat[row, group["period_column"]]), "row": row}
        for row in range(group["header"] + 1, len(df))
        if clean_text(df.iat[row, group["period_column"]])
        and is_capacity(df.iat[row, group["capacity_column"]])
    ]


def extract_base_prices(df, group, period):
    prices = {}
    row = period["row"]
    while row < len(df) and is_capacity(df.iat[row, group["capacity_column"]]):
        if row > period["row"] and clean_text(df.iat[row, group["period_column"]]):
            break
        capacity = normalize_capacity(df.iat[row, group["capacity_column"]])
        for column, category in group["categories"]:
            value = parse_money(df.iat[row, column])
            if value is not None:
                prices[(category, capacity)] = value
        row += 1
    return prices


def detect_io_columns(row):
    columns = []
    for index, value in enumerate(row):
        name = clean_text(value).upper()
        if "-" not in name:
            continue
        category, capacity = name.rsplit("-", 1)
        if category and is_capacity(capacity):
            columns.append(
                (index, normalize_category(category), normalize_capacity(capacity))
            )
    return columns


def extract_io_rate_rows(df):
    rows = []
    columns = []
    context = ""
    for row_index in range(len(df)):
        detected = detect_io_columns(df.iloc[row_index])
        if detected:
            columns = detected
            previous = df.iloc[row_index - 1, :4] if row_index else []
            context = " | ".join(clean_text(value) for value in previous if clean_text(value))
            continue
        if not columns:
            continue
        first_cells = remove_accents(
            " | ".join(clean_text(value).upper() for value in df.iloc[row_index, :5])
        )
        if "MELHOR PRECO" not in first_cells and "MELHOR TARIFA" not in first_cells:
            continue
        name = clean_text(df.iat[row_index, 0]) or f"Linha {row_index + 1}"
        rows.append(
            {
                "label": " | ".join(part for part in (context, name) if part),
                "name": name,
                "row": row_index,
                "columns": columns.copy(),
            }
        )
    return rows


def extract_io_prices(df, rate):
    prices = {}
    expected = set()
    for column, category, capacity in rate["columns"]:
        key = (category, capacity)
        expected.add(key)
        value = parse_money(df.iat[rate["row"], column])
        if value is not None:
            prices[key] = value
    return prices, expected


def compare_prices(
    base_prices,
    io_prices,
    expected,
    source,
    sheet,
    io_period,
    base_period,
):
    results = []
    for category, capacity in sorted(set(base_prices) | set(io_prices)):
        key = (category, capacity)
        base_value = base_prices.get(key)
        io_value = io_prices.get(key)
        if io_value is None and key not in expected:
            continue
        if base_value is None:
            status = "Faltando na Base"
            classification = "Grande diferença"
        elif io_value is None:
            status = "Faltando no IO"
            classification = "Grande diferença"
        elif io_value <= 0 < base_value:
            status = "Valor Zerado no IO"
            classification = "Grande diferença"
        elif abs(base_value - io_value) <= 0.01:
            status = "Igual"
            classification = "Igual"
        elif abs(base_value - io_value) <= 10:
            status = "Correto (ajuste de até R$ 10)"
            classification = "Próximo"
        else:
            status = "Preço Diferente"
            classification = "Grande diferença"

        difference = (
            io_value - base_value
            if base_value is not None and io_value is not None
            else None
        )
        results.append(
            {
                "Arquivo IO": source,
                "Aba IO": sheet,
                "Período IO": io_period,
                "Período Base Usado": base_period,
                "Categoria": category,
                "Ocupação": capacity,
                "Valor Base": base_value,
                "Valor IO": io_value,
                "Diferença": round(difference, 2) if difference is not None else None,
                "Classificação": classification,
                "Status": status,
            }
        )
    return results


def export_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Comparação")
    return output.getvalue()


def color_result_rows(row):
    classification = str(row.get("Classificação", ""))
    if "Igual" in classification:
        color = "background-color: rgba(16, 185, 129, 0.18)"
    elif "Próximo" in classification:
        color = "background-color: rgba(245, 158, 11, 0.20)"
    else:
        color = "background-color: rgba(239, 68, 68, 0.18)"
    return [color] * len(row)


def run_app():
    st.set_page_config(
        page_title="Auditoria de Tarifas Carmel", page_icon="🏨", layout="wide"
    )
    st.markdown(
        """
        <style>
        .main .block-container { padding-top: 2rem; padding-bottom: 2rem; }
        h1 { color: #1E3A8A; font-weight: 700; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("🏨 Painel de Auditoria de Tarifas")
    st.caption("Compare a tabela oficial com a IO de forma visual e categorizada.")

    with st.container(border=True):
        left, right = st.columns(2)
        base_file = left.file_uploader(
            "📂 Tabela oficial", type=["xls", "xlsx", "csv"]
        )
        io_files = right.file_uploader(
            "📂 Relatórios IO",
            type=["xls", "xlsx", "csv"],
            accept_multiple_files=True,
        )
    if not base_file:
        return

    try:
        base_sheet = st.selectbox("Aba da tabela oficial", list_sheets(base_file))
        base_df = load_file(base_file, base_sheet)
        groups = find_base_groups(base_df)
        if not groups:
            st.error("Não encontrei os blocos de tarifas na tabela oficial.")
            return
        group_name = st.selectbox("Bloco de preços", [group["name"] for group in groups])
        group = next(group for group in groups if group["name"] == group_name)
        periods = extract_base_periods(base_df, group)
    except Exception as error:
        st.error(f"Erro ao abrir a tabela oficial: {error}")
        return

    reports = []
    for index, file in enumerate(io_files or []):
        try:
            sheet = st.selectbox(
                f"Aba de {file.name}", list_sheets(file), key=f"io_sheet_{index}"
            )
            df = load_file(file, sheet)
            rates = extract_io_rate_rows(df)
            preferred = [
                rate
                for rate in rates
                if "MELHOR PRECO DISPONIVEL"
                in remove_accents(rate["name"].upper())
            ]
            if not preferred:
                names = sorted({rate["name"] for rate in rates})
                st.warning(
                    f"{file.name}: não encontrei 'Melhor Preço Disponível'. "
                    f"Linhas encontradas: {', '.join(names) or 'nenhuma'}. "
                    "O relatório não será comparado com a tarifa base."
                )
                continue

            base_categories = {category for _, category in group["categories"]}
            io_categories = {
                category
                for rate in preferred
                for _, category, _ in rate["columns"]
            }
            if not base_categories.intersection(io_categories):
                st.warning(
                    f"{file.name}: as categorias não correspondem à base. "
                    f"IO: {', '.join(sorted(io_categories))}. "
                    "O relatório foi ignorado para evitar falsos erros."
                )
                continue

            years_in_name = extract_years(file.name)
            years_in_rates = rate_years(preferred)
            if years_in_name and years_in_rates and not years_in_name.intersection(years_in_rates):
                st.warning(
                    f"{file.name}: o ano no nome do arquivo "
                    f"({', '.join(map(str, sorted(years_in_name)))}) não bate "
                    f"com as datas internas "
                    f"({', '.join(map(str, sorted(years_in_rates)))}). "
                    "Confira a planilha antes de confiar na comparação."
                )

            reports.append((file.name, sheet, df, preferred))
        except Exception as error:
            st.error(f"Erro ao abrir {file.name}: {error}")

    if st.button("Comparar", type="primary", disabled=not reports):
        results = []
        for name, sheet, io_df, rates in reports:
            for rate in rates:
                period_index = infer_period_index(rate["label"], periods)
                if period_index is None:
                    st.warning(f"Período não encontrado: {rate['label']}")
                    continue
                period = periods[period_index]
                base_prices = extract_base_prices(base_df, group, period)
                io_prices, expected = extract_io_prices(io_df, rate)
                results.extend(
                    compare_prices(
                        base_prices,
                        io_prices,
                        expected,
                        name,
                        sheet,
                        rate["label"],
                        period["label"],
                    )
                )
        st.session_state["results"] = pd.DataFrame(results)

    result = st.session_state.get("results")
    if result is None or result.empty:
        return

    st.divider()
    st.subheader("📊 Visão geral da auditoria")
    equal = result["Classificação"].eq("Igual").sum()
    close = result["Classificação"].eq("Próximo").sum()
    different = result["Classificação"].eq("Grande diferença").sum()
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Iguais", equal)
    col2.metric("Próximos", close)
    col3.metric("Grande diferença", different)
    col4.metric("Total auditado", len(result))
    st.caption(
        "Diferenças de até R$ 10 são consideradas corretas, mas permanecem "
        "sinalizadas como próximas."
    )

    colors = {
        "Igual": "#10B981",
        "Próximo": "#F59E0B",
        "Grande diferença": "#EF4444",
    }
    chart_left, chart_right = st.columns(2)
    counts = result["Classificação"].value_counts().reset_index()
    counts.columns = ["Classificação", "Quantidade"]
    with chart_left:
        pie = px.pie(
            counts,
            names="Classificação",
            values="Quantidade",
            hole=0.5,
            color="Classificação",
            color_discrete_map=colors,
            title="Distribuição dos resultados",
        )
        pie.update_traces(textinfo="percent+label", showlegend=False)
        pie.update_layout(margin=dict(t=50, b=10, l=10, r=10))
        st.plotly_chart(pie, width="stretch", config={"displayModeBar": False})

    with chart_right:
        large = result[result["Classificação"] == "Grande diferença"]
        if large.empty:
            st.success("Nenhuma diferença grande encontrada.")
        else:
            by_category = (
                large["Categoria"]
                .value_counts()
                .rename_axis("Categoria")
                .reset_index(name="Quantidade")
            )
            bar = px.bar(
                by_category,
                x="Categoria",
                y="Quantidade",
                text="Quantidade",
                color_discrete_sequence=["#EF4444"],
                title="Grandes diferenças por categoria",
            )
            bar.update_layout(showlegend=False, xaxis_title="", yaxis_title="Quantidade")
            st.plotly_chart(bar, width="stretch", config={"displayModeBar": False})

    st.divider()
    selected = st.radio(
        "Filtre a visão da tabela:",
        ["Todos", "Iguais", "Próximos", "Grande diferença"],
        horizontal=True,
    )
    labels = {
        "Iguais": "Igual",
        "Próximos": "Próximo",
        "Grande diferença": "Grande diferença",
    }
    view = (
        result
        if selected == "Todos"
        else result[result["Classificação"] == labels[selected]]
    )

    st.subheader(f"📋 Detalhamento — {selected}")
    if view.empty:
        st.info("Nenhum item encontrado para este filtro.")
    else:
        for category in sorted(view["Categoria"].unique()):
            category_view = view[view["Categoria"] == category]
            with st.expander(
                f"🏢 {category} — {len(category_view)} item(ns)",
                expanded=selected != "Todos",
            ):
                display = category_view[
                    [
                        "Ocupação",
                        "Período IO",
                        "Período Base Usado",
                        "Valor Base",
                        "Valor IO",
                        "Diferença",
                        "Classificação",
                        "Status",
                    ]
                ].copy()
                display["Classificação"] = display["Classificação"].map(
                    {
                        "Igual": "🟢 Igual",
                        "Próximo": "🟡 Próximo",
                        "Grande diferença": "🔴 Grande diferença",
                    }
                )
                st.dataframe(
                    display.style.apply(color_result_rows, axis=1),
                    hide_index=True,
                    width="stretch",
                    column_config={
                        "Valor Base": st.column_config.NumberColumn(format="R$ %.2f"),
                        "Valor IO": st.column_config.NumberColumn(format="R$ %.2f"),
                        "Diferença": st.column_config.NumberColumn(format="R$ %.2f"),
                    },
                )

    st.divider()
    st.subheader("💾 Exportar resultados")
    st.download_button(
        "⬇️ Baixar planilha (.xlsx)",
        export_excel(view),
        "Auditoria_Tarifas_Carmel.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


if __name__ == "__main__":
    run_app()
