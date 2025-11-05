# app.py
# -*- coding: utf-8 -*-
# Интерактивный дэшборд (Streamlit) для визуализации План/Факт.
# Запуск:
#   pip install streamlit pandas plotly openpyxl xlsxwriter
#   streamlit run app.py

import re
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ----------------------- УТИЛИТЫ -----------------------

def _coerce_number(series: pd.Series) -> pd.Series:
    """
    Приведение к числу: чистим пробелы/NBSP, табы, заменяем запятую на точку.
    Дополнительно выбрасываем любые посторонние символы (кроме цифр, - и .).
    """
    if series.dtype.kind in ("i", "f"):
        return series.astype(float)
    s = series.astype(str)
    # нормализуем
    s = (
        s.str.replace("\u00a0", "", regex=False)  # NBSP
         .str.replace(" ", "", regex=False)
         .str.replace("\t", "", regex=False)
         .str.replace(",", ".", regex=False)
    )
    # оставляем только [-0-9.] (защита от «1.234,5$», «~», и т.п.)
    s = s.str.replace(r"[^0-9\.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")

def find_year_columns(columns: List[str]) -> Tuple[Dict[int, str], Dict[int, str]]:
    fact_pattern = re.compile(r"^Факт in mark\s*(\d{4}),\s*\$")
    plan_pattern = re.compile(r"^План in mark\s*(\d{4}),\s*\$")
    fact_cols, plan_cols = {}, {}
    for c in columns:
        if not isinstance(c, str):
            continue
        c2 = c.strip()
        mf = fact_pattern.match(c2)
        mp = plan_pattern.match(c2)
        if mf:
            fact_cols[int(mf.group(1))] = c
        if mp:
            plan_cols[int(mp.group(1))] = c
    return fact_cols, plan_cols

def build_tidy(df: pd.DataFrame) -> pd.DataFrame:
    """Собирает длинную таблицу: Год / План,$ / Факт,$ / (Продукт/Дивизион)."""
    df = df.copy()
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    fact_cols, plan_cols = find_year_columns(list(df.columns))
    years = sorted(set(fact_cols.keys()) | set(plan_cols.keys()))
    id_cols = [c for c in ("Продукт", "Дивизион") if c in df.columns]

    records = []
    for _, row in df.iterrows():
        for y in years:
            plan_val = row.get(plan_cols.get(y))
            fact_val = row.get(fact_cols.get(y))
            if pd.notna(plan_val) or pd.notna(fact_val):
                rec = {"Год": int(y), "План,$": plan_val, "Факт,$": fact_val}
                for ic in id_cols:
                    rec[ic] = row.get(ic, None)
                records.append(rec)

    tidy = pd.DataFrame(records)
    for c in ["План,$", "Факт,$"]:
        tidy[c] = _coerce_number(tidy[c])
    # выбросим строки, где обе метрики пустые или обе == 0
    tidy = tidy.dropna(how="all", subset=["План,$", "Факт,$"])
    return tidy

def percent(numerator: float, denominator: float) -> Optional[float]:
    if denominator and denominator != 0 and pd.notna(numerator) and pd.notna(denominator):
        return numerator / denominator * 100.0
    return None

def to_excel_download(df_dict: Dict[str, pd.DataFrame]) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for name, df in df_dict.items():
            sheet = name[:31] if len(name) > 31 else name
            df.to_excel(writer, sheet_name=sheet, index=False)
    buf.seek(0)
    return buf.read()

def pick_last_year_with_data(df: pd.DataFrame, years_range: Tuple[int, int]) -> Optional[int]:
    """Возвращает максимальный год в диапазоне, по которому есть ненулевой Факт после фильтров."""
    cand = (df.groupby("Год", as_index=False)["Факт,$"].sum()
              .query("@years_range[0] <= Год <= @years_range[1]"))
    if cand.empty:
        return None
    cand = cand[cand["Факт,$"].fillna(0) != 0]
    if cand.empty:
        return None
    return int(cand["Год"].max())


# ----------------------- UI -----------------------

st.set_page_config(page_title="A&P Dashboard", layout="wide")
st.title("📊 A&P Sales Dashboard")

with st.sidebar:
    st.header("Загрузка файла")
    f = st.file_uploader("Загрузи Excel (.xlsx)", type=["xlsx"])
    selected_sheet = None
    file_bytes = None
    sheets: List[str] = []

    if f is not None:
        try:
            file_bytes = f.read()
            xls_tmp = pd.ExcelFile(BytesIO(file_bytes))
            sheets = xls_tmp.sheet_names
            selected_sheet = st.selectbox("Лист:", sheets, index=0)
        except Exception as e:
            st.error(f"Ошибка чтения файла: {e}")

    st.markdown("---")
    show_debug = st.checkbox("Показать диагностику", value=False)
    st.caption("Ожидаются колонки «План in mark {год}, $» и «Факт in mark {год}, $». "
               "Опционально — «Продукт», «Дивизион».")

if file_bytes is None or selected_sheet is None:
    st.info("Загрузи файл слева, выбери лист — и дэшборд отрисуется.")
    st.stop()

# Читаем выбранный лист
try:
    df_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=selected_sheet)
except Exception as e:
    st.error(f"Не удалось прочитать лист «{selected_sheet}»: {e}")
    st.stop()

if df_raw.empty:
    st.warning("Выбранный лист пуст.")
    st.stop()

# Tidy
tidy = build_tidy(df_raw)
if tidy.empty:
    st.warning("Не найдено пар колонок План/Факт по годам. Проверь названия.")
    st.stop()

# ----------------------- ФИЛЬТРЫ -----------------------

years = sorted(tidy["Год"].dropna().unique().tolist())
min_year, max_year = years[0], years[-1]

col_y1, col_y2 = st.columns([1, 2])
with col_y1:
    year_range = st.slider("Годы", min_value=int(min_year), max_value=int(max_year),
                           value=(int(min_year), int(max_year)), step=1)

divisions = sorted(tidy["Дивизион"].dropna().unique().tolist()) if "Дивизион" in tidy.columns else []
products = sorted(tidy["Продукт"].dropna().unique().tolist()) if "Продукт" in tidy.columns else []

with col_y2:
    filt_cols = st.columns(2)
    with filt_cols[0]:
        # ВАЖНО: по умолчанию выбираем ВСЕ дивизионы, а не первые 10
        default_divs = divisions  # все
        sel_divs = st.multiselect("Дивизион", options=divisions, default=default_divs)
    with filt_cols[1]:
        product_search = st.text_input("Поиск по продукту (подстрока)", value="")

# Применяем фильтры
mask = tidy["Год"].between(year_range[0], year_range[1])
if divisions and sel_divs:
    mask &= tidy["Дивизион"].isin(sel_divs)
if product_search.strip() and "Продукт" in tidy.columns:
    sub = product_search.strip().lower()
    mask &= tidy["Продукт"].fillna("").str.lower().str.contains(sub)

tidy_f = tidy.loc[mask].copy()
if tidy_f.empty:
    st.warning("По текущим фильтрам нет данных. Убери лишние фильтры или расширь диапазон лет.")
    st.stop()

if show_debug:
    st.info(
        f"Строк после фильтров: {len(tidy_f):,}. "
        f"Года с данными: {sorted(tidy_f['Год'].unique().tolist())}"
    )

# ----------------------- KPI -----------------------

kpi = tidy_f.groupby("Год", as_index=False)[["План,$", "Факт,$"]].sum().sort_values("Год")
total_plan = float(kpi["План,$"].sum())
total_fact = float(kpi["Факт,$"].sum())
total_perf = percent(total_fact, total_plan)

col1, col2, col3, col4 = st.columns(4)
col1.metric("План (∑), $", f"{total_plan:,.0f}")
col2.metric("Факт (∑), $", f"{total_fact:,.0f}")
col3.metric("% выполнения (∑)", f"{total_perf:.1f}%" if total_perf is not None else "—")
col4.metric("Отклонение, $", f"{(total_fact - total_plan):,.0f}")

# ----------------------- ГРАФИКИ -----------------------

# 1) Линия План/Факт по годам (итого)
line_fig = go.Figure()
line_fig.add_trace(go.Scatter(x=kpi["Год"], y=kpi["План,$"], mode="lines+markers", name="План, $"))
line_fig.add_trace(go.Scatter(x=kpi["Год"], y=kpi["Факт,$"], mode="lines+markers", name="Факт, $"))
line_fig.update_layout(
    title="План vs Факт по годам (итого, с учётом фильтров)",
    xaxis_title="Год", yaxis_title="$",
    hovermode="x unified", height=420, margin=dict(l=40, r=30, t=60, b=40)
)
st.plotly_chart(line_fig, use_container_width=True)

# Подбираем год для ТОПов и долей: последний год в диапазоне, где есть Факт
auto_top_year = pick_last_year_with_data(tidy_f, year_range)
if auto_top_year is None:
    st.warning("В выбранном диапазоне лет нет данных для ТОП-10/дивизионов.")
else:
    top_year = auto_top_year

    top_block_cols = st.columns(2)

    # 2) ТОП-10 продуктов по Факту за выбранный (автовыбранный) год
    if "Продукт" in tidy_f.columns and not tidy_f[tidy_f["Год"] == top_year].empty:
        top_df = (tidy_f[tidy_f["Год"] == top_year]
                  .groupby("Продукт", as_index=False)["Факт,$"].sum()
                  .sort_values("Факт,$", ascending=False).head(10))
        with top_block_cols[0]:
            bar_fig = px.bar(
                top_df.sort_values("Факт,$"),
                x="Факт,$", y="Продукт", orientation="h",
                title=f"ТОП-10 продуктов по Факту, {top_year}",
            )
            bar_fig.update_layout(height=500, margin=dict(l=10, r=10, t=60, b=20))
            st.plotly_chart(bar_fig, use_container_width=True)
    else:
        with top_block_cols[0]:
            st.info(f"Нет данных по продуктам в {top_year} году.")

    # 3) Факт по дивизионам за выбранный год
    if "Дивизион" in tidy_f.columns and not tidy_f[tidy_f["Год"] == top_year].empty:
        div_df = (tidy_f[tidy_f["Год"] == top_year]
                  .groupby("Дивизион", as_index=False)["Факт,$"].sum()
                  .sort_values("Факт,$", ascending=False))
        with top_block_cols[1]:
            pie_fig = px.pie(div_df, values="Факт,$", names="Дивизион",
                             title=f"Факт по дивизионам, {top_year}", hole=0.35)
            pie_fig.update_layout(height=500, margin=dict(l=10, r=10, t=60, b=20))
            st.plotly_chart(pie_fig, use_container_width=True)
    else:
        with top_block_cols[1]:
            st.info(f"Нет данных по дивизионам в {top_year} году.")

# 4) Тренды
st.markdown("### Детализация трендов")
trend_cols = st.columns(2)
if "Продукт" in tidy_f.columns and len(products) > 0:
    with trend_cols[0]:
        prod_sel = st.selectbox("Продукт (для тренда)", ["—"] + products)
        if prod_sel != "—":
            p_df = (tidy_f[tidy_f["Продукт"] == prod_sel]
                    .groupby("Год", as_index=False)[["План,$", "Факт,$"]].sum()
                    .sort_values("Год"))
            if not p_df.empty:
                pf = go.Figure()
                pf.add_trace(go.Scatter(x=p_df["Год"], y=p_df["План,$"], mode="lines+markers", name="План, $"))
                pf.add_trace(go.Scatter(x=p_df["Год"], y=p_df["Факт,$"], mode="lines+markers", name="Факт, $"))
                pf.update_layout(title=f"Тренд: {prod_sel}", xaxis_title="Год", yaxis_title="$", height=420)
                st.plotly_chart(pf, use_container_width=True)
            else:
                st.info("Нет данных для выбранного продукта.")

if "Дивизион" in tidy_f.columns and len(divisions) > 0:
    with trend_cols[1]:
        div_sel = st.selectbox("Дивизион (для тренда)", ["—"] + divisions)
        if div_sel != "—":
            d_df = (tidy_f[tidy_f["Дивизион"] == div_sel]
                    .groupby("Год", as_index=False)[["План,$", "Факт,$"]].sum()
                    .sort_values("Год"))
            if not d_df.empty:
                df = go.Figure()
                df.add_trace(go.Scatter(x=d_df["Год"], y=d_df["План,$"], mode="lines+markers", name="План, $"))
                df.add_trace(go.Scatter(x=d_df["Год"], y=d_df["Факт,$"], mode="lines+markers", name="Факт, $"))
                df.update_layout(title=f"Тренд по дивизиону: {div_sel}", xaxis_title="Год", yaxis_title="$", height=420)
                st.plotly_chart(df, use_container_width=True)
            else:
                st.info("Нет данных для выбранного дивизиона.")

# ----------------------- ТАБЛИЦЫ -----------------------

st.markdown("### Таблицы")
tab1, tab2, tab3 = st.tabs(["Сырые строки (после фильтров)", "Итоги по годам", "Свод по дивизионам/годам"])

with tab1:
    sort_cols = [c for c in ["Год", "Дивизион", "Продукт"] if c in tidy_f.columns]
    st.dataframe(tidy_f.sort_values(sort_cols, na_position="last"),
                 use_container_width=True, height=420)

with tab2:
    year_summary = tidy_f.groupby("Год", as_index=False)[["План,$", "Факт,$"]].sum()
    year_summary = year_summary.rename(columns={"План,$": "План итого, $",
                                                "Факт,$": "Факт итого, $"})
    year_summary["% выполнения"] = (
        (year_summary["Факт итого, $"] / year_summary["План итого, $"]) * 100.0
    ).round(1)
    st.dataframe(year_summary, use_container_width=True, height=360)

with tab3:
    if "Дивизион" in tidy_f.columns:
        div_year = tidy_f.groupby(["Дивизион", "Год"], as_index=False)[["Факт,$", "План,$"]].sum()
        div_year = div_year.rename(columns={"Факт,$": "Факт, $", "План,$": "План, $"})
        st.dataframe(div_year.sort_values(["Год", "Факт, $"], ascending=[True, False]),
                     use_container_width=True, height=360)
    else:
        st.info("Колонки «Дивизион» нет — этот срез недоступен.")

# ----------------------- ВЫГРУЗКА -----------------------

st.markdown("### Экспорт")
exp_cols = st.columns(3)

with exp_cols[0]:
    csv_bytes = tidy_f.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Скачать отфильтрованные строки (CSV)", data=csv_bytes,
                       file_name="filtered_rows.csv", mime="text/csv")

with exp_cols[1]:
    xls_bytes = to_excel_download({
        "tidy_filtered": tidy_f,
        "year_summary": year_summary if 'year_summary' in locals() else pd.DataFrame(),
        "div_year": div_year if 'div_year' in locals() else pd.DataFrame(),
    })
    st.download_button("⬇️ Скачать своды (Excel)", data=xls_bytes,
                       file_name="dashboard_exports.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with exp_cols[2]:
    st.caption("Готово. При желании добавим кварталы/месяцы, PowerPoint/PDF и пресеты фильтров.")
