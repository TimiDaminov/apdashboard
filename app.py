# app.py
# -*- coding: utf-8 -*-
# Интерактивный дэшборд (Streamlit) для визуализации План/Факт и бюджета A&P с YoY-сравнением.
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
    Приведение к числу: чистим пробелы/NBSP, табы, заменяем запятую на точку,
    оставляем только [-0-9.] (защита от $ и прочих символов).
    """
    if series.dtype.kind in ("i", "f"):
        return series.astype(float)
    s = series.astype(str)
    s = (
        s.str.replace("\u00a0", "", regex=False)  # NBSP
         .str.replace(" ", "", regex=False)
         .str.replace("\t", "", regex=False)
         .str.replace(",", ".", regex=False)
         .str.replace(r"[^0-9\.\-]", "", regex=True)
    )
    return pd.to_numeric(s, errors="coerce")

def find_year_columns_sales(columns: List[str]) -> Tuple[Dict[int, str], Dict[int, str]]:
    """Находим 'План in mark {год}, $' и 'Факт in mark {год}, $'."""
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

def find_year_columns_ap(columns: List[str]) -> Dict[int, str]:
    """
    Находим 'План A&P{год},$' или 'План A&P {год},$' (+ допускаем хвосты вроде '.1').
    Примеры: 'План A&P2025,$', 'План A&P 2026,$', 'План A&P2025,$.1'
    """
    ap_patterns = [
        re.compile(r"^План A&P\s*(\d{4}),\s*\$"),
        re.compile(r"^План A&P\s*(\d{4})")  # запасной вариант, если нет ', $'
    ]
    ap_cols: Dict[int, str] = {}
    for c in columns:
        if not isinstance(c, str):
            continue
        c2 = c.strip()
        for pat in ap_patterns:
            m = pat.match(c2)
            if m:
                year = int(m.group(1))
                # не перетирать уже найденное — берём первый встретившийся столбец
                if year not in ap_cols:
                    ap_cols[year] = c
                break
    return ap_cols

def build_tidy_sales(df: pd.DataFrame) -> pd.DataFrame:
    """Длинная таблица продаж: Год / План,$ / Факт,$ / (Продукт/Дивизион)."""
    df = df.copy()
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    fact_cols, plan_cols = find_year_columns_sales(list(df.columns))
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
    tidy = tidy.dropna(how="all", subset=["План,$", "Факт,$"])
    return tidy

def build_tidy_ap(df: pd.DataFrame) -> pd.DataFrame:
    """Длинная таблица бюджета A&P: Год / A&P план,$ / (Продукт/Дивизион при наличии)."""
    df = df.copy()
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    ap_cols = find_year_columns_ap(list(df.columns))
    years = sorted(ap_cols.keys())
    id_cols = [c for c in ("Продукт", "Дивизион") if c in df.columns]

    records = []
    for _, row in df.iterrows():
        for y in years:
            val = row.get(ap_cols.get(y))
            if pd.notna(val):
                rec = {"Год": int(y), "A&P план,$": val}
                for ic in id_cols:
                    rec[ic] = row.get(ic, None)
                records.append(rec)

    tidy_ap = pd.DataFrame(records) if records else pd.DataFrame(columns=["Год", "A&P план,$"])
    if not tidy_ap.empty:
        tidy_ap["A&P план,$"] = _coerce_number(tidy_ap["A&P план,$"])
    return tidy_ap

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

def pick_last_year_with_data(df: pd.DataFrame, years_range: Tuple[int, int], value_col: str) -> Optional[int]:
    """Последний год в диапазоне, где сумма value_col > 0."""
    if df.empty:
        return None
    cand = (df.groupby("Год", as_index=False)[value_col].sum()
              .query("@years_range[0] <= Год <= @years_range[1]"))
    if cand.empty:
        return None
    cand = cand[cand[value_col].fillna(0) != 0]
    if cand.empty:
        return None
    return int(cand["Год"].max())

def yoy_values(series_by_year: Dict[int, float], year: int) -> Tuple[Optional[float], Optional[float]]:
    """Возвращает (delta_abs, delta_pct) для выбранного года vs предыдущий."""
    prev = year - 1
    cur_v = series_by_year.get(year)
    prev_v = series_by_year.get(prev)
    if cur_v is None or prev_v is None or prev_v == 0:
        return None, None
    delta_abs = cur_v - prev_v
    delta_pct = (delta_abs / prev_v) * 100.0
    return delta_abs, delta_pct


# ----------------------- UI -----------------------

st.set_page_config(page_title="A&P Dashboard", layout="wide")
st.title("📊 A&P Sales Dashboard")

with st.sidebar:
    st.header("Загрузка файла")
    f = st.file_uploader("", type=["xlsx"])
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

if file_bytes is None or selected_sheet is None:
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

# Tidy по продажам и A&P
tidy_sales = build_tidy_sales(df_raw)
tidy_ap = build_tidy_ap(df_raw)  # может быть пустым, если в файле нет столбцов A&P
if tidy_sales.empty:
    st.warning("Не найдено пар колонок План/Факт по годам. Проверь названия.")
    st.stop()

# ----------------------- ФИЛЬТРЫ -----------------------

years = sorted(tidy_sales["Год"].dropna().unique().tolist())
min_year, max_year = years[0], years[-1]

col_y1, col_y2 = st.columns([1, 2])
with col_y1:
    year_range = st.slider("Годы", min_value=int(min_year), max_value=int(max_year),
                           value=(int(min_year), int(max_year)), step=1)

divisions = sorted(tidy_sales["Дивизион"].dropna().unique().tolist()) if "Дивизион" in tidy_sales.columns else []
products = sorted(tidy_sales["Продукт"].dropna().unique().tolist()) if "Продукт" in tidy_sales.columns else []

with col_y2:
    fcols = st.columns(2)
    with fcols[0]:
        sel_divs = st.multiselect("Дивизион", options=divisions, default=divisions)
    with fcols[1]:
        product_search = st.text_input("Поиск по продукту (подстрока)", value="")

# Применяем фильтры к tidy_sales
mask_sales = tidy_sales["Год"].between(year_range[0], year_range[1])
if divisions and sel_divs:
    mask_sales &= tidy_sales["Дивизион"].isin(sel_divs)
if product_search.strip() and "Продукт" in tidy_sales.columns:
    sub = product_search.strip().lower()
    mask_sales &= tidy_sales["Продукт"].fillna("").str.lower().str.contains(sub)

sales_f = tidy_sales.loc[mask_sales].copy()
if sales_f.empty:
    st.warning("По текущим фильтрам нет данных (продажи). Измени фильтры.")
    st.stop()

# Те же фильтры применим к tidy_ap (если он есть)
if not tidy_ap.empty:
    mask_ap = tidy_ap["Год"].between(year_range[0], year_range[1])
    if "Дивизион" in tidy_ap.columns and sel_divs:
        mask_ap &= tidy_ap["Дивизион"].isin(sel_divs)
    if product_search.strip() and "Продукт" in tidy_ap.columns:
        sub = product_search.strip().lower()
        mask_ap &= tidy_ap["Продукт"].fillna("").str.lower().str.contains(sub)
    ap_f = tidy_ap.loc[mask_ap].copy()
else:
    ap_f = pd.DataFrame(columns=["Год", "A&P план,$"])

# ----------------------- KPI -----------------------

kpi = sales_f.groupby("Год", as_index=False)[["План,$", "Факт,$"]].sum().sort_values("Год")
total_plan = float(kpi["План,$"].sum())
total_fact = float(kpi["Факт,$"].sum())
total_perf = percent(total_fact, total_plan)

c1, c2, c3, c4 = st.columns(4)
c1.metric("План (∑), $", f"{total_plan:,.0f}")
c2.metric("Факт (∑), $", f"{total_fact:,.0f}")
c3.metric("% выполнения (∑)", f"{total_perf:.1f}%" if total_perf is not None else "—")
c4.metric("Отклонение, $", f"{(total_fact - total_plan):,.0f}")

# ----------------------- ГРАФИКИ -----------------------

# 1) Линия План/Факт по годам (итого)
line_fig = go.Figure()
line_fig.add_trace(go.Scatter(x=kpi["Год"], y=kpi["План,$"], mode="lines+markers", name="План, $"))
line_fig.add_trace(go.Scatter(x=kpi["Год"], y=kpi["Факт,$"], mode="lines+markers", name="Факт, $"))
line_fig.update_layout(
    title="План vs Факт по годам (итого, с учётом фильтров)",
    xaxis_title="Год", yaxis_title="$",
    hovermode="x unified", height=420, margin=dict(l=40, r=30, t=60, b=40),
)
st.plotly_chart(line_fig, use_container_width=True)

# 2) ТОП-10 и доли по дивизионам (на последний год с данными по Факту)
def pick_last_year(df: pd.DataFrame) -> Optional[int]:
    return pick_last_year_with_data(df, year_range, "Факт,$")

top_year = pick_last_year(sales_f)
if top_year is not None:
    cols_top = st.columns(2)

    if "Продукт" in sales_f.columns and not sales_f[sales_f["Год"] == top_year].empty:
        top_df = (sales_f[sales_f["Год"] == top_year]
                  .groupby("Продукт", as_index=False)["Факт,$"].sum()
                  .sort_values("Факт,$", ascending=False).head(10))
        with cols_top[0]:
            bar_fig = px.bar(
                top_df.sort_values("Факт,$"),
                x="Факт,$", y="Продукт", orientation="h",
                title=f"ТОП-10 продуктов по Факту, {top_year}",
            )
            bar_fig.update_layout(height=500, margin=dict(l=10, r=10, t=60, b=20))
            st.plotly_chart(bar_fig, use_container_width=True)
    else:
        with cols_top[0]:
            st.info(f"Нет данных по продуктам в {top_year}.")

    if "Дивизион" in sales_f.columns and not sales_f[sales_f["Год"] == top_year].empty:
        div_df = (sales_f[sales_f["Год"] == top_year]
                  .groupby("Дивизион", as_index=False)["Факт,$"].sum()
                  .sort_values("Факт,$", ascending=False))
        with cols_top[1]:
            pie_fig = px.pie(div_df, values="Факт,$", names="Дивизион",
                             title=f"Факт по дивизионам, {top_year}", hole=0.35)
            pie_fig.update_layout(height=500, margin=dict(l=10, r=10, t=60, b=20))
            st.plotly_chart(pie_fig, use_container_width=True)

# ----------------------- Y0Y СРАВНЕНИЕ (ПЛАН vs A&P) -----------------------

st.markdown("## 🔁 Сравнение с прошлым годом (YoY)")

# Свод по годам (с учётом фильтров)
sales_plan_by_year = sales_f.groupby("Год", as_index=False)["План,$"].sum()
sales_plan_map = dict(zip(sales_plan_by_year["Год"], sales_plan_by_year["План,$"]))

if not ap_f.empty:
    ap_by_year = ap_f.groupby("Год", as_index=False)["A&P план,$"].sum()
    ap_map = dict(zip(ap_by_year["Год"], ap_by_year["A&P план,$"]))
    ap_years = sorted(ap_map.keys())
else:
    ap_by_year = pd.DataFrame(columns=["Год", "A&P план,$"])
    ap_map = {}
    ap_years = []

# Годы, по которым можно сравнивать (нужны текущий и предыдущий)
candidate_years = sorted(set(sales_plan_map.keys()) | set(ap_map.keys()))
candidate_years = [y for y in candidate_years if (y - 1) in candidate_years]
if not candidate_years:
    st.info("Недостаточно данных для YoY-сравнения (нужен год и предыдущий).")
else:
    # Выбор года сравнения
    default_year = max([y for y in candidate_years if year_range[0] <= y <= year_range[1]], default=candidate_years[-1])
    yoY_year = st.selectbox(
        "Год сравнения (будет сравнен с предыдущим)",
        options=sorted(candidate_years),
        index=sorted(candidate_years).index(default_year)
    )

    # Расчёт YoY для плана продаж
    plan_delta_abs, plan_delta_pct = yoy_values(sales_plan_map, yoY_year)

    # Расчёт YoY для A&P (если есть)
    if ap_map:
        ap_delta_abs, ap_delta_pct = yoy_values(ap_map, yoY_year)
    else:
        ap_delta_abs = ap_delta_pct = None

    m1, m2, m3, m4 = st.columns(4)
    if plan_delta_abs is not None:
        m1.metric("План: прирост, $", f"{plan_delta_abs:,.0f}")
        m2.metric("План: прирост, %", f"{plan_delta_pct:.1f}%")
    else:
        m1.metric("План: прирост, $", "—")
        m2.metric("План: прирост, %", "—")

    if ap_delta_abs is not None:
        m3.metric("A&P: прирост, $", f"{ap_delta_abs:,.0f}")
        m4.metric("A&P: прирост, %", f"{ap_delta_pct:.1f}%")
    else:
        m3.metric("A&P: прирост, $", "—")
        m4.metric("A&P: прирост, %", "—")

    # График 1: прошлый vs текущий (две группы: План и A&P)
    comp_fig = go.Figure()
    xcats = ["План продаж", "Бюджет A&P"]

    prev_vals = [
        sales_plan_map.get(yoY_year - 1, np.nan),
        ap_map.get(yoY_year - 1, np.nan) if ap_map else np.nan
    ]
    curr_vals = [
        sales_plan_map.get(yoY_year, np.nan),
        ap_map.get(yoY_year, np.nan) if ap_map else np.nan
    ]

    comp_fig.add_trace(go.Bar(x=xcats, y=prev_vals, name=f"{yoY_year-1}"))
    comp_fig.add_trace(go.Bar(x=xcats, y=curr_vals, name=f"{yoY_year}"))
    comp_fig.update_layout(
        barmode="group",
        title=f"Сравнение {yoY_year} vs {yoY_year-1}: План и A&P",
        yaxis_title="$",
        height=420,
        margin=dict(l=40, r=30, t=60, b=40)
    )
    st.plotly_chart(comp_fig, use_container_width=True)

    # График 2: %-прирост по двум метрикам
    growth_vals = [
        plan_delta_pct if plan_delta_pct is not None else 0,
        ap_delta_pct if ap_delta_pct is not None else 0
    ]
    growth_fig = px.bar(
        x=xcats, y=growth_vals, labels={"x": "Метрика", "y": "%"},
        title=f"Прирост, % (YoY): {yoY_year} vs {yoY_year-1}"
    )
    growth_fig.update_layout(height=380, margin=dict(l=40, r=30, t=60, b=40))
    st.plotly_chart(growth_fig, use_container_width=True)

# ----------------------- ТАБЛИЦЫ -----------------------

st.markdown("### Таблицы")
tab1, tab2, tab3, tab4 = st.tabs([
    "Сырые строки (после фильтров)",
    "Итоги по годам (План/Факт)",
    "Свод по дивизионам/годам",
    "A&P по годам"
])

with tab1:
    sort_cols = [c for c in ["Год", "Дивизион", "Продукт"] if c in sales_f.columns]
    st.dataframe(sales_f.sort_values(sort_cols, na_position="last"),
                 use_container_width=True, height=420)

with tab2:
    year_summary = sales_f.groupby("Год", as_index=False)[["План,$", "Факт,$"]].sum()
    year_summary = year_summary.rename(columns={"План,$": "План итого, $",
                                                "Факт,$": "Факт итого, $"})
    year_summary["% выполнения"] = (
        (year_summary["Факт итого, $"] / year_summary["План итого, $"]) * 100.0
    ).round(1)
    st.dataframe(year_summary, use_container_width=True, height=360)

with tab3:
    if "Дивизион" in sales_f.columns:
        div_year = sales_f.groupby(["Дивизион", "Год"], as_index=False)[["Факт,$", "План,$"]].sum()
        div_year = div_year.rename(columns={"Факт,$": "Факт, $", "План,$": "План, $"})
        st.dataframe(div_year.sort_values(["Год", "Факт, $"], ascending=[True, False]),
                     use_container_width=True, height=360)
    else:
        st.info("Колонки «Дивизион» нет — этот срез недоступен.")

with tab4:
    if not ap_f.empty:
        ap_summary = ap_f.groupby("Год", as_index=False)["A&P план,$"].sum()
        st.dataframe(ap_summary, use_container_width=True, height=300)
    else:
        st.info("В файле не обнаружены столбцы вида «План A&P{год},$». "
                "Если нужны YoY-графики по бюджету, добавь их в Excel.")

# ----------------------- ВЫГРУЗКА -----------------------

st.markdown("### Экспорт")
exp_cols = st.columns(3)

with exp_cols[0]:
    csv_bytes = sales_f.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Сырые строки (CSV)", data=csv_bytes,
                       file_name="filtered_sales_rows.csv", mime="text/csv")

with exp_cols[1]:
    xls_bytes = to_excel_download({
        "sales_filtered": sales_f,
        "year_summary": year_summary if 'year_summary' in locals() else pd.DataFrame(),
        "div_year": div_year if 'div_year' in locals() else pd.DataFrame(),
        "ap_filtered": ap_f
    })
    st.download_button("⬇️ Своды (Excel)", data=xls_bytes,
                       file_name="dashboard_exports.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
