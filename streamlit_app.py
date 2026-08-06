from __future__ import annotations

from datetime import date, datetime
import hashlib
import html
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from reconciliation import (
    ReconciliationError,
    Result,
    basis_config,
    export_excel,
    module_config,
    reconcile_module_all,
)


APP_NAME = "POOWARD ReconcileFlow"
MAX_COMBINED_UPLOAD_BYTES = 150 * 1024 * 1024
MODULES = {"接单差异": "order", "出货差异": "shipment"}
BASES = {"人民币 RMB": "rmb", "原币": "original"}

DOCUMENT_PREFERRED = [
    "DAY", "BIL", "CO", "CPO NO", "JO", "PART-DWG", "QTY", "CURR",
    "UP-CPO", "TP-CPO", "EX-CH", "VAT  PRICE", "VAT PRICE", "INVOICE",
    "匹配方式", "关联流水号", "来源表", "来源行号",
]
ORDER_PREFERRED = [
    "数据来源", "接单时间", "客户代码", "订单流水号", "客户PO", "零件名称",
    "订单数量", "单价", "交易金额", "交易币种", "接单汇率", "接单金额(RMB)",
    "出货运费(原币)", "出货运费(RMB)", "出货装箱单号", "实际出厂日期",
    "出货通知单号", "来源表", "来源行号",
]
SHIPMENT_PREFERRED = [
    "数据来源", "出货日期", "实际出货日期", "实际出厂日期", "客户代码",
    "订单流水号", "客户PO", "零件名称", "出货数量", "订单数量", "出货金额",
    "实际出货金额", "交易币种", "出货金额(RMB)", "实际出货金额(RMB)",
    "出货运费(原币)", "出货运费(RMB)", "出货装箱单号", "出货通知单号",
    "来源表", "来源行号",
]
INTERNAL_COLUMNS = {"客户代码_标准", "订单流水号_标准", "匹配键"}


def load_css() -> None:
    css_path = Path(__file__).with_name("assets") / "styles.css"
    st.markdown(f"<style>{css_path.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)


def init_state() -> None:
    defaults: dict[str, Any] = {
        "result_cache": {},
        "export_cache": {},
        "module_revision": {"order": 0, "shipment": 0},
        "document_digest": None,
        "last_success": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def file_bytes(uploaded: Any) -> bytes:
    return uploaded.getvalue()


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def format_amount(value: Any, basis: str) -> str:
    try:
        places = 2 if basis == "rmb" else 4
        return f"{float(value):,.{places}f}"
    except (TypeError, ValueError):
        return "—"


def format_cell(value: Any) -> str:
    if value is None:
        return "—"
    try:
        if bool(pd.isna(value)):
            return "—"
    except (TypeError, ValueError):
        pass
    if isinstance(value, (pd.Timestamp, datetime)):
        timestamp = pd.Timestamp(value)
        if timestamp.hour == timestamp.minute == timestamp.second == timestamp.microsecond == 0:
            return timestamp.strftime("%Y-%m-%d")
        return timestamp.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float):
        return f"{value:,.4f}".rstrip("0").rstrip(".")
    text = str(value).strip()
    return text or "—"


def render_brand() -> None:
    st.sidebar.markdown(
        """
        <div class="rf-brand">
          <span class="rf-brand-mark" aria-hidden="true">P</span>
          <span><strong>POOWARD</strong><small>ReconcileFlow</small></span>
        </div>
        <p class="rf-brand-note">接单与出货差异核对中心</p>
        """,
        unsafe_allow_html=True,
    )


def render_header(module: str, basis: str) -> None:
    module_info = module_config(module)
    config = basis_config(basis)
    if module == "order":
        english = "ORDER RECONCILIATION"
    else:
        english = "SHIPMENT RECONCILIATION"
    rule = (
        f"系统金额 = {module_info['amount_labels'][basis]} + {config['freight_label']}；"
        f"文控金额 = “{module_info['document_keyword']}”子表的 "
        f"{'VAT PRICE' if basis == 'rmb' else 'TP-CPO'}。"
    )
    st.markdown(
        f"""
        <header class="rf-hero">
          <div>
            <p class="rf-eyebrow">{english}</p>
            <h1>{html.escape(module_info['label'])}</h1>
            <p class="rf-lede">{html.escape(rule)}</p>
          </div>
          <div class="rf-formula" aria-label="差异计算公式">
            <span class="rf-formula-system">系统金额</span>
            <b>−</b>
            <span class="rf-formula-document">文控金额</span>
            <b>=</b>
            <em>差异</em>
          </div>
        </header>
        """,
        unsafe_allow_html=True,
    )


def render_upload_form(module: str) -> tuple[Any, Any, Any, float, bool]:
    module_info = module_config(module)
    st.markdown(
        f"""
        <div class="rf-section-heading">
          <span class="rf-step">01</span>
          <div><h2>上传{module_info['short_label']}核对文件</h2>
          <p>每次只需三份文件；上传框支持点击选择或直接拖入。</p></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.form(f"upload_form_{module}", clear_on_submit=False, border=True):
        document_col, amount_col, freight_col = st.columns([1.08, 1, 1], gap="medium")
        with document_col:
            st.markdown('<p class="rf-source-tag shared">共用数据 · SHARED</p>', unsafe_allow_html=True)
            document = st.file_uploader(
                "文控登记表 · Document Control Ledger",
                type=["xlsx", "xlsm"],
                key="shared_document",
                help="接单和出货模块共用；切换模块后无需重新选择。",
            )
            st.caption(f"读取名称包含“{module_info['document_keyword']}”的子表")
        with amount_col:
            st.markdown(f'<p class="rf-source-tag">{module_info["short_label"]}金额 · AMOUNT</p>', unsafe_allow_html=True)
            amount = st.file_uploader(
                f"{module_info['amount_source_label']} · Amount Detail",
                type=["xlsx", "xlsm"],
                key=f"{module}_amount_file",
            )
            st.caption(f"人民币：{module_info['amount_labels']['rmb']} ｜ 原币：{module_info['amount_labels']['original']}")
        with freight_col:
            st.markdown(f'<p class="rf-source-tag">{module_info["short_label"]}运费 · FREIGHT</p>', unsafe_allow_html=True)
            freight = st.file_uploader(
                f"{module_info['freight_source_label']} · Freight Detail",
                type=["xlsx", "xlsm"],
                key=f"{module}_freight_file",
            )
            st.caption("读取出货运费(RMB)及出货运费(原币)")

        controls, action = st.columns([1, 1.5], gap="large", vertical_alignment="bottom")
        with controls:
            tolerance = st.number_input(
                "差异容差",
                min_value=0.0,
                value=1.0,
                step=0.01,
                format="%.2f",
                help="绝对差异不超过该数值时视为一致。",
            )
        with action:
            st.caption("提交一次，同时计算并缓存人民币和原币结果。")
            submitted = st.form_submit_button(
                f"开始{module_info['short_label']}核对",
                type="primary",
                use_container_width=True,
            )
    return document, amount, freight, float(tolerance), submitted


def run_reconciliation(module: str, document: Any, amount: Any, freight: Any, tolerance: float) -> None:
    if not all((document, amount, freight)):
        st.error("请上传文控登记表、金额明细和运费明细三份文件。", icon="⚠️")
        return

    blobs = {
        "document": file_bytes(document),
        "amount": file_bytes(amount),
        "freight": file_bytes(freight),
    }
    total_size = sum(len(value) for value in blobs.values())
    if total_size > MAX_COMBINED_UPLOAD_BYTES:
        st.error("三份文件合计不能超过 150 MB。请精简无关工作表后再试。", icon="⚠️")
        return

    module_info = module_config(module)
    names = {
        "document": document.name,
        "amount": amount.name,
        "freight": freight.name,
    }
    document_digest = digest(blobs["document"])

    try:
        with st.status(
            f"正在读取三份文件并生成{module_info['short_label']}双口径结果…",
            expanded=True,
        ) as status:
            st.write("正在识别工作表和字段…")
            results = reconcile_module_all(
                blobs["document"],
                blobs["amount"],
                blobs["freight"],
                module=module,
                tolerance=tolerance,
                source_names=names,
            )
            st.write("人民币与原币结果已计算完成，正在写入当前会话缓存…")

            previous_digest = st.session_state.document_digest
            cache = dict(st.session_state.result_cache)
            if previous_digest and previous_digest != document_digest:
                cache = {}
                st.session_state.export_cache = {}
            cache[module] = results
            st.session_state.result_cache = cache
            st.session_state.document_digest = document_digest

            revisions = dict(st.session_state.module_revision)
            revisions[module] = int(revisions.get(module, 0)) + 1
            st.session_state.module_revision = revisions
            st.session_state.export_cache = {
                key: value
                for key, value in st.session_state.export_cache.items()
                if not str(key).startswith(f"{module}:")
            }
            st.session_state.last_success = module
            status.update(label="核对完成：两个口径均已缓存", state="complete", expanded=False)
        st.toast(f"{module_info['label']}已完成", icon="✅")
    except ReconciliationError as exc:
        st.error(str(exc), icon="⚠️")
    except Exception:
        st.error("处理文件时发生异常。请确认三份文件均为有效 Excel，并检查表头后重试。", icon="⚠️")
        st.caption("如需排查，请将此页面截图和三份文件的表头发给维护人员。")


def selected_rows(event: Any) -> list[int]:
    try:
        return list(event.selection.rows)
    except (AttributeError, TypeError):
        try:
            return list(event.get("selection", {}).get("rows", []))
        except (AttributeError, TypeError):
            return []


def comparison_frame(frame: pd.DataFrame, result: Result, *, line_level: bool) -> pd.DataFrame:
    config = basis_config(result.basis)
    module_info = module_config(result.module)
    rename = {
        "主金额": module_info["amount_labels"][result.basis],
        "运费": config["freight_label"],
        "系统金额": config["system_label"],
        "文控金额": config["document_label"],
        "差异": config["difference_label"],
    }
    if line_level:
        columns = [
            "订单流水号", "主金额", "运费", "系统金额", "文控金额", "差异",
            "状态", "系统行数", "文控行数",
        ]
    else:
        columns = [
            "客户代码", "主金额", "运费", "系统金额", "文控金额", "差异",
            "状态", "系统行数", "文控行数",
        ]
    return frame.reindex(columns=columns).rename(columns=rename).reset_index(drop=True)


def comparison_styler(frame: pd.DataFrame, result: Result) -> pd.io.formats.style.Styler:
    config = basis_config(result.basis)
    module_info = module_config(result.module)
    system_label = config["system_label"]
    document_label = config["document_label"]
    difference_label = config["difference_label"]
    numeric_columns = [
        module_info["amount_labels"][result.basis],
        config["freight_label"],
        system_label,
        document_label,
        difference_label,
    ]
    number_pattern = "{:,.2f}" if result.basis == "rmb" else "{:,.4f}"
    formats = {column: number_pattern for column in numeric_columns if column in frame.columns}
    styler = frame.style.format(formats, na_rep="—")
    styler = styler.set_properties(
        subset=[system_label],
        **{
            "background-color": "#e3f3f0",
            "color": "#0a5f58",
            "font-weight": "800",
            "font-size": "15px",
            "text-align": "right",
            "border-right": "3px solid #2f847c",
            "box-shadow": "inset -10px 0 14px -14px #0b5f58",
        },
    )
    styler = styler.set_properties(
        subset=[document_label],
        **{
            "background-color": "#fff6e7",
            "color": "#70480e",
            "font-weight": "800",
            "font-size": "15px",
            "text-align": "left",
            "box-shadow": "inset 10px 0 14px -14px #7b5319",
        },
    )
    if difference_label in frame.columns:
        styler = styler.map(
            lambda value: "color: #a13f3f; font-weight: 800;" if float(value) != 0 else "color: #526b67;",
            subset=[difference_label],
        )
    if "状态" in frame.columns:
        status_styles = {
            "有差异": "background-color: #fff0d6; color: #8b5107; font-weight: 800;",
            "仅文控表": "background-color: #fde9e9; color: #923b3b; font-weight: 800;",
            "仅系统表": "background-color: #e9edf8; color: #405789; font-weight: 800;",
        }
        styler = styler.map(lambda value: status_styles.get(str(value), ""), subset=["状态"])

    system_index = frame.columns.get_loc(system_label)
    document_index = frame.columns.get_loc(document_label)
    styler = styler.set_table_styles(
        [
            {
                "selector": f"th.col_heading.level0.col{system_index}",
                "props": [
                    ("background-color", "#cfeae5"), ("color", "#075b55"),
                    ("font-weight", "800"), ("text-align", "right"),
                    ("border-right", "3px solid #2f847c"),
                ],
            },
            {
                "selector": f"th.col_heading.level0.col{document_index}",
                "props": [
                    ("background-color", "#f6e7cd"), ("color", "#70480e"),
                    ("font-weight", "800"), ("text-align", "left"),
                ],
            },
        ],
        overwrite=False,
    )
    return styler


def comparison_column_config(result: Result, identity: str) -> dict[str, Any]:
    config = basis_config(result.basis)
    module_info = module_config(result.module)
    number_format = "%,.2f" if result.basis == "rmb" else "%,.4f"
    return {
        identity: st.column_config.TextColumn(identity, width="large", pinned=True),
        module_info["amount_labels"][result.basis]: st.column_config.NumberColumn(
            module_info["amount_labels"][result.basis], width=155, alignment="right", format=number_format,
        ),
        config["freight_label"]: st.column_config.NumberColumn(
            config["freight_label"], width=155, alignment="right", format=number_format,
        ),
        config["system_label"]: st.column_config.NumberColumn(
            config["system_label"], width=180, alignment="right", format=number_format,
        ),
        config["document_label"]: st.column_config.NumberColumn(
            config["document_label"], width=180, alignment="left", format=number_format,
        ),
        config["difference_label"]: st.column_config.NumberColumn(
            config["difference_label"], width=160, alignment="right", format=number_format,
        ),
        "状态": st.column_config.TextColumn("状态", width=110, alignment="center"),
        "系统行数": st.column_config.NumberColumn("系统行数", width=100, alignment="center", format="%d"),
        "文控行数": st.column_config.NumberColumn("文控行数", width=100, alignment="center", format="%d"),
    }


def render_kpis(result: Result) -> None:
    stats = result.statistics
    config = basis_config(result.basis)
    system_total = format_amount(stats.get("系统总额"), result.basis)
    document_total = format_amount(stats.get("文控总额"), result.basis)
    total_difference = format_amount(stats.get("总差异"), result.basis)
    st.markdown(
        f"""
        <section class="rf-kpi-grid" aria-label="核对关键指标">
          <article class="rf-kpi"><span>差异客户</span><strong>{int(stats.get('差异客户数', 0))}</strong><small>个客户代码</small></article>
          <article class="rf-kpi"><span>差异流水号</span><strong>{int(stats.get('差异流水号数', 0))}</strong><small>条待核对记录</small></article>
          <article class="rf-amount-axis">
            <div class="rf-amount system"><span>{html.escape(config['system_label'])}</span><strong>{system_total}</strong></div>
            <div class="rf-axis" aria-hidden="true"></div>
            <div class="rf-amount document"><span>{html.escape(config['document_label'])}</span><strong>{document_total}</strong></div>
          </article>
          <article class="rf-kpi difference"><span>{html.escape(config['difference_label'])}</span><strong>{total_difference}</strong><small>系统 − 文控</small></article>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_record(frame: pd.DataFrame, *, preferred: list[str], empty_message: str) -> None:
    if frame.empty:
        st.info(empty_message)
        return
    for ordinal, (_, row) in enumerate(frame.iterrows(), start=1):
        source = format_cell(row.get("来源表"))
        row_number = format_cell(row.get("来源行号"))
        label = f"第 {ordinal} 条 · {source} · Excel 第 {row_number} 行"
        with st.expander(label, expanded=len(frame) == 1):
            available = [column for column in preferred if column in row.index and column not in INTERNAL_COLUMNS]
            remaining = [
                str(column) for column in row.index
                if str(column) not in available and str(column) not in INTERNAL_COLUMNS
            ]
            ordered = available + remaining
            cells = []
            for column in ordered:
                value = format_cell(row[column])
                cells.append(
                    '<div class="rf-record-item">'
                    f'<dt>{html.escape(str(column))}</dt><dd>{html.escape(value)}</dd>'
                    "</div>"
                )
            st.markdown(f'<dl class="rf-record-grid">{"".join(cells)}</dl>', unsafe_allow_html=True)


def render_results(result: Result) -> None:
    module_info = module_config(result.module)
    config = basis_config(result.basis)
    revision = int(st.session_state.module_revision.get(result.module, 0))

    st.markdown(
        f"""
        <div class="rf-results-head">
          <div><p class="rf-eyebrow">{module_info['short_label'].upper()} · {config['short_label']}</p>
          <h2>{html.escape(module_info['label'])}结果</h2></div>
          <span class="rf-cache-pill"><i></i>双口径已缓存</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    export_key = f"{result.module}:{result.basis}:{revision}"
    if export_key not in st.session_state.export_cache:
        st.session_state.export_cache[export_key] = export_excel(result)
    filename = f"{module_info['short_label']}_{config['short_label']}_差异核对结果.xlsx"
    st.download_button(
        "下载当前口径 Excel 结果",
        data=st.session_state.export_cache[export_key],
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="secondary",
        use_container_width=False,
    )

    render_kpis(result)

    if result.warnings:
        with st.expander(f"核对提示 · {len(result.warnings)} 条", expanded=False, icon="⚠️"):
            for warning in result.warnings:
                st.markdown(f"- {html.escape(str(warning))}")

    st.markdown(
        """
        <div class="rf-section-heading result-section">
          <span class="rf-step">02</span>
          <div><h2>客户差异汇总</h2><p>仅展示非零差异。点击任意一行，下钻查看该客户的差异流水号。</p></div>
        </div>
        <div class="rf-axis-legend"><span class="system">系统金额向中线靠右对齐</span><i></i><span class="document">文控金额从中线向右展开</span></div>
        """,
        unsafe_allow_html=True,
    )

    search = st.text_input(
        "搜索客户代码",
        placeholder="例如 CB229",
        key=f"customer_search_{result.module}_{result.basis}_{revision}",
    ).strip().upper()
    customer_source = result.customer_differences.reset_index(drop=True)
    if search:
        customer_source = customer_source[
            customer_source["客户代码"].astype(str).str.upper().str.contains(search, regex=False)
        ].reset_index(drop=True)
    if customer_source.empty:
        if search:
            st.info("没有找到匹配的差异客户代码。")
        else:
            st.success("当前口径没有非零客户差异。", icon="✅")
        return

    customer_display = comparison_frame(customer_source, result, line_level=False)
    customer_event = st.dataframe(
        comparison_styler(customer_display, result),
        key=f"customers_{result.module}_{result.basis}_{revision}",
        hide_index=True,
        width="stretch",
        height=min(680, 42 * (len(customer_display) + 1) + 8),
        on_select="rerun",
        selection_mode="single-row",
        column_config=comparison_column_config(result, "客户代码"),
    )
    customer_rows = selected_rows(customer_event)
    if not customer_rows:
        st.caption("选择一位客户后，这里将出现该客户的差异流水号。")
        return

    customer_position = customer_rows[0]
    if customer_position >= len(customer_source):
        return
    customer = str(customer_source.iloc[customer_position]["客户代码"])
    line_source = result.line_differences[
        result.line_differences["客户代码"].astype(str).eq(customer)
    ].reset_index(drop=True)

    st.markdown(
        f"""
        <div class="rf-section-heading result-section">
          <span class="rf-step">03</span>
          <div><h2>{html.escape(customer)} 的差异流水号</h2>
          <p>共 {len(line_source)} 条非零差异；点击任意一行查看两侧原始数据。</p></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if line_source.empty:
        st.info("该客户没有非零流水号差异。")
        return

    line_display = comparison_frame(line_source, result, line_level=True)
    customer_key = hashlib.sha1(customer.encode("utf-8")).hexdigest()[:10]
    line_event = st.dataframe(
        comparison_styler(line_display, result),
        key=f"lines_{result.module}_{result.basis}_{revision}_{customer_key}",
        hide_index=True,
        width="stretch",
        height=min(620, 42 * (len(line_display) + 1) + 8),
        on_select="rerun",
        selection_mode="single-row",
        column_config=comparison_column_config(result, "订单流水号"),
    )
    line_rows = selected_rows(line_event)
    if not line_rows:
        st.caption("选择一条流水号后，下方将并排展示文控与系统原始行。")
        return

    line_position = line_rows[0]
    if line_position >= len(line_source):
        return
    line = line_source.iloc[line_position]
    match_key = str(line["匹配键"])
    order_number = str(line["订单流水号"])
    document_rows = result.document_rows[
        result.document_rows["客户代码_标准"].astype(str).eq(customer)
        & result.document_rows["匹配键"].astype(str).eq(match_key)
    ]
    system_rows = result.system_rows[
        result.system_rows["客户代码_标准"].astype(str).eq(customer)
        & result.system_rows["匹配键"].astype(str).eq(match_key)
    ]

    st.markdown(
        f"""
        <div class="rf-section-heading result-section">
          <span class="rf-step">04</span>
          <div><h2>流水号 {html.escape(order_number)} · 原始数据</h2>
          <p>文控与系统数据并排展示，来源表及 Excel 行号均保留。</p></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    document_col, divider_col, system_col = st.columns([1, 0.035, 1], gap="small")
    with document_col:
        st.markdown(
            f'<div class="rf-raw-head document"><span>文控登记表</span><strong>{len(document_rows)} 行</strong></div>',
            unsafe_allow_html=True,
        )
        render_record(document_rows, preferred=DOCUMENT_PREFERRED, empty_message="文控侧没有对应原始行。")
    with divider_col:
        st.markdown('<div class="rf-raw-divider" aria-hidden="true"></div>', unsafe_allow_html=True)
    with system_col:
        st.markdown(
            f'<div class="rf-raw-head system"><span>系统{module_info["short_label"]}数据</span><strong>{len(system_rows)} 行</strong></div>',
            unsafe_allow_html=True,
        )
        preferred = ORDER_PREFERRED if result.module == "order" else SHIPMENT_PREFERRED
        render_record(system_rows, preferred=preferred, empty_message="系统侧没有对应原始行。")


def main() -> None:
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="🔎",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    load_css()
    init_state()

    render_brand()
    module_label = st.sidebar.radio(
        "核对模块",
        options=list(MODULES),
        index=0,
        key="active_module_label",
    )
    module = MODULES[module_label]
    basis_label = st.sidebar.radio(
        "显示口径",
        options=list(BASES),
        index=0,
        key="active_basis_label",
    )
    basis = BASES[basis_label]
    st.sidebar.markdown(
        """
        <div class="rf-sidebar-note">
          <strong>一次提交，双口径缓存</strong>
          <span>人民币与原币会同时计算。切换模块时，已选择的文控表继续保留。</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.sidebar.caption("文件仅在当前浏览器会话中处理；关闭会话后缓存会失效。")

    render_header(module, basis)
    document, amount, freight, tolerance, submitted = render_upload_form(module)
    if submitted:
        run_reconciliation(module, document, amount, freight, tolerance)

    module_results = st.session_state.result_cache.get(module)
    if module_results and basis in module_results:
        st.divider()
        render_results(module_results[basis])
    else:
        st.markdown(
            """
            <div class="rf-empty-state">
              <span>等待核对</span>
              <strong>上传当前模块的三份文件</strong>
              <p>完成后可直接切换人民币与原币口径，并逐级查看客户、流水号及原始行。</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown('<footer class="rf-footer">POOWARD · ReconcileFlow · Streamlit Edition</footer>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
