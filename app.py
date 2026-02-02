# -*- coding: utf-8 -*-
"""简体中文 Streamlit 界面：销售透视报表生成器

运行：
    streamlit run app.py

说明：
- 生成的 Pivot 报表 Total 行使用 SUBTOTAL(109,...) 公式，
  在 Excel 里筛选/隐藏行时会自动跟随变化。
"""

from __future__ import annotations

import os
from datetime import datetime, date

import streamlit as st

from process import (
    ReportError,
    generate_pivot_report_from_upload,
    DATE_BASIS_EX_FTY,
    DATE_BASIS_CUSTOMER,
    DATE_BASIS_COLUMN_MAP,
)


def _pretty_list(items: list[str], max_items: int = 10) -> str:
    items = [str(x) for x in items if str(x).strip()]
    if not items:
        return "（无）"
    if len(items) <= max_items:
        return "、".join(items)
    return "、".join(items[:max_items]) + f" …（共 {len(items)} 个）"


def main() -> None:
    st.set_page_config(page_title="销售透视报表生成器", layout="wide")

    st.title("销售透视报表生成器")
    st.caption("上传订单状态 Excel，系统会自动清洗并生成 Pivot 报表（含 9 个 sheet：3 个工厂 × 3 个指标）。")

    with st.expander("使用说明", expanded=False):
        st.markdown(
            """
- 支持 `.xlsx` / `.xls`（`.xls` 需要安装 `xlrd`）。
- 不需要手动选择表头/列名：系统会自动扫描并匹配。
- Team 会按规则自动归类（如 SW / Sports / Fancy 等）。
- Customer 字段固定写入 `ALL`。
- Total 行使用 `SUBTOTAL(109, ...)`，筛选后会自动更新。
            """
        )

    uploaded = st.file_uploader("上传 Excel 文件", type=["xlsx", "xls"])

    if uploaded is None:
        st.info("请选择并上传一个 Excel 文件。")
        return

    def _set_report_date(d: date) -> None:
        """统一维护报表日期（date 对象），并同步到 date picker 的 widget state。"""
        st.session_state["report_date_obj"] = d
        st.session_state["report_date_picker"] = d

    def _reset_report_date_to_today() -> None:
        _set_report_date(datetime.now().date())

    # 新文件：清理缓存，并把"报表日期"重置为今天，"日期基准"重置为默认值
    if st.session_state.get("uploaded_name") != uploaded.name:
        st.session_state.pop("pivot_bytes", None)
        st.session_state.pop("stats", None)
        st.session_state["uploaded_name"] = uploaded.name
        st.session_state["date_basis"] = DATE_BASIS_EX_FTY
        _reset_report_date_to_today()

    base_name = os.path.splitext(uploaded.name)[0]

    # 统一以 date 对象存储（方便 date picker），在真正写入报表时再格式化为 "Jan-19" 这种字符串
    report_date_obj: date = st.session_state.get("report_date_obj") or datetime.now().date()
    report_date = report_date_obj.strftime("%b-%d")

    # 如果用户更改了报表日期，则让之前生成的文件失效，避免“日期已变但仍下载旧文件”的误解
    if "pivot_bytes" in st.session_state:
        last_gen = st.session_state.get("generated_report_date")
        if last_gen and last_gen != report_date:
            st.session_state.pop("pivot_bytes", None)
            st.session_state.pop("stats", None)
            st.session_state.pop("generated_report_date", None)
            st.info("你已更改报表日期，请重新点击“生成报表”以应用新日期。")

    # Date basis selector - controls which date column is used for month/year grouping
    DATE_BASIS_OPTIONS = {
        DATE_BASIS_EX_FTY: "Ex-Fty (Request Garment Delivery)",
        DATE_BASIS_CUSTOMER: "Customer Delivery Date",
    }
    current_date_basis = st.session_state.get("date_basis", DATE_BASIS_EX_FTY)
    
    # If date_basis changed, invalidate cached report
    if "pivot_bytes" in st.session_state:
        last_basis = st.session_state.get("generated_date_basis")
        if last_basis and last_basis != current_date_basis:
            st.session_state.pop("pivot_bytes", None)
            st.session_state.pop("stats", None)
            st.session_state.pop("generated_date_basis", None)
            st.info("你已更改日期基准，请重新点击「生成报表」以应用新设置。")
    
    c0, c1, c2, c3 = st.columns([1.2, 1, 1, 0.7])
    with c0:
        selected_basis = st.selectbox(
            "日期基准",
            options=list(DATE_BASIS_OPTIONS.keys()),
            format_func=lambda x: DATE_BASIS_OPTIONS[x],
            index=list(DATE_BASIS_OPTIONS.keys()).index(current_date_basis),
            key="date_basis_selector",
            help="选择用于生成月份/年份列的日期列",
        )
        st.session_state["date_basis"] = selected_basis
    with c1:
        run_btn = st.button("生成报表", type="primary", use_container_width=True)
    with c2:
        st.text_input("报表日期", value=report_date, disabled=True)
    with c3:
        # Streamlit 的 date_input 自带日历选择器，默认打开当前月份，可切换月份
        # 为了满足“按按钮弹出日历”的体验：优先用 popover（若版本不支持则降级为 expander）
        if hasattr(st, "popover"):
            with st.popover("选择日期"):
                picked = st.date_input(
                    "选择报表日期",
                    value=report_date_obj,
                    key="report_date_picker",
                )
                # 把 widget 结果同步到真正用于生成报表的 key
                st.session_state["report_date_obj"] = picked

                st.button(
                    "重置为今天",
                    on_click=_reset_report_date_to_today,
                    use_container_width=True,
                )
        else:
            with st.expander("选择日期", expanded=False):
                picked = st.date_input(
                    "选择报表日期",
                    value=report_date_obj,
                    key="report_date_picker",
                )
                st.session_state["report_date_obj"] = picked
                st.button(
                    "重置为今天",
                    on_click=_reset_report_date_to_today,
                    use_container_width=True,
                )

    if run_btn:
        excel_bytes = uploaded.getvalue()
        date_basis = st.session_state.get("date_basis", DATE_BASIS_EX_FTY)

        with st.status("正在处理……", expanded=True) as status:
            try:
                st.write("1）读取并校验工作簿……")
                pivot_bytes, stats = generate_pivot_report_from_upload(
                    excel_bytes=excel_bytes,
                    filename=uploaded.name,
                    report_date=report_date,
                    date_basis=date_basis,
                )

                st.write("2）生成 Excel 报表完成。")

                st.session_state["pivot_bytes"] = pivot_bytes
                st.session_state["stats"] = stats
                st.session_state["generated_report_date"] = report_date
                st.session_state["generated_date_basis"] = date_basis

                status.update(label="处理完成", state="complete")

            except ReportError as e:
                status.update(label="处理失败", state="error")
                st.error(str(e))
                return
            except Exception as e:
                status.update(label="发生未知错误", state="error")
                st.error(f"发生未知错误：{e}")
                return

    if "pivot_bytes" in st.session_state:
        stats = st.session_state.get("stats", {})

        st.success("报表已生成，可以下载。")

        warnings_text = (stats.get("warnings") or "").strip()
        if warnings_text:
            st.warning(warnings_text)

        m1, m2, m3 = st.columns(3)
        m1.metric("参与计算的行数", stats.get("rows_used", 0))
        m2.metric("工厂数量", len(stats.get("factories", [])))
        m3.metric("产品类型数量", len(stats.get("product_types", [])))

        with st.expander("更多信息", expanded=False):
            st.write(f"报表日期：{stats.get('report_date', '')}")
            st.write(f"日期基准：{stats.get('date_column', '')}")
            st.write(f"工厂：{_pretty_list(stats.get('factories', []), max_items=20)}")
            st.write(f"产品类型：{_pretty_list(stats.get('product_types', []), max_items=20)}")

        st.download_button(
            "下载报表（.xlsx）",
            data=st.session_state["pivot_bytes"],
            file_name=f"{base_name}_pivot.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
