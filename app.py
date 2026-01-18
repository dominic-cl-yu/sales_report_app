# app.py
# -*- coding: utf-8 -*-
"""简体中文 Streamlit 界面：销售透视报表生成器

运行：
    streamlit run app.py

依赖：见 requirements.txt
"""

from __future__ import annotations

import os
from datetime import datetime

import streamlit as st

from process import ReportError, generate_pivot_report_from_upload


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
            """
        )

    uploaded = st.file_uploader("上传 Excel 文件", type=["xlsx", "xls"])

    if uploaded is None:
        st.info("请选择并上传一个 Excel 文件。")
        return

    # 新文件：清理缓存，并固定一次报表日期
    if st.session_state.get("uploaded_name") != uploaded.name:
        st.session_state.pop("pivot_bytes", None)
        st.session_state.pop("stats", None)
        st.session_state["uploaded_name"] = uploaded.name
        st.session_state["auto_report_date"] = datetime.now().strftime("%b-%d")

    base_name = os.path.splitext(uploaded.name)[0]
    ext = os.path.splitext(uploaded.name)[1].lower()

    report_date = st.session_state.get("auto_report_date") or datetime.now().strftime("%b-%d")

    c1, c2 = st.columns([1, 1])
    with c1:
        run_btn = st.button("生成报表", type="primary", use_container_width=True)
    with c2:
        st.text_input("报表日期（自动）", value=report_date, disabled=True)

    if run_btn:
        excel_bytes = uploaded.getvalue()

        with st.status("正在处理……", expanded=True) as status:
            try:
                st.write("1）读取并校验工作簿……")
                pivot_bytes, stats = generate_pivot_report_from_upload(
                    excel_bytes=excel_bytes,
                    filename=uploaded.name,
                    report_date=report_date,
                )

                st.write("2）生成 Excel 报表完成。")

                st.session_state["pivot_bytes"] = pivot_bytes
                st.session_state["stats"] = stats

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
