"""
Streamlit app — Manual grading assistant for exam .docx files.

Usage:
    streamlit run grading_app.py
"""

from __future__ import annotations

import json
import os
import tempfile
from typing import Dict, List, Optional

import streamlit as st
from PIL import Image

from docx_parser import (
    StudentSubmission,
    MajorQuestion,
    parse_reference_docx,
    parse_student_docx,
    question_signature,
)
from scoring import SCORE_CONFIG, export_to_xlsx

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="考试批改助手",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Session state keys
# ---------------------------------------------------------------------------

STATE_DEFAULTS = {
    "reference_docx": None,          # UploadedFile
    "reference_questions": None,     # List[MajorQuestion]
    "reference_signature": None,     # str
    "reference_filename": None,      # str
    "student_files": [],             # List[UploadedFile]
    "submissions": {},               # Dict[str, StudentSubmission]  (UploadedFile.name -> submission)
    "scores": {},                    # Dict[str, Dict[str, float]]   (filename -> {key: score})
    "current_student": None,         # str — filename key
    "current_tab": 0,                # int — question tab
}

for key, val in STATE_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = val


# ---------------------------------------------------------------------------
# Helper: score key builder
# ---------------------------------------------------------------------------

def skey(major: int, sub: int, section: str) -> str:
    return f"{major}-{sub}-{section}"

# ---------------------------------------------------------------------------
# Helper: ensure scores dict structure
# ---------------------------------------------------------------------------

def init_scores(filename: str):
    if filename not in st.session_state.scores:
        st.session_state.scores[filename] = {}

# ---------------------------------------------------------------------------
# Auto-save / auto-restore scores to/from JSON
# ---------------------------------------------------------------------------

AUTOSAVE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output", "autosave.json")


def save_scores():
    """Write current scores dict + reference signature to autosave JSON."""
    try:
        os.makedirs(os.path.dirname(AUTOSAVE_PATH), exist_ok=True)
        data = {
            "ref_sig": st.session_state.reference_signature,
            "scores": st.session_state.scores,
        }
        with open(AUTOSAVE_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # silent — don't disrupt the UI for a save failure


def load_scores():
    """Restore scores from autosave JSON (only if reference signature matches)."""
    if not os.path.exists(AUTOSAVE_PATH):
        return
    try:
        with open(AUTOSAVE_PATH, "r", encoding="utf-8") as f:
            saved = json.load(f)
        # Signature guard — only restore if the same reference doc is loaded
        if saved.get("ref_sig") != st.session_state.reference_signature:
            # Delete stale autosave from old/other reference
            try:
                os.remove(AUTOSAVE_PATH)
            except Exception:
                pass
            return
        for fname, fdata in saved.get("scores", {}).items():
            if fname not in st.session_state.scores:
                st.session_state.scores[fname] = {}
            for k, v in fdata.items():
                if v is not None:
                    st.session_state.scores[fname][k] = v
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Incomplete grading check
# ---------------------------------------------------------------------------

def check_incomplete_grading(student_names: List[str]) -> List[tuple]:
    """Return list of (student_name_or_id, missing_keys) for students with unfinished grading."""
    incomplete = []
    for fname in student_names:
        scores = st.session_state.scores.get(fname, {})
        sub = st.session_state.submissions.get(fname)
        if sub is None:
            continue
        missing = []
        for q in sub.major_questions:
            for s in q.sub_questions:
                if scores.get(skey(q.index, s.index, "程序")) is None:
                    missing.append(f"Q{q.index}({s.index})程序")
                if scores.get(skey(q.index, s.index, "结果")) is None:
                    missing.append(f"Q{q.index}({s.index})结果")
        if missing:
            info = sub.student_info
            name = info.name or info.student_id or fname
            incomplete.append((name, missing))
    return incomplete


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

with st.sidebar:
    st.title("📝 考试批改助手")

    # ---- Reference docx ----
    st.subheader("📄 参考答案文档")
    ref_file = st.file_uploader(
        "上传参考答案 .docx",
        type=["docx"],
        key="ref_uploader",
        help="上传教师提供的参考答案Word文档",
    )

    if ref_file is not None and ref_file != st.session_state.reference_docx:
        with st.spinner("正在解析参考答案..."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    tmp.write(ref_file.getbuffer())
                    tmp_path = tmp.name
                st.session_state.reference_questions = parse_reference_docx(tmp_path)
                st.session_state.reference_signature = question_signature(
                    st.session_state.reference_questions
                )
                st.session_state.reference_docx = ref_file
                st.session_state.reference_filename = ref_file.name
                os.unlink(tmp_path)
                st.success(f"已加载参考答案: {ref_file.name}")
                # 参考答案就绪后，尝试恢复匹配的 autosave 分数
                load_scores()
            except Exception as e:
                st.error(f"解析失败: {e}")
                st.session_state.reference_questions = None

    if st.session_state.reference_questions:
        qs = st.session_state.reference_questions
        st.info(f"✅ 参考答案已加载 ({len(qs)} 大题)")
        ref_sig = st.session_state.reference_signature or ""
        # Show short preview
        for q in qs:
            st.caption(f"  Q{q.index}: {q.title[:50]}…" if len(q.title) > 50 else f"  Q{q.index}: {q.title}")
    else:
        st.warning("请上传参考答案文档")

    st.divider()

    # ---- Student docx files ----
    st.subheader("👥 学生答卷")
    stu_files = st.file_uploader(
        "上传学生答卷 .docx (可多选)",
        type=["docx"],
        accept_multiple_files=True,
        key="stu_uploader",
        help="选择学生提交的Word答卷文件，可批量上传",
    )

    if stu_files:
        # Process new files
        new_count = 0
        for f in stu_files:
            if f.name not in st.session_state.submissions:
                with st.spinner(f"解析: {f.name}..."):
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                            tmp.write(f.getbuffer())
                            tmp_path = tmp.name
                        sub = parse_student_docx(tmp_path)
                        sub.filename = f.name  # 用原始文件名而非临时路径，保证与 scores key 一致
                        st.session_state.submissions[f.name] = sub
                        init_scores(f.name)
                        os.unlink(tmp_path)
                        new_count += 1
                    except Exception as e:
                        st.error(f"{f.name}: 解析失败 — {e}")

        # Update file list
        st.session_state.student_files = stu_files
        if new_count > 0:
            st.success(f"已解析 {new_count} 份答卷")

    # ---- Student selector ----
    student_names = list(st.session_state.submissions.keys())
    if student_names:
        st.divider()
        st.subheader("🧑 当前学生")

        # Current student index
        current = st.session_state.current_student
        if current is None or current not in st.session_state.submissions:
            current = student_names[0]
            st.session_state.current_student = current

        idx = student_names.index(current) if current in student_names else 0

        col1, col2, col3 = st.columns([1, 2, 1])
        with col1:
            if st.button("◀ 上一人", use_container_width=True) and idx > 0:
                st.session_state.current_student = student_names[idx - 1]
                st.rerun()
        with col2:
            st.caption(f"{idx + 1} / {len(student_names)}")
        with col3:
            if st.button("下一人 ▶", use_container_width=True) and idx < len(student_names) - 1:
                st.session_state.current_student = student_names[idx + 1]
                st.rerun()

        # Dropdown selector
        sel = st.selectbox(
            "选择学生",
            student_names,
            index=idx,
            label_visibility="collapsed",
            key="student_selector",
        )
        if sel != current:
            st.session_state.current_student = sel
            st.rerun()

        # Show student info
        sub = st.session_state.submissions[current]
        info = sub.student_info
        if info.name or info.student_id:
            st.markdown(f"**{info.name or '?'}**  ({info.student_id or '?'})")
            if info.class_name:
                st.caption(info.class_name)

        # ---- Progress (current student, sub-question count) ----
        sub_qs = sum(len(q.sub_questions) for q in sub.major_questions)  # e.g. 20
        graded_sub = sum(
            1 for q in sub.major_questions
            for s in q.sub_questions
            if (st.session_state.scores.get(current, {}).get(skey(q.index, s.index, "程序"), None) is not None
                and st.session_state.scores.get(current, {}).get(skey(q.index, s.index, "结果"), None) is not None)
        )
        sub_pct = graded_sub / sub_qs if sub_qs > 0 else 0
        st.progress(sub_pct, text=f"已批改: {graded_sub}/{sub_qs} 小题")

        if st.button("🔄 重置当前学生分数", use_container_width=True):
            if current in st.session_state.scores:
                st.session_state.scores[current] = {}
                save_scores()
                st.rerun()

        # ---- Variant check ----
        if st.session_state.reference_signature:
            stu_sig = question_signature(sub.major_questions)
            if stu_sig != st.session_state.reference_signature:
                st.warning("⚠️ 学生答卷与参考答案的题目不一致，请确认！")

        # ---- Export ----
        st.divider()
        if st.button("📊 导出成绩 (.xlsx)", use_container_width=True, type="primary"):
            # Check for incomplete grading
            incomplete = check_incomplete_grading(student_names)
            if incomplete:
                msg_parts = []
                for name, missing in incomplete:
                    shown = missing[:3]
                    rest  = len(missing) - 3
                    detail = "、".join(shown)
                    if rest > 0:
                        detail += f"…等{rest+3}项"
                    msg_parts.append(f"• {name}: {detail}")
                st.warning("⚠️ 以下学生存在未批改的题目:\n" + "\n".join(msg_parts))
            # Proceed with export
            export_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "output",
                "批改成绩.xlsx",
            )
            try:
                subs = [st.session_state.submissions[f] for f in student_names]
                path = export_to_xlsx(subs, st.session_state.scores, export_path)
                st.success(f"已导出: {path}")
                # Offer download
                with open(path, "rb") as fb:
                    st.download_button(
                        "⬇ 下载 xlsx",
                        data=fb,
                        file_name="批改成绩.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as e:
                st.error(f"导出失败: {e}")
    else:
        st.info("请上传学生答卷文件")


# ---------------------------------------------------------------------------
# Main area — grading view
# ---------------------------------------------------------------------------

def render_student_answer(
    sub_file: str,
    ref_questions: Optional[List[MajorQuestion]],
):
    """Render grading UI for one student."""
    submission = st.session_state.submissions.get(sub_file)
    if not submission:
        st.warning("请先选择一名学生")
        return

    student_name = submission.student_info.name or sub_file
    student_id   = submission.student_info.student_id or ""
    scores       = st.session_state.scores.get(sub_file, {})

    # Clamp current_tab in case question count changed
    q_count = len(submission.major_questions)
    if st.session_state.current_tab >= q_count:
        st.session_state.current_tab = max(0, q_count - 1)

    # Title
    title = f"🧑 **{student_name}**"
    if student_id:
        title += f"　学号: {student_id}"
    st.markdown(f"# {title}")

    # Progress summary
    scores_local = scores
    total_subs = sum(len(q.sub_questions) for q in submission.major_questions)
    graded_subs = sum(
        1 for q in submission.major_questions
        for s in q.sub_questions
        if scores_local.get(skey(q.index, s.index, "程序")) is not None
           and scores_local.get(skey(q.index, s.index, "结果")) is not None
    )
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric("已批改", f"{graded_subs}/{total_subs} 小题")
    with col_m2:
        st.metric("未批改", f"{total_subs - graded_subs} 小题")
    with col_m3:
        if graded_subs > 0:
            total_pts = sum(
                min(
                    (scores_local.get(skey(q.index, s.index, "程序"), 0.0) +
                     scores_local.get(skey(q.index, s.index, "结果"), 0.0)),
                    SCORE_CONFIG["sub_max"],
                )
                for q in submission.major_questions
                for s in q.sub_questions
                if scores_local.get(skey(q.index, s.index, "程序")) is not None
                   and scores_local.get(skey(q.index, s.index, "结果")) is not None
            )
            st.metric("当前总分", f"{total_pts:.1f} / {total_subs * int(SCORE_CONFIG['sub_max'])}")
    st.divider()

    # ---- Custom question tabs (top + bottom) ----
    q_count = min(4, len(submission.major_questions))
    tab_labels = [f"📌 第{'一二三四'[i]}题" for i in range(q_count)]

    def render_tab_bar(key_prefix: str):
        """Render a row of question-selector buttons."""
        cols = st.columns(q_count)
        for i in range(q_count):
            with cols[i]:
                is_active = st.session_state.current_tab == i
                if st.button(
                    tab_labels[i],
                    use_container_width=True,
                    type="primary" if is_active else "secondary",
                    key=f"{key_prefix}_tab_{i}",
                ):
                    st.session_state.current_tab = i
                    st.rerun()

    # Top tab bar
    render_tab_bar("top")

    # Render active question
    q_idx = st.session_state.current_tab
    if q_idx < len(submission.major_questions):
        stu_q = submission.major_questions[q_idx]
        ref_q = ref_questions[q_idx] if ref_questions and q_idx < len(ref_questions) else None

        # Major question title
        q_title = stu_q.title or f"第{q_idx+1}题"
        st.subheader(f"**{q_title}**　(25分)")

        # Render each sub-question
        for sub_idx in range(5):
            render_sub_question(q_idx + 1, sub_idx + 1, stu_q, ref_q, sub_file)

        # Major question total
        major_total = sum(
            min(
                ((scores.get(skey(q_idx + 1, s, "程序")) or 0.0) +
                 (scores.get(skey(q_idx + 1, s, "结果")) or 0.0)),
                SCORE_CONFIG["sub_max"],
            )
            for s in range(1, 6)
        )
        st.markdown(f"**第{q_idx+1}题得分: {major_total:.1f} / 25**")
    else:
        st.info("该大题无数据")

    # Bottom tab bar
    st.markdown("---")
    render_tab_bar("bottom")


def render_sub_question(
    q_idx: int,
    s_idx: int,
    stu_q: MajorQuestion,
    ref_q: Optional[MajorQuestion],
    sub_file: str,
):
    """Render one sub-question with reference, student answer, and scoring."""
    scores = st.session_state.scores.get(sub_file, {})

    # Find sub-question data
    stu_sub = None
    if stu_q:
        for sq in stu_q.sub_questions:
            if sq.index == s_idx:
                stu_sub = sq
                break

    ref_sub = None
    if ref_q:
        for sq in ref_q.sub_questions:
            if sq.index == s_idx:
                ref_sub = sq
                break

    has_prog_t = scores.get(skey(q_idx, s_idx, "程序")) is not None
    has_res_t  = scores.get(skey(q_idx, s_idx, "结果")) is not None
    both_done  = has_prog_t and has_res_t
    status_icon = "✅" if both_done else "⏳"
    sub_title = f"{status_icon} ({s_idx})　{stu_sub.question_text[:80] if stu_sub else ''}"

    # Left accent border + main content
    accent_color = "#4ade80" if both_done else "#e2e8f0"
    accent, main_col = st.columns([0.03, 0.97])
    with accent:
        st.markdown(f'<div style="background:{accent_color};border-radius:3px;height:100%;min-height:80px;">&nbsp;</div>', unsafe_allow_html=True)
    with main_col:
        st.markdown(f"**{sub_title}**　(5分 — 程序3分 + 结果2分)")

        # ---- 程序 section ----
        col1, col2, col3 = st.columns([3, 3, 2])
        with col1:
            st.caption("📖 参考答案 — 程序")
            if ref_sub and ref_sub.program and ref_sub.program.image_bytes:
                st.image(ref_sub.program.image_bytes, use_container_width=True)
            else:
                st.info("无参考数据")

        with col2:
            st.caption("📸 学生答案 — 程序")
            if stu_sub and stu_sub.program:
                if stu_sub.program.image_bytes:
                    st.image(stu_sub.program.image_bytes, use_container_width=True)
                elif stu_sub.program.text_content:
                    st.text_area("文字答案", stu_sub.program.text_content, height=100, disabled=True,
                                key=f"{sub_file}-{q_idx}-{s_idx}-prog-text")
                else:
                    st.error("❌ 未作答")
            else:
                st.error("❌ 未作答")

        with col3:
            score_key_prog = skey(q_idx, s_idx, "程序")
            current_prog = scores.get(score_key_prog, None)
            prog_options = [0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0]
            prog_idx = None if current_prog is None else (prog_options.index(current_prog) if current_prog in prog_options else None)
            prog_score = st.radio(
                f"程序分",
                options=prog_options,
                index=prog_idx,
                format_func=lambda x: f"{x:.0f}" if x == int(x) else f"{x:.1f}",
                horizontal=True,
                key=f"{sub_file}-{q_idx}-{s_idx}-prog",
                help="程序(代码截图) 满分3分",
                label_visibility="collapsed",
            )
            if prog_score is not None and prog_score != current_prog:
                st.session_state.scores[sub_file][score_key_prog] = prog_score
                save_scores()

        # ---- 结果 section ----
        col1, col2, col3 = st.columns([3, 3, 2])
        with col1:
            st.caption("📖 参考答案 — 结果")
            if ref_sub and ref_sub.result and ref_sub.result.image_bytes:
                st.image(ref_sub.result.image_bytes, use_container_width=True)
            else:
                st.info("无参考数据")

        with col2:
            st.caption("📸 学生答案 — 结果")
            if stu_sub and stu_sub.result:
                if stu_sub.result.image_bytes:
                    st.image(stu_sub.result.image_bytes, use_container_width=True)
                elif stu_sub.result.text_content:
                    st.text_area("文字答案", stu_sub.result.text_content, height=100, disabled=True,
                                key=f"{sub_file}-{q_idx}-{s_idx}-res-text")
                else:
                    st.error("❌ 未作答")
            else:
                st.error("❌ 未作答")

        with col3:
            score_key_res = skey(q_idx, s_idx, "结果")
            current_res = scores.get(score_key_res, None)
            res_options = [0, 0.5, 1.0, 1.5, 2.0]
            res_idx = None if current_res is None else (res_options.index(current_res) if current_res in res_options else None)
            res_score = st.radio(
                f"结果分",
                options=res_options,
                index=res_idx,
                format_func=lambda x: f"{x:.0f}" if x == int(x) else f"{x:.1f}",
                horizontal=True,
                key=f"{sub_file}-{q_idx}-{s_idx}-res",
                help="结果(输出截图) 满分2分",
                label_visibility="collapsed",
            )
            if res_score is not None and res_score != current_res:
                st.session_state.scores[sub_file][score_key_res] = res_score
                save_scores()

        # Sub-total for this sub-question (only when both scored)
        prog_val = scores.get(skey(q_idx, s_idx, "程序"))
        res_val  = scores.get(skey(q_idx, s_idx, "结果"))
        if prog_val is not None and res_val is not None:
            sub_total = min(prog_val + res_val, SCORE_CONFIG["sub_max"])
            st.caption(f"本小题得分: {sub_total:.1f} / 5 ✅")
        else:
            st.caption("⏳ 尚未批改")
    st.divider()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    # Restore autosave scores if reference signature matches
    load_scores()

    # ---- Custom CSS for better appearance ----
    st.markdown("""
    <style>
        /* Score radio buttons as clickable chips */
        div[data-testid="stRadio"] > div {
            gap: 0.15rem !important;
            flex-wrap: wrap;
        }
        div[data-testid="stRadio"] label {
            border: 1px solid #d0d5dd;
            border-radius: 6px;
            padding: 0.15rem 0.45rem;
            margin: 0.1rem 0.05rem !important;
            background: #f8f9fa;
            font-size: 0.85rem !important;
            transition: all 0.1s ease;
        }
        div[data-testid="stRadio"] label:hover {
            background: #e2e8f0;
            border-color: #94a3b8;
        }
        div[data-testid="stRadio"] label[data-selected="true"],
        div[data-testid="stRadio"] input:checked + div {
            background: #dbeafe !important;
            border-color: #3b82f6 !important;
            font-weight: 600;
        }
        /* Compact captions */
        .stCaption {
            font-size: 0.8rem;
        }
        /* Sub-question card with left accent border */
        .sub-card {
            padding: 0.5rem 0 0.75rem 0.75rem;
            border-left: 3px solid #e2e8f0;
            margin-bottom: 0.25rem;
        }
        .sub-card.graded {
            border-left-color: #4ade80;
        }
        /* Tab bar buttons — connected tab look */
        div.row-widget.stButton > button {
            border-radius: 8px 8px 0 0 !important;
            border-bottom: none !important;
            font-size: 0.9rem !important;
            font-weight: 500;
            transition: all 0.15s ease;
        }
        div.row-widget.stButton > button[kind="primary"] {
            background: #dbeafe !important;
            color: #1e40af !important;
            border-color: #93c5fd !important;
            border-bottom: 2px solid #dbeafe !important;
            font-weight: 600;
        }
        div.row-widget.stButton > button[kind="secondary"] {
            background: #f8fafc !important;
            color: #64748b !important;
            border-color: #e2e8f0 !important;
        }
        div.row-widget.stButton > button[kind="secondary"]:hover {
            background: #f1f5f9 !important;
            color: #334155 !important;
        }
    </style>
    """, unsafe_allow_html=True)

    ref_qs = st.session_state.reference_questions
    current = st.session_state.current_student

    if not current or current not in st.session_state.submissions:
        if st.session_state.submissions:
            st.session_state.current_student = list(st.session_state.submissions.keys())[0]
            st.rerun()
        else:
            st.info("👈 请先在左侧上传参考答案和学生答卷文件")
            st.markdown("""
            ### 使用说明
            1. 上传 **参考答案 .docx** 文档
            2. 上传 **学生答卷 .docx** 文件（可多选）
            3. 在左侧选择学生，逐题批改
            4. 点击 **导出成绩** 生成 xlsx 文件
            """)
            return

    render_student_answer(current, ref_qs)


if __name__ == "__main__":
    main()
