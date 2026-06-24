"""
Streamlit app — Manual grading assistant for exam .docx files.

Usage:
    streamlit run grading_app.py
"""

from __future__ import annotations

import base64
import datetime
import json
import os
import re
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
from ipynb_parser import parse_reference_ipynb
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
    "reset_version": 0,              # int — bumped on student score reset, appended to widget keys
    "scoring_rubric": None,          # dict — loaded from points.json
    "scoring_rubric_file": None,     # UploadedFile — the uploaded .json file
    "scoring_rubric_filename": None,  # str
    "checkmarks": {},                # Dict[str, Dict[str, bool]]  (filename -> {checkmark_key: bool})
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
# Helper: scoring rubric lookup
# ---------------------------------------------------------------------------

_SECTION_KEY_MAP = {
    "程序": "program",
    "结果": "result",
}


def get_rubric_for(q_idx: int, s_idx: int) -> Optional[dict]:
    """Look up scoring rubric for a (major, sub) question from points.json."""
    rubric = st.session_state.scoring_rubric
    if not rubric:
        return None
    sections = rubric.get("sections", [])
    section_idx = q_idx - 1
    if section_idx < 0 or section_idx >= len(sections):
        return None
    subs = sections[section_idx].get("sub_questions", [])
    sub_idx = s_idx - 1
    if sub_idx < 0 or sub_idx >= len(subs):
        return None
    return subs[sub_idx]


def compute_section_score(fname: str, q_idx: int, s_idx: int, section_cn: str) -> float:
    """Compute numeric section score from checked rubric points."""
    rubric_data = get_rubric_for(q_idx, s_idx)
    if not rubric_data:
        return 0.0
    section_en = _SECTION_KEY_MAP.get(section_cn)
    if not section_en or section_en not in rubric_data:
        return 0.0
    section_data = rubric_data[section_en]
    points = section_data.get("points", [])
    max_score = section_data.get("max_score", 3.0)
    if not points:
        return 0.0
    total = 0.0
    cm = st.session_state.checkmarks.get(fname, {})
    for i, p in enumerate(points):
        cm_key = f"{q_idx}-{s_idx}-{section_cn}-{i}"
        if cm.get(cm_key, False):
            total += p.get("score", 0)
    return min(total, max_score)


def render_scoring_checkboxes(
    fname: str,
    q_idx: int,
    s_idx: int,
    section_cn: str,
    rubric_data: Optional[dict],
):
    """Render checkbox-based scoring for one section (程序/结果).

    Each rubric point gets a checkbox; total = sum of checked points.
    """
    section_en = _SECTION_KEY_MAP.get(section_cn)
    if not rubric_data or not section_en or section_en not in rubric_data:
        st.caption("📋 暂无评分标准")
        return

    section_data = rubric_data[section_en]
    points = section_data.get("points", [])
    max_score = section_data.get("max_score", 3)

    if not points:
        st.caption("📋 暂无评分标准")
        return

    st.markdown(f"**{section_cn}得分:**")

    total = 0.0
    changed = False
    cm = st.session_state.checkmarks.get(fname, {})
    rv = st.session_state.reset_version

    for i, p in enumerate(points):
        cm_key = f"{q_idx}-{s_idx}-{section_cn}-{i}"
        widget_key = f"ck-{fname}-{q_idx}-{s_idx}-{section_cn}-{i}-v{rv}"
        is_checked = cm.get(cm_key, False)

        checked = st.checkbox(
            f"{p['item']}  ({p['score']}分)",
            value=is_checked,
            key=widget_key,
        )

        if checked != is_checked:
            if fname not in st.session_state.checkmarks:
                st.session_state.checkmarks[fname] = {}
            st.session_state.checkmarks[fname][cm_key] = checked
            changed = True

        if checked:
            total += p["score"]

    total = min(total, max_score)

    # Show computed total
    st.markdown(f"**合计: {total:.1f} / {max_score}**")

    if changed:
        # Sync numeric scores dict so export / progress work as before
        if fname not in st.session_state.scores:
            st.session_state.scores[fname] = {}
        st.session_state.scores[fname][skey(q_idx, s_idx, section_cn)] = total
        save_scores()


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
    """Write current scores + checkmarks + reference signature to autosave JSON."""
    try:
        os.makedirs(os.path.dirname(AUTOSAVE_PATH), exist_ok=True)
        data = {
            "ref_sig": st.session_state.reference_signature,
            "scores": st.session_state.scores,
            "checkmarks": st.session_state.checkmarks,
        }
        with open(AUTOSAVE_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # silent — don't disrupt the UI for a save failure


def load_scores():
    """Restore scores and checkmarks from autosave JSON (only if reference signature matches)."""
    if not os.path.exists(AUTOSAVE_PATH):
        return
    try:
        with open(AUTOSAVE_PATH, "r", encoding="utf-8") as f:
            saved = json.load(f)
        # Signature guard — only restore if the same reference doc is loaded.
        # Never delete the file — it survives page refreshes so the user can
        # upload the same reference doc later and pick up where they left off.
        if saved.get("ref_sig") != st.session_state.reference_signature:
            return
        for fname, fdata in saved.get("scores", {}).items():
            if fname not in st.session_state.scores:
                st.session_state.scores[fname] = {}
            for k, v in fdata.items():
                if v is not None:
                    st.session_state.scores[fname][k] = v
        # Restore checkmarks
        saved_cm = saved.get("checkmarks", {})
        for fname, cm_data in saved_cm.items():
            if fname not in st.session_state.checkmarks:
                st.session_state.checkmarks[fname] = {}
            for k, v in cm_data.items():
                st.session_state.checkmarks[fname][k] = bool(v)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# student_points.txt — per-point scoring record
# ---------------------------------------------------------------------------
# Generate points text (no longer auto-saves — download buttons used instead)
# ---------------------------------------------------------------------------


def generate_points_text() -> str:
    """Build a human-readable per-point scoring record for all students."""
    lines = []
    lines.append("=" * 80)
    lines.append("学生评分明细记录")
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines.append(f"生成时间: {now_str}")
    lines.append("=" * 80)

    for fname in sorted(st.session_state.submissions.keys()):
        sub = st.session_state.submissions.get(fname)
        if not sub:
            continue
        info = sub.student_info
        scores = st.session_state.scores.get(fname, {})

        lines.append("")
        lines.append("-" * 80)
        lines.append(f"学生: {info.name or '未知'}  |  学号: {info.student_id or '未知'}  |  班级: {info.class_name or '未知'}")
        lines.append(f"文件名: {fname}")
        lines.append("-" * 80)

        for q in sub.major_questions:
            lines.append("")
            lines.append(f"  ◆ 第{q.index}题  {q.title[:80]}")

            for s in q.sub_questions:
                q_text = re.sub(r"^[（(]\d+[）)]\s*", "", s.question_text)[:60]
                lines.append(f"    ({s.index}) {q_text}")

                for section_cn in ["程序", "结果"]:
                    rubric_data = get_rubric_for(q.index, s.index)
                    section_en = _SECTION_KEY_MAP.get(section_cn)
                    points = []
                    max_score = 0
                    if rubric_data and section_en and section_en in rubric_data:
                        points = rubric_data[section_en].get("points", [])
                        max_score = rubric_data[section_en].get("max_score", 3)

                    checked_items = []
                    cm = st.session_state.checkmarks.get(fname, {})
                    for i, p in enumerate(points):
                        cm_key = f"{q.index}-{s.index}-{section_cn}-{i}"
                        if cm.get(cm_key, False):
                            checked_items.append(f"           ✓ {p['item']} (+{p['score']}分)")

                    total = scores.get(skey(q.index, s.index, section_cn), 0.0)
                    lines.append(f"      {section_cn}: {total:.1f}/{max_score} 分")
                    if checked_items:
                        lines.extend(checked_items)
                    else:
                        lines.append(f"        (无得分点)")

                sub_total = min(
                    scores.get(skey(q.index, s.index, "程序"), 0.0)
                    + scores.get(skey(q.index, s.index, "结果"), 0.0),
                    SCORE_CONFIG["sub_max"],
                )
                lines.append(f"      → 本小题: {sub_total:.1f}/{SCORE_CONFIG['sub_max']} 分")

            # Major total
            q_total = sum(
                min(
                    scores.get(skey(q.index, s.index, "程序"), 0.0)
                    + scores.get(skey(q.index, s.index, "结果"), 0.0),
                    SCORE_CONFIG["sub_max"],
                )
                for s in q.sub_questions
            )
            lines.append(f"    ★ 第{q.index}题总分: {q_total:.1f} / 25")

        # Grand total
        grand_total = sum(
            min(
                scores.get(skey(q.index, s.index, "程序"), 0.0)
                + scores.get(skey(q.index, s.index, "结果"), 0.0),
                SCORE_CONFIG["sub_max"],
            )
            for q in sub.major_questions
            for s in q.sub_questions
        )
        lines.append(f"")
        lines.append(f"  ★★★ 总分: {grand_total:.1f} / 100 ★★★")

    return "\n".join(lines)


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
# Navigation warning — warn on incomplete page when switching students
# ---------------------------------------------------------------------------

def get_missing_on_current_page(fname: str) -> List[str]:
    """Return missing score items for the currently displayed major question."""
    if not fname:
        return []
    q_idx = st.session_state.current_tab
    scores = st.session_state.scores.get(fname, {})
    missing = []
    for s in range(1, 6):
        if scores.get(skey(q_idx + 1, s, "程序")) is None:
            missing.append(f"Q{q_idx+1}({s})程序")
        if scores.get(skey(q_idx + 1, s, "结果")) is None:
            missing.append(f"Q{q_idx+1}({s})结果")
    return missing


def navigate_student(target_student: str):
    """Switch current_student, showing warning if current page has ungraded items."""
    current = st.session_state.current_student
    if current:
        missing = get_missing_on_current_page(current)
        if missing:
            q_cn = "一二三四"[st.session_state.current_tab]
            st.session_state._nav_warning = (
                f"⚠️ 第{q_cn}题还有 {len(missing)} 项未批改！"
            )
    st.session_state.current_student = target_student
    st.rerun()


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

with st.sidebar:
    st.title("📝 考试批改助手")

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
        # ---- Sync: remove submissions for files no longer in uploader ----
        current_names = {f.name for f in stu_files}
        removed = [k for k in st.session_state.submissions if k not in current_names]
        old_count = len(st.session_state.submissions)
        for k in removed:
            del st.session_state.submissions[k]
            if k in st.session_state.scores:
                del st.session_state.scores[k]
            if k in st.session_state.checkmarks:
                del st.session_state.checkmarks[k]
        if removed:
            st.rerun()

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
                        sub.filename = f.name
                        st.session_state.submissions[f.name] = sub
                        init_scores(f.name)
                        os.unlink(tmp_path)
                        new_count += 1
                    except Exception as e:
                        st.error(f"{f.name}: 解析失败 — {e}")

        st.session_state.student_files = stu_files
        if new_count > 0:
            st.success(f"已解析 {new_count} 份答卷")
    else:
        # DON'T auto-clear submissions — prevents data loss when uploader briefly returns empty on rerun
        pass

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
            if st.button("◀ 上一人", use_container_width=True, key="side_prev_stu"):
                if idx > 0:
                    navigate_student(student_names[idx - 1])
        with col2:
            st.caption(f"{idx + 1} / {len(student_names)}")
        with col3:
            if st.button("下一人 ▶", use_container_width=True, key="side_next_stu"):
                if idx < len(student_names) - 1:
                    navigate_student(student_names[idx + 1])

        # Dropdown selector — track last user-selected value to detect real interactions
        sel = st.selectbox(
            "选择学生",
            student_names,
            index=idx,
            label_visibility="collapsed",
            key="student_selector",
        )
        # Only update current_student when the user explicitly selects from the dropdown
        # (not when a nav button or other code changed current_student)
        last_sel = st.session_state.get("_last_sel_student")
        if last_sel is None:
            st.session_state._last_sel_student = sel
        elif sel != last_sel:
            # User selected a different student from the dropdown
            st.session_state._last_sel_student = sel
            st.session_state.current_student = sel
            st.rerun()
        elif st.session_state.current_student != sel:
            # Nav button changed current_student — sync tracker to selectbox's displayed value
            st.session_state._last_sel_student = sel

        # Show student info
        sub = st.session_state.submissions[current]
        info = sub.student_info
        if info.name or info.student_id:
            st.markdown(f"**{info.name or '?'}**  ({info.student_id or '?'})")
            if info.class_name:
                st.caption(info.class_name)

        # ---- Progress (current student, sub-question count) ----
        sub_qs = sum(len(q.sub_questions) for q in sub.major_questions)
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
            if current in st.session_state.checkmarks:
                st.session_state.checkmarks[current] = {}
            st.session_state.reset_version += 1
            save_scores()
            st.rerun()

        # ---- Variant check ----
        if st.session_state.reference_signature:
            stu_sig = question_signature(sub.major_questions)
            if stu_sig != st.session_state.reference_signature:
                st.warning("⚠️ 学生答卷与参考答案的题目不一致，请确认！")

        # ---- Question selector (global) ----
        st.divider()
        st.subheader("📌 当前批改题号")
        q_count = 4
        q_cols = st.columns(q_count)
        for i in range(q_count):
            with q_cols[i]:
                is_active = st.session_state.current_tab == i
                if st.button(
                    f"第{'一二三四'[i]}题",
                    use_container_width=True,
                    type="primary" if is_active else "secondary",
                    key=f"sidebar_q_{i}",
                ):
                    st.session_state.current_tab = i
                    st.rerun()
        # Progress for current question across all students
        q_idx = st.session_state.current_tab
        total_stu = len(student_names)
        graded_stu = sum(
            1 for fname in student_names
            if all(
                st.session_state.scores.get(fname, {}).get(skey(q_idx + 1, s, "程序")) is not None
                and st.session_state.scores.get(fname, {}).get(skey(q_idx + 1, s, "结果")) is not None
                for s in range(1, 6)
            )
        )
        if total_stu > 0:
            st.progress(graded_stu / total_stu,
                        text=f"第{'一二三四'[q_idx]}题 — {graded_stu}/{total_stu} 人已完成")

            # Auto-advance to next question when all students are done for this one
            question_labels = "一二三四"
            if graded_stu == total_stu and total_stu > 0:
                if q_idx < 3:  # not the last question
                    st.session_state._auto_advance = f"🎉 第{question_labels[q_idx]}题全部批改完成，已自动切换到第{question_labels[q_idx+1]}题"
                    st.session_state.current_tab = q_idx + 1
                    st.session_state.current_student = student_names[0]
                    st.rerun()
                else:
                    # Check if ALL four questions are fully graded
                    all_done = True
                    for qi in range(4):
                        done = sum(
                            1 for fname in student_names
                            if all(
                                st.session_state.scores.get(fname, {}).get(skey(qi + 1, s, "程序")) is not None
                                and st.session_state.scores.get(fname, {}).get(skey(qi + 1, s, "结果")) is not None
                                for s in range(1, 6)
                            )
                        )
                        if done < total_stu:
                            all_done = False
                            break
                    if all_done and not st.session_state.get("_all_done_shown"):
                        st.session_state._all_done_shown = True
                        st.balloons()
                        st.success("🎉🎉🎉 全部四道大题批改完成！可以导出成绩了 🎉🎉🎉")

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
                xlsx_bytes = export_to_xlsx(subs, st.session_state.scores)
                points_text = generate_points_text()
                st.session_state._export_xlsx = xlsx_bytes
                st.session_state._export_points = points_text
                st.session_state._export_ready = True
                st.rerun()
            except Exception as e:
                st.error(f"导出失败: {e}")

        # Show download buttons after export data is generated
        if st.session_state.get("_export_ready"):
            st.divider()
            st.subheader("📥 下载文件")
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "\U0001F4E5 批改成绩.xlsx",
                    data=st.session_state._export_xlsx,
                    file_name="批改成绩.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col2:
                st.download_button(
                    "\U0001F4E5 得分点记录.txt",
                    data=st.session_state._export_points,
                    file_name="student_points.txt",
                    mime="text/plain",
                    use_container_width=True,
                )
            if st.button("\U0001F504 重新生成", use_container_width=True):
                st.session_state._export_ready = False
                st.rerun()
    else:
        st.info("请上传学生答卷文件")

    st.divider()

    # ---- Autosave management ----
    if os.path.exists(AUTOSAVE_PATH):
        if st.button("🗑️ 清除自动保存记录 (autosave.json)", use_container_width=True,
                     key="clear_autosave"):
            try:
                os.remove(AUTOSAVE_PATH)
                st.success("已清除自动保存记录")
                st.rerun()
            except Exception as e:
                st.error(f"清除失败: {e}")

    st.divider()

    # ---- Combined reference / rubric uploader ----
    st.subheader("📄 参考答案 / 📋 评分标准")
    ref_rubric_files = st.file_uploader(
        "上传 .docx/.ipynb (参考答案) 或 .json (评分标准)",
        type=["docx", "ipynb", "json"],
        accept_multiple_files=True,
        key="ref_rubric_uploader",
        help="上传参考答案文档(.docx/.ipynb)和/或评分标准文件(.json)，根据文件扩展名自动识别",
    )

    # Determine what file types are currently uploaded
    current_docx = None
    current_ipynb = None
    current_json = None
    if ref_rubric_files:
        for f in ref_rubric_files:
            if f.name.lower().endswith(".docx"):
                current_docx = f
            elif f.name.lower().endswith(".ipynb"):
                current_ipynb = f
            elif f.name.lower().endswith(".json"):
                current_json = f

    # --- Process reference docx (filename-based comparison) ---
    if current_docx is not None:
        if st.session_state.reference_filename != current_docx.name:
            with st.spinner("正在解析参考答案 (.docx)..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                        tmp.write(current_docx.getbuffer())
                        tmp_path = tmp.name
                    st.session_state.reference_questions = parse_reference_docx(tmp_path)
                    st.session_state.reference_signature = question_signature(
                        st.session_state.reference_questions
                    )
                    st.session_state.reference_docx = current_docx
                    st.session_state.reference_filename = current_docx.name
                    os.unlink(tmp_path)
                    st.success(f"已加载参考答案: {current_docx.name}")
                    load_scores()
                except Exception as e:
                    st.error(f"参考答案解析失败: {e}")
    else:
        pass

    # --- Process reference ipynb ---
    if current_ipynb is not None:
        if st.session_state.reference_filename != current_ipynb.name:
            with st.spinner("正在解析参考答案 (.ipynb)..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".ipynb") as tmp:
                        tmp.write(current_ipynb.getbuffer())
                        tmp_path = tmp.name
                    st.session_state.reference_questions = parse_reference_ipynb(tmp_path)
                    st.session_state.reference_signature = question_signature(
                        st.session_state.reference_questions
                    )
                    st.session_state.reference_docx = current_ipynb
                    st.session_state.reference_filename = current_ipynb.name
                    os.unlink(tmp_path)
                    st.success(f"已加载参考答案: {current_ipynb.name}")
                    load_scores()
                except Exception as e:
                    st.error(f"参考答案解析失败: {e}")
    else:
        pass

    # --- Process rubric json (filename-based comparison) ---
    if current_json is not None:
        if st.session_state.scoring_rubric_filename != current_json.name:
            with st.spinner("正在加载评分标准..."):
                try:
                    content = json.loads(current_json.getvalue().decode("utf-8"))
                    if "sections" not in content:
                        raise ValueError("JSON 文件中缺少 'sections' 字段")
                    st.session_state.scoring_rubric = content
                    st.session_state.scoring_rubric_file = current_json
                    st.session_state.scoring_rubric_filename = current_json.name
                    st.success(f"已加载: {current_json.name}")
                except Exception as e:
                    st.error(f"评分标准加载失败: {e}")
                    # Don't clear old state on parse failure
    else:
        # DON'T auto-clear — prevents data loss on transient rerun
        pass

    # Show status of what's loaded
    if st.session_state.reference_questions:
        qs = st.session_state.reference_questions
        ref_type = "📓" if st.session_state.reference_filename.endswith(".ipynb") else "📄"
        st.info(f"{ref_type} 参考答案: {st.session_state.reference_filename} ({len(qs)} 大题)")
    elif current_docx is not None or current_ipynb is not None:
        st.warning("请上传参考答案 — 正在处理...")
    elif st.session_state.reference_filename is None:
        st.warning("请上传参考答案 (.docx / .ipynb)")

    if st.session_state.scoring_rubric:
        section_count = len(st.session_state.scoring_rubric.get("sections", []))
        rubric_name = st.session_state.scoring_rubric_filename or "?"
        st.info(f"✅ 评分标准: {rubric_name} ({section_count} 道大题)")
    elif current_json is not None:
        st.warning("请上传评分标准 (.json) — 正在处理...")
    elif st.session_state.scoring_rubric_filename is None:
        st.warning("请上传评分标准 (.json)")

    # Reload button — clears reference/rubric + wipes checkmarks for a fresh start
    if st.session_state.reference_questions or st.session_state.scoring_rubric:
        if st.button("🔄 重新加载", use_container_width=True, key="clear_ref_rubric"):
            st.session_state.reference_docx = None
            st.session_state.reference_questions = None
            st.session_state.reference_signature = None
            st.session_state.reference_filename = None
            st.session_state.scoring_rubric = None
            st.session_state.scoring_rubric_file = None
            st.session_state.scoring_rubric_filename = None
            # Clear checkmarks since they are only valid for the previous rubric
            st.session_state.checkmarks = {}
            st.session_state._all_done_shown = False
            st.rerun()


# ---------------------------------------------------------------------------
# Clickable image helper
# ---------------------------------------------------------------------------

def _render_zoomable_image(
    image_bytes: bytes,
    sub_file: str,
    q_idx: int,
    s_idx: int,
    section: str,
):
    """Render a student answer image that zooms in a full-screen overlay on click.

    Uses a pure-CSS :target lightbox — click the image to zoom, click the
    dark backdrop (or press Esc) to close.
    """
    import hashlib

    b64 = base64.b64encode(image_bytes).decode()
    # Unique but stable ID so the :target hash doesn't collide across renders
    uid = hashlib.md5(f"{sub_file}-{q_idx}-{s_idx}-{section}".encode()).hexdigest()[:8]
    zoom_id = f"zoom-{uid}"

    st.markdown(
        f"""<style>
#{zoom_id} {{
    display: none;
    position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(0,0,0,0.88); z-index: 99999;
    cursor: zoom-out;
}}
#{zoom_id}:target {{
    display: flex; align-items: center; justify-content: center;
}}
#{zoom_id} img {{
    max-width: 95vw; max-height: 95vh;
    border-radius: 6px; box-shadow: 0 8px 40px rgba(0,0,0,.5);
    cursor: default;
}}
#{zoom_id} .close-btn {{
    position: absolute; top: 18px; right: 28px;
    color: #fff; font-size: 36px; line-height: 1;
    text-decoration: none; cursor: pointer;
    font-weight: 300; opacity: 0.8;
}}
#{zoom_id} .close-btn:hover {{ opacity: 1; }}
</style>

<a href="#{zoom_id}" title="点击放大" style="display:inline-block;line-height:0;">
  <img src="data:image/png;base64,{b64}"
       style="max-width:100%;cursor:zoom-in;border:1px solid #e2e8f0;border-radius:6px;"
       alt="Q{q_idx}({s_idx}) {section}">
</a>

<div id="{zoom_id}">
  <a href="#" class="close-btn">&times;</a>
  <img src="data:image/png;base64,{b64}" alt="Q{q_idx}({s_idx}) {section}">
</div>""",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Main area — grading view
# ---------------------------------------------------------------------------

def render_student_answer(
    sub_file: str,
    ref_questions: Optional[List[MajorQuestion]],
):
    """Render grading UI for one student."""
    # Show navigation warning (set by navigate_student if current page has ungraded items)
    nav_warning = st.session_state.get("_nav_warning")
    if nav_warning:
        st.warning(nav_warning)
        del st.session_state["_nav_warning"]

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

    # ---- Prev / Next student navigation in main area ----
    student_names = list(st.session_state.submissions.keys())
    if len(student_names) > 1:
        idx = student_names.index(sub_file)
        nav_cols = st.columns([1, 2, 1])
        with nav_cols[0]:
            if st.button("◀ 上一人", use_container_width=True, key="main_prev_stu"):
                if idx > 0:
                    navigate_student(student_names[idx - 1])
        with nav_cols[1]:
            st.caption(f"**{idx + 1} / {len(student_names)}**")
        with nav_cols[2]:
            if st.button("下一人 ▶", use_container_width=True, key="main_next_stu"):
                if idx < len(student_names) - 1:
                    navigate_student(student_names[idx + 1])
        st.divider()

    # Progress: current question status for this student
    q_idx = st.session_state.current_tab
    graded_q_subs = sum(
        1 for s in range(1, 6)
        if scores.get(skey(q_idx + 1, s, "程序")) is not None
           and scores.get(skey(q_idx + 1, s, "结果")) is not None
    )
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric(f"第{'一二三四'[q_idx]}题 已批改", f"{graded_q_subs}/5 小题")
    with col_m2:
        st.metric(f"第{'一二三四'[q_idx]}题 未批改", f"{5 - graded_q_subs} 小题")
    with col_m3:
        if graded_q_subs > 0:
            q_total = sum(
                min(
                    (scores.get(skey(q_idx + 1, s, "程序"), 0.0) +
                     scores.get(skey(q_idx + 1, s, "结果"), 0.0)),
                    SCORE_CONFIG["sub_max"],
                )
                for s in range(1, 6)
                if scores.get(skey(q_idx + 1, s, "程序")) is not None
                   and scores.get(skey(q_idx + 1, s, "结果")) is not None
            )
            st.metric("本题得分", f"{q_total:.1f} / 25")
    st.divider()

    # Render active question
    q_idx = st.session_state.current_tab
    if q_idx < len(submission.major_questions):
        stu_q = submission.major_questions[q_idx]
        ref_q = ref_questions[q_idx] if ref_questions and q_idx < len(ref_questions) else None

        # Major question title
        q_title = stu_q.title or f"第{q_idx+1}题"
        st.subheader(f"**{q_title}**")

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

        # Bottom navigation — prev/next student
        if len(student_names) > 1:
            st.divider()
            bot_cols = st.columns([1, 2, 1])
            with bot_cols[0]:
                if st.button("◀ 上一人", use_container_width=True, key="bot_prev_stu"):
                    if idx > 0:
                        navigate_student(student_names[idx - 1])
            with bot_cols[1]:
                st.caption(f"**{idx + 1} / {len(student_names)}**")
            with bot_cols[2]:
                if st.button("下一人 ▶", use_container_width=True, key="bot_next_stu"):
                    if idx < len(student_names) - 1:
                        navigate_student(student_names[idx + 1])
    else:
        st.info("该大题无数据")


def render_sub_question(
    q_idx: int,
    s_idx: int,
    stu_q: MajorQuestion,
    ref_q: Optional[MajorQuestion],
    sub_file: str,
):
    """Render one sub-question with reference, student answer, and checkbox scoring."""
    scores = st.session_state.scores.get(sub_file, {})
    rubric_data = get_rubric_for(q_idx, s_idx)

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
    q_text = ""
    if stu_sub:
        q_text = re.sub(r"^[（(]\d+[）)]\s*", "", stu_sub.question_text)
        q_text = q_text[:80]
    sub_title = f"{status_icon} ({s_idx})　{q_text}"

    # Left accent border + main content
    accent_color = "#4ade80" if both_done else "#e2e8f0"
    accent, main_col = st.columns([0.03, 0.97])
    with accent:
        st.markdown(f'<div style="background:{accent_color};border-radius:3px;height:100%;min-height:80px;">&nbsp;</div>', unsafe_allow_html=True)
    with main_col:
        st.markdown(f"**{sub_title}**")

        # ---- 程序 section ----
        col1, col2, col3 = st.columns([4, 4, 2])
        with col1:
            st.caption("📖 参考答案 — 程序")
            if ref_sub and ref_sub.program:
                if ref_sub.program.text_content:
                    st.code(ref_sub.program.text_content, language="python", line_numbers=True)
                elif ref_sub.program.image_bytes:
                    st.image(ref_sub.program.image_bytes, width='stretch')
                elif ref_sub.program.is_missing:
                    st.info("无参考数据")
            else:
                st.info("无参考数据")

        with col2:
            st.caption("📸 学生答案 — 程序")
            if stu_sub and stu_sub.program and stu_sub.program.image_bytes:
                _render_zoomable_image(stu_sub.program.image_bytes, sub_file, q_idx, s_idx, "程序")
            else:
                st.error("❌ 未作答")

        with col3:
            render_scoring_checkboxes(sub_file, q_idx, s_idx, "程序", rubric_data)

        # ---- 结果 section ----
        col1, col2, col3 = st.columns([4, 4, 2])
        with col1:
            st.caption("📖 参考答案 — 结果")
            if ref_sub and ref_sub.result:
                if ref_sub.result.image_bytes:
                    st.image(ref_sub.result.image_bytes, width='stretch')
                elif ref_sub.result.text_content:
                    st.code(ref_sub.result.text_content, language="text")
                elif ref_sub.result.is_missing:
                    st.info("无参考数据")
            else:
                st.info("无参考数据")

        with col2:
            st.caption("📸 学生答案 — 结果")
            if stu_sub and stu_sub.result and stu_sub.result.image_bytes:
                _render_zoomable_image(stu_sub.result.image_bytes, sub_file, q_idx, s_idx, "结果")
            else:
                st.error("❌ 未作答")

        with col3:
            render_scoring_checkboxes(sub_file, q_idx, s_idx, "结果", rubric_data)

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
        /* Score radio buttons as larger clickable chips */
        div[data-testid="stRadio"] > div {
            gap: 0.3rem !important;
            flex-wrap: wrap;
        }
        div[data-testid="stRadio"] label {
            border: 1px solid #d0d5dd;
            border-radius: 8px;
            padding: 0.3rem 0.65rem;
            margin: 0.15rem 0.1rem !important;
            background: #f8f9fa;
            font-size: 1rem !important;
            font-weight: 500;
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
            font-weight: 700;
        }
        /* Compact checkboxes in scoring column */
        div[data-testid="stCheckbox"] label {
            font-size: 0.82rem !important;
            padding: 0.15rem 0 !important;
            gap: 0.3rem !important;
        }
        div[data-testid="stCheckbox"] {
            min-height: 1.6rem !important;
        }
        /* Score label styling */
        .score-label {
            font-size: 1rem;
            font-weight: 600;
            margin-bottom: 0.2rem;
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
            1. 上传 **参考答案**（.docx 或 .ipynb 格式）
            2. 上传 **评分标准 .json** 文件
            3. 上传 **学生答卷 .docx** 文件（可多选）
            4. 勾选评分标准中的得分点进行批改
            5. 点击 **导出成绩** 生成 xlsx 文件
            """)
            return

    # Show auto-advance notification (persists across rerun)
    auto_msg = st.session_state.get("_auto_advance")
    if auto_msg:
        st.success(auto_msg)
        del st.session_state["_auto_advance"]

    render_student_answer(current, ref_qs)


if __name__ == "__main__":
    main()
