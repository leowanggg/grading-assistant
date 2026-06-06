"""
IPYNB Parser — Extract reference answer structure from Jupyter Notebook (.ipynb) files.

Convention for reference notebooks:
  - Markdown cells with # 一/二/三/四 headers → major questions
  - Markdown cells with (N) or N. markers → sub-questions
  - Code cells after a sub-question → 程序 (as formatted code text)
  - Outputs of code cells → 结果 (as text and/or images)

Example structure:
  # 一、Python数据分析与可视化
  ## (1) 创建销售数据数组
  ```python
  import numpy as np
  np.random.seed(2)
  sales = np.random.randint(30, 101, size=(5, 3))
  sales
  ```
  [output → 结果]
  ## (2) 转换为DataFrame...
  ...
  # 二、数据预处理（heart.csv）
  ...
"""

from __future__ import annotations

import base64
import json
import re
from typing import Dict, List, Optional

from docx_parser import ImageSection, MajorQuestion, SubQuestion


# ---------------------------------------------------------------------------
# Notebook cell helpers
# ---------------------------------------------------------------------------

def _get_source_text(cell: dict) -> str:
    """Get the full text of a cell's source (handles list-of-strings or single string)."""
    source = cell.get("source", [])
    if isinstance(source, list):
        return "".join(source)
    return str(source)


def _get_output_text(output: dict) -> str:
    """Get plain-text representation of one cell output."""
    output_type = output.get("output_type", "")
    if output_type == "stream":
        text = output.get("text", [])
    elif output_type in ("execute_result", "display_data"):
        data = output.get("data", {})
        text = data.get("text/plain", [])
    else:
        return ""

    if isinstance(text, list):
        return "".join(text)
    return str(text)


def _get_output_image(output: dict) -> Optional[bytes]:
    """Extract the first PNG/JPEG image from a cell output, if any."""
    data = output.get("data", {})
    for mime in ("image/png", "image/jpeg", "image/gif"):
        b64_data = data.get(mime)
        if b64_data:
            raw = b64_data if isinstance(b64_data, str) else "".join(b64_data)
            try:
                return base64.b64decode(raw)
            except Exception:
                return None
    return None


# ---------------------------------------------------------------------------
# Heading detection
# ---------------------------------------------------------------------------

# Match "一、", "二.", "三．" etc. at start after optional # characters
_MAJOR_HEADER_RE = re.compile(
    r"^#*\s*([一二三四])\s*[、.．\s]"
)

# Match "(1)", "（1）" or leading "1." / "1、" in heading text
_SUB_HEADER_RE = re.compile(r"[（(](\d)[）)]")
_SUB_HEADER_RE2 = re.compile(r"^#*\s*(\d)\s*[.、．]")


def _is_major_heading(text: str) -> Optional[tuple[int, str]]:
    """Scan text lines for a major-question heading.
    Returns (1-based index, cleaned_title) or None.
    Handles headings embedded in cells with intro text or separators.
    """
    for line in text.split('\n'):
        stripped = line.strip()
        m = _MAJOR_HEADER_RE.match(stripped)
        if m:
            cn = m.group(1)
            idx = {"一": 1, "二": 2, "三": 3, "四": 4}.get(cn)
            if idx:
                # Strip leading # markers for the title
                title = re.sub(r'^#+\s*', '', stripped).strip()
                return (idx, title)
    return None


def _extract_sub_index(text: str) -> Optional[int]:
    """Try to extract a sub-question index (1-5) from heading text.
    Scans each line. Prioritizes line-starting "N." / "N、" patterns over
    parenthesized numbers to avoid false matches like （5分）→5.
    """
    for line in text.split('\n'):
        stripped = line.strip()
        # Priority: "1." or "1、" at line start
        m = _SUB_HEADER_RE2.match(stripped)
        if m:
            return int(m.group(1))
        # Fallback: "(1)" or "（1）" anywhere in line
        m = _SUB_HEADER_RE.search(stripped)
        if m:
            return int(m.group(1))
    return None


# ---------------------------------------------------------------------------
# Clean heading text for display
# ---------------------------------------------------------------------------

def _strip_heading(text: str) -> str:
    """Remove leading # characters and trim."""
    return text.lstrip("#").strip()


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_reference_ipynb(file_path: str) -> List[MajorQuestion]:
    """
    Parse a Jupyter Notebook (.ipynb) reference answer file.

    Returns a list of up to 4 MajorQuestion objects, each with 5 SubQuestions.

    The notebook is expected to follow this convention:
      - `# 一、Title`  → Major Question 1
      - `## (1) Title` → Sub-Question 1
      - Code cell      → 程序 (program code as text)
      - Output cell    → 结果 (result text and/or images)
    """
    with open(file_path, "r", encoding="utf-8") as f:
        nb = json.load(f)

    cells = nb.get("cells", [])

    questions: List[MajorQuestion] = []
    cur_q: Optional[MajorQuestion] = None
    cur_sub: Optional[SubQuestion] = None
    awaiting_prog = False  # True when we expect the next code cell to be 程序
    prog_texts: List[str] = []
    output_texts: List[str] = []
    output_image: Optional[bytes] = None

    def _flush_code_cell():
        """Attach buffered code & outputs to the current sub-question."""
        nonlocal prog_texts, output_texts, output_image
        if cur_sub is None:
            prog_texts, output_texts, output_image = [], [], None
            return

        if prog_texts:
            code = "\n".join(prog_texts)
            cur_sub.program = ImageSection(
                section_type="程序",
                text_content=code,
                is_missing=False,
            )

        if output_texts or output_image:
            text = "\n".join(output_texts).strip()
            cur_sub.result = ImageSection(
                section_type="结果",
                text_content=text,
                image_bytes=output_image,
                is_missing=not text and output_image is None,
            )

        prog_texts, output_texts, output_image = [], [], None

    for cell in cells:
        cell_type = cell.get("cell_type", "")
        source = _get_source_text(cell).strip()

        if cell_type == "markdown":
            if not source:
                continue

            # ---- Major question detection (scan lines) ----
            major = _is_major_heading(source)
            if major is not None:
                _flush_code_cell()
                if len(questions) >= 4:
                    cur_q = None
                    cur_sub = None
                    continue
                q_idx, q_title = major
                cur_q = MajorQuestion(index=q_idx, title=q_title)
                questions.append(cur_q)
                cur_sub = None
                awaiting_prog = True
                continue

            # ---- Sub-question detection ----
            sub_idx = _extract_sub_index(source)
            if sub_idx is not None and cur_q is not None:
                _flush_code_cell()
                if 1 <= sub_idx <= 5:
                    cur_sub = SubQuestion(index=sub_idx, question_text=_strip_heading(source))
                    cur_q.sub_questions.append(cur_sub)
                    awaiting_prog = True
                continue

            # ---- Description text for current sub-question ----
            if cur_sub is not None and source:
                cur_sub.question_text += "\n" + source

        elif cell_type == "code":
            code_text = _get_source_text(cell).strip()

            if cur_sub is None:
                # Code cell before any sub-question — skip unless we can attach to last sub
                continue

            # Flush any previous code block that wasn't attached yet
            if awaiting_prog and prog_texts:
                _flush_code_cell()

            # Store code as program text
            prog_texts.append(code_text)

            # Process outputs
            outputs = cell.get("outputs", [])
            for out in outputs:
                out_type = out.get("output_type", "")
                # Skip error outputs
                if out_type == "error":
                    prog_texts.append(f"# [错误]: {_get_output_text(out)[:200]}")
                    continue

                img = _get_output_image(out)
                if img and output_image is None:
                    output_image = img

                txt = _get_output_text(out)
                if txt:
                    output_texts.append(txt)

            # For error outputs with no text lines, note it
            if not output_texts and not output_image:
                has_error = any(o.get("output_type") == "error" for o in outputs)
                if has_error:
                    output_texts.append("[运行错误]")

            # If there's only one code cell per question, flush immediately
            # (but buffer first so multiple code cells are combined)
            awaiting_prog = False

    # Flush any remaining buffered code
    _flush_code_cell()

    # ---- Ensure 4 major questions with 5 sub-questions each ----
    seen_major_indices = {q.index for q in questions}

    # Sort questions by index
    questions.sort(key=lambda q: q.index)

    # Fill missing major questions
    for i in range(1, 5):
        if i not in seen_major_indices:
            q = MajorQuestion(index=i, title=f"第{'一二三四'[i-1]}题")
            questions.append(q)

    questions.sort(key=lambda q: q.index)

    # Fill missing sub-questions for each major
    for q in questions:
        existing = {s.index for s in q.sub_questions}
        for i in range(1, 6):
            if i not in existing:
                q.sub_questions.append(
                    SubQuestion(index=i, question_text=f"({i})")
                )
        q.sub_questions.sort(key=lambda s: s.index)

    return questions


def question_signature_ipynb(questions: List[MajorQuestion]) -> str:
    """
    Generate a signature for a parsed notebook (same format as docx_parser.question_signature).
    Reuses the docx_parser function for consistency.
    """
    # Import here to avoid circular dependency
    from docx_parser import question_signature
    return question_signature(questions)


# ---------------------------------------------------------------------------
# Quick test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else r"eg/reference.ipynb"
    qs = parse_reference_ipynb(path)
    print(f"Questions: {len(qs)}")
    for q in qs:
        print(f"  Q{q.index}: {q.title[:60]}")
        for s in q.sub_questions:
            prog_note = "代码"
            if s.program:
                prog_note = f"代码({len(s.program.text_content)}字符)" if s.program.text_content else "无"
            res_note = "输出"
            if s.result:
                if s.result.text_content:
                    res_note = f"文本({len(s.result.text_content)}字符)"
                elif s.result.image_bytes:
                    res_note = f"图片({len(s.result.image_bytes)}字节)"
                else:
                    res_note = "无"
            print(f"    ({s.index}) 程序:{prog_note} 结果:{res_note}")
