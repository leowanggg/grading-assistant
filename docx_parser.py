"""
DOCX Parser — Extract exam answers from student .docx files.

Structure:
  Cover page (student info) → exam header → 4 major questions,
  each with 5 sub-questions, each having 程序 (program) and 结果 (result) sections
  answered as screenshot images.
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from typing import List, Optional, Tuple
from xml.etree import ElementTree as ET

from docx import Document
from docx.image.image import Image as DocxImage
from docx.opc.constants import RELATIONSHIP_TYPE as RT


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class ImageSection:
    """One scorable section: 程序 (program) or 结果 (result)."""
    section_type: str               # "程序" or "结果"
    image_bytes: Optional[bytes] = None
    text_content: str = ""
    is_missing: bool = False


@dataclass
class SubQuestion:
    """Single sub-question (1-5) inside a major question."""
    index: int                      # 1-based
    question_text: str = ""
    program: Optional[ImageSection] = None
    result: Optional[ImageSection] = None


@dataclass
class MajorQuestion:
    """One major question (一 to 四)."""
    index: int                      # 1-based
    title: str = ""
    sub_questions: List[SubQuestion] = field(default_factory=list)


@dataclass
class StudentInfo:
    """Student identity from cover page."""
    student_id: str = ""
    name: str = ""
    class_name: str = ""


@dataclass
class StudentSubmission:
    """Parsed student answer document."""
    filename: str
    student_info: StudentInfo = field(default_factory=StudentInfo)
    major_questions: List[MajorQuestion] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Namespace helpers
# ---------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS  = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

WP_INLINE   = f"{{{WP_NS}}}inline"
WP_ANCHOR   = f"{{{WP_NS}}}anchor"
A_BLIP      = f"{{{A_NS}}}blip"
R_EMBED     = f"{{{R_NS}}}embed"
W_DRAWING   = f"{{{W_NS}}}drawing"


# ---------------------------------------------------------------------------
# Image extraction
# ---------------------------------------------------------------------------

def get_image_bytes(para, document: Document) -> Optional[bytes]:
    """Extract the first inline or floating image from a paragraph, return raw bytes."""
    element = para._element

    # Search for wp:inline > a:blip
    container = element.find(f".//{WP_INLINE}")
    # Fallback: search for wp:anchor > a:blip (floating images)
    if container is None:
        container = element.find(f".//{WP_ANCHOR}")
    # Fallback: search for w:drawing > wp:inline
    if container is None:
        drawing = element.find(f".//{W_DRAWING}")
        if drawing is not None:
            container = drawing.find(f".//{WP_INLINE}")
    if container is None:
        return None

    blip = container.find(f".//{A_BLIP}")
    if blip is None:
        return None

    r_id = blip.get(R_EMBED)
    if r_id is None:
        return None

    try:
        rel = document.part.rels[r_id]
        return rel.target_part.blob
    except (KeyError, AttributeError):
        return None


def has_image(para) -> bool:
    """Quick check if a paragraph contains an inline image."""
    element = para._element
    if element.find(f".//{WP_INLINE}") is not None:
        return True
    return element.find(f".//{W_DRAWING}") is not None


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def para_text(para) -> str:
    """Normalized paragraph text (spaces stripped)."""
    return para.text.strip() if para.text else ""


# ---------------------------------------------------------------------------
# Student info extraction (cover page)
# ---------------------------------------------------------------------------

def extract_student_info(doc: Document) -> StudentInfo:
    """Extract 学号/姓名/专业班级 from cover page (first ~15 paragraphs)."""
    info = StudentInfo()

    # Strategy 1 — scan paragraphs for label: value patterns
    pat_id    = re.compile(r"学\s*号\s*[：:]\s*(.+)")
    pat_name  = re.compile(r"姓\s*名\s*[：:]\s*(.+)")
    pat_class = re.compile(r"专业班级\s*[：:]\s*(.+)")

    # Also handle underline-only value after label in same paragraph
    pat_id2    = re.compile(r"学\s*号")
    pat_name2  = re.compile(r"姓\s*名")
    pat_class2 = re.compile(r"专业班级")

    for para in doc.paragraphs[:20]:
        text = para_text(para)
        if not text:
            continue

        m = pat_id.search(text)
        if m and not info.student_id:
            info.student_id = m.group(1).strip()
            continue
        m = pat_name.search(text)
        if m and not info.name:
            info.name = m.group(1).strip()
            continue
        m = pat_class.search(text)
        if m and not info.class_name:
            info.class_name = m.group(1).strip()
            continue

        # Second pass — look for label and grab last non-empty run
        runs = [r.text for r in para.runs if r.text and r.text.strip()]
        joined = "".join(runs).strip()
        if pat_id2.search(text) and not info.student_id:
            # Value is everything after the label
            idx = text.find("号")
            if idx >= 0:
                candidate = text[idx+1:].strip().lstrip("：:").strip()
                if candidate:
                    info.student_id = candidate
        if pat_name2.search(text) and not info.name:
            idx = text.find("名")
            if idx >= 0:
                candidate = text[idx+1:].strip().lstrip("：:").strip()
                if candidate:
                    info.name = candidate
        if pat_class2.search(text) and not info.class_name:
            idx = text.find("班级")
            if idx >= 0:
                candidate = text[idx+2:].strip().lstrip("：:").strip()
                if candidate:
                    info.class_name = candidate

    # Strategy 2 — check first table if no info found
    if not info.student_id or not info.name:
        for table in doc.tables:
            for row in table.rows:
                row_text = "".join(cell.text for cell in row.cells)
                m = pat_id.search(row_text)
                if m:
                    info.student_id = m.group(1).strip()
                m = pat_name.search(row_text)
                if m:
                    info.name = m.group(1).strip()
                m = pat_class.search(row_text)
                if m:
                    info.class_name = m.group(1).strip()
            if info.name:
                break

    # Strategy 3 — regex fallback over all first-page text
    if not info.student_id:
        all_text = " ".join(para_text(p) for p in doc.paragraphs[:20])
        nums = re.findall(r"\b(\d{8,12})\b", all_text)
        if nums:
            info.student_id = nums[0]
    if not info.name:
        all_text = " ".join(para_text(p) for p in doc.paragraphs[:20])
        names = re.findall(r"姓\s*名\s*[：:]?\s*([^\s]{2,4})", all_text)
        if not names:
            names = re.findall(r"([一-鿿]{2,4})", all_text)
            # Filter out common non-name words
            skip = {"实践考核", "须知", "数据挖掘", "课程考核", "专业班级", "任课教师"}
            names = [n for n in names if n not in skip and len(n) >= 2]
        if names:
            info.name = names[0]

    return info


# ---------------------------------------------------------------------------
# Question structure parsing
# ---------------------------------------------------------------------------

_MAJOR_RE = re.compile(r"^(?:[一二三四]|\d+)\s*[.、].{4,}")  # require meaningful content after the number
_SUB_RE   = re.compile(r"^[（(](\d+)[）)]")
_PROG_RE  = re.compile(r"^程序[：:]")
_RESULT_RE = re.compile(r"^结果[：:]")
_EMPTY_PAT = re.compile(r"^\s*$")


def parse_questions(doc: Document) -> List[MajorQuestion]:
    """
    Core state-machine parser.

    States: SEEK_MAJOR → IN_MAJOR(accumulate title) → IN_SUB →
            (on 程序/结果 label) → capture next image/text → back to IN_SUB
    """
    questions: List[MajorQuestion] = []
    cur_q: Optional[MajorQuestion] = None
    cur_sub: Optional[SubQuestion] = None
    state = "SEEK_MAJOR"
    # buffer for sub-question question_text
    sub_text_buffer: List[str] = []
    # track pending section to fill
    pending_section: Optional[str] = None  # "程序" or "结果"

    for para in doc.paragraphs:
        text = para_text(para)

        # ---------- SEEK_MAJOR ----------
        if state == "SEEK_MAJOR":
            if _MAJOR_RE.search(text):
                if len(questions) >= 4:
                    continue  # ignore extra matches past 4 questions
                cur_q = MajorQuestion(index=len(questions) + 1, title=text)
                questions.append(cur_q)
                state = "IN_SUB"
                sub_text_buffer = []
            continue

        # ---------- IN_MAJOR (accumulate title) ----------
        if state == "IN_MAJOR":
            if _MAJOR_RE.search(text):
                cur_q.title = text
                state = "IN_SUB"
                sub_text_buffer = []
            elif _SUB_RE.search(text):
                # Sub-question found — save title buffer as title
                cur_q.title = " ".join(sub_text_buffer) if sub_text_buffer else text
                sub_text_buffer = []
                m = _SUB_RE.match(text)
                sub_idx = int(m.group(1))
                cur_sub = SubQuestion(index=sub_idx, question_text=text)
                cur_q.sub_questions.append(cur_sub)
                state = "IN_SUB"
            else:
                sub_text_buffer.append(text)
            continue

        # ---------- IN_SUB ----------
        if state == "IN_SUB":
            # Check for next major question
            if _MAJOR_RE.search(text):
                if len(questions) >= 4:
                    continue  # ignore extra matches past 4 questions
                cur_q = MajorQuestion(index=len(questions) + 1, title=text)
                questions.append(cur_q)
                state = "IN_SUB"
                sub_text_buffer = []
                cur_sub = None
                continue

            # Check for sub-question marker
            m = _SUB_RE.match(text)
            if m:
                sub_idx = int(m.group(1))
                cur_sub = SubQuestion(index=sub_idx, question_text=text)
                cur_q.sub_questions.append(cur_sub)
                continue

            # Check for 程序 label
            if _PROG_RE.match(text):
                # If image is in the SAME paragraph as the label, capture immediately
                img_bytes = get_image_bytes(para, doc)
                if img_bytes is not None:
                    if cur_sub is not None:
                        cur_sub.program = ImageSection(section_type="程序", image_bytes=img_bytes)
                    continue
                pending_section = "程序"
                state = "CAPTURE_IMAGE"
                continue

            # Check for 结果 label
            if _RESULT_RE.match(text):
                # If image is in the SAME paragraph as the label, capture immediately
                img_bytes = get_image_bytes(para, doc)
                if img_bytes is not None:
                    if cur_sub is not None:
                        cur_sub.result = ImageSection(section_type="结果", image_bytes=img_bytes)
                    continue
                pending_section = "结果"
                state = "CAPTURE_IMAGE"
                continue

            # Otherwise accumulate as question description for current sub
            if cur_sub is not None and text:
                cur_sub.question_text += "\n" + text
            continue

        # ---------- CAPTURE_IMAGE ----------
        if state == "CAPTURE_IMAGE":
            # Check if a new sub-question starts before we captured anything
            if _SUB_RE.match(text):
                # Mark current section as missing, close it, start new sub
                section = ImageSection(
                    section_type=pending_section, is_missing=True
                )
                if cur_sub is not None:
                    if pending_section == "程序":
                        cur_sub.program = section
                    elif pending_section == "结果":
                        cur_sub.result = section
                pending_section = None
                # Now process this paragraph as a new sub-question
                m = _SUB_RE.match(text)
                sub_idx = int(m.group(1))
                cur_sub = SubQuestion(index=sub_idx, question_text=text)
                cur_q.sub_questions.append(cur_sub)
                state = "IN_SUB"
                continue

            # Try to get image from this paragraph
            img_bytes = get_image_bytes(para, doc)
            section = ImageSection(section_type=pending_section)

            if img_bytes is not None:
                section.image_bytes = img_bytes
            elif _EMPTY_PAT.match(text) or not text:
                # Empty paragraph — could be blank line before image
                continue
            elif text:
                section.text_content = text

            if section.image_bytes is None and not section.text_content:
                section.is_missing = True

            # Attach to current sub-question
            if cur_sub is not None:
                if pending_section == "程序":
                    cur_sub.program = section
                elif pending_section == "结果":
                    cur_sub.result = section

            pending_section = None
            state = "IN_SUB"
            continue

    return questions


# ---------------------------------------------------------------------------
# Question signature for variant matching
# ---------------------------------------------------------------------------

def question_signature(questions: List[MajorQuestion]) -> str:
    """Short hash of question titles for variant detection."""
    parts = []
    for q in questions:
        clean = re.sub(r"\s+", "", q.title)[:60]
        parts.append(clean)
    return "|".join(parts)


# ---------------------------------------------------------------------------
# Public parsing API
# ---------------------------------------------------------------------------

def parse_student_docx(file_path: str) -> StudentSubmission:
    """
    Parse a student .docx file.
    Returns StudentSubmission with info, questions, and answer images.
    """
    doc = Document(file_path)
    info = extract_student_info(doc)
    questions = parse_questions(doc)

    # Ensure 4 major questions with 5 sub-questions each
    for q_idx, q in enumerate(questions):
        # Fill missing sub-questions with placeholders
        existing = {s.index for s in q.sub_questions}
        for i in range(1, 6):
            if i not in existing:
                q.sub_questions.append(
                    SubQuestion(index=i, question_text=f"({i})")

                )
        q.sub_questions.sort(key=lambda s: s.index)

    return StudentSubmission(
        filename=file_path,
        student_info=info,
        major_questions=questions,
    )


def parse_reference_docx(file_path: str) -> List[MajorQuestion]:
    """
    Parse a reference answer .docx.
    Returns list of MajorQuestion (no student info).
    """
    doc = Document(file_path)
    return parse_questions(doc)


# ---------------------------------------------------------------------------
# Quick test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else r"c:\Users\Silin WANG\Projects\test\ansr.docx"
    sub = parse_student_docx(path)
    print(f"Student: {sub.student_info.name} / {sub.student_info.student_id}")
    print(f"Questions: {len(sub.major_questions)}")
    for q in sub.major_questions:
        print(f"  Q{q.index}: {q.title[:60]}")
        for s in q.sub_questions:
            prog = "✓" if s.program and s.program.image_bytes else ("✗" if s.program and s.program.is_missing else "?")
            res  = "✓" if s.result and s.result.image_bytes else ("✗" if s.result and s.result.is_missing else "?")
            print(f"    ({s.index}) 程序:{prog} 结果:{res}")
