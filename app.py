"""
Voice Marks → Excel  (Faculty Dictation App)
=============================================
Faculty speaks while checking copies (wired earphone mic on Android phone):
  "Name Pratham, roll number 2K23CSUN001,
   question 1 A 5 marks, question 1 B 3 marks,
   question 2 zero, question 5 four marks …"

App parses the speech/text → structured table → Excel with auto-sum formulas.

Works best on Android: open app → tap text box → use Gboard voice typing.
Also accepts audio file upload (optional OpenAI Whisper if API key given).
"""

import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Voice Marks → Excel",
    page_icon="🎤",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "students" not in st.session_state:
    st.session_state.students = []
if "last_transcript" not in st.session_state:
    st.session_state.last_transcript = ""


# ------------------------------------------------------------------
# Parser: spoken / typed marks script → structured dict
# ------------------------------------------------------------------
def normalize_transcript(text: str) -> str:
    text = text.lower()
    # Common spoken number words
    num_map = {
        "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4",
        "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9",
        "ten": "10", "eleven": "11", "twelve": "12", "thirteen": "13",
        "fourteen": "14", "fifteen": "15", "sixteen": "16", "seventeen": "17",
        "eighteen": "18", "nineteen": "19", "twenty": "20",
        "oh": "0", "o": "0",
    }
    for w, d in num_map.items():
        text = re.sub(rf"\b{w}\b", d, text)

    # Spoken letters for parts
    letter_map = {
        "ay": "a", "a.": "a", "bee": "b", "b.": "b", "see": "c", "c.": "c",
        "dee": "d", "d.": "d", "ee": "e", "e.": "e",
    }
    for w, d in letter_map.items():
        text = re.sub(rf"\b{w}\b", d, text)

    text = text.replace("roll no.", "roll number")
    text = text.replace("roll no", "roll number")
    text = text.replace("q.", "question ")
    text = text.replace("ques ", "question ")
    text = text.replace("question no.", "question ")
    text = text.replace("question number", "question ")
    text = text.replace("marks", " marks ")
    text = text.replace("mark ", " marks ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def extract_roll(text: str) -> str:
    # Patterns like 2k23csun001, 2K23UON0104, 2k23 csun 001
    patterns = [
        r"\b(2[kK]\d{2}[a-zA-Z]{2,6}\d{3,6})\b",
        r"\b(2\s*[kK]\s*\d{2}\s*[a-zA-Z]{2,6}\s*\d{3,6})\b",
        r"roll\s*(?:number|no\.?)?\s*[:\-]?\s*([0-9a-zA-Z]{8,16})",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            roll = re.sub(r"\s+", "", m.group(1)).upper()
            return roll
    return ""


def extract_name(text: str) -> str:
    patterns = [
        r"(?:student\s*)?name\s*(?:is|:)?\s*([a-zA-Z][a-zA-Z\s]{1,30}?)(?:\s*,|\s+roll|\s+question|\s+q\d|$)",
        r"(?:name)\s+([a-zA-Z][a-zA-Z\s]{1,25}?)(?:\s+roll)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            name = re.sub(r"[^a-zA-Z\s]", "", m.group(1)).strip()
            name = re.sub(r"\s+", " ", name)
            if len(name) >= 2 and name.lower() not in ("is", "the", "student"):
                return name.title()
    return ""


def extract_marks(text: str) -> dict:
    """
    Returns dict like {"q1a": 5, "q1b": 3, "q2": 0, "q5": 4, ...}
    """
    marks = {}
    t = text.lower()

    # Pattern set 1: "question 1 a 5 marks" / "question 1 (a) 5"
    patterns = [
        # question 1 a 5 / question 1 (a) 5 marks
        r"question\s*(\d{1,2})\s*[\(\[]?\s*([a-j])\s*[\)\]]?\s*[:\-]?\s*(\d{1,2})\b",
        # q1a 5 / q1-a 5 / q1 (a) 5
        r"\bq\s*(\d{1,2})\s*[\-\s\(\[]*\s*([a-j])\s*[\)\]]*\s*[:\-]?\s*(\d{1,2})\b",
        # 1 a 5 marks / 1.(a) 5
        r"\b(\d{1,2})\s*[\.\)]\s*[\(\[]?\s*([a-j])\s*[\)\]]?\s*[:\-]?\s*(\d{1,2})\b",
    ]
    for pat in patterns:
        for m in re.finditer(pat, t):
            q, part, val = m.group(1), m.group(2).lower(), int(m.group(3))
            key = f"q{int(q)}{part}"
            marks[key] = val

    # Pattern set 2: question without part – "question 2 0 marks" / "q5 4"
    patterns2 = [
        r"question\s*(\d{1,2})\s*[:\-]?\s*(\d{1,2})\s*(?:marks?)?\b",
        r"\bq\s*(\d{1,2})\s*[:\-]?\s*(\d{1,2})\b",
    ]
    for pat in patterns2:
        for m in re.finditer(pat, t):
            q, val = int(m.group(1)), int(m.group(2))
            # Only if no part already captured for this question
            if not any(k.startswith(f"q{q}") and len(k) > len(f"q{q}") for k in marks):
                marks[f"q{q}"] = val

    # "question 2 zero" already normalized to digit
    return marks


def parse_transcript(raw: str) -> dict:
    text = normalize_transcript(raw)
    name = extract_name(text)
    roll = extract_roll(text)
    marks = extract_marks(text)

    # Build full student record
    student = {
        "name": name,
        "roll": roll,
        "raw": raw.strip(),
    }
    # Initialize all possible keys
    for q in range(1, 11):
        for p in list("abcdefghij") + [""]:
            key = f"q{q}{p}" if p else f"q{q}"
            student[key] = marks.get(key)

    # Prefer part-wise; if only q1 exists and no q1a, keep q1
    # Calculate totals
    grand = 0
    for q in range(1, 11):
        q_total = 0
        has_parts = False
        for p in "abcdefghij":
            v = student.get(f"q{q}{p}")
            if v is not None:
                q_total += int(v)
                has_parts = True
        if not has_parts:
            v = student.get(f"q{q}")
            if v is not None:
                q_total = int(v)
        student[f"q{q}_total"] = q_total
        grand += q_total
    student["grand_total"] = grand
    return student


# ------------------------------------------------------------------
# Excel export
# ------------------------------------------------------------------
def create_excel(students: list) -> bytes:
    wb = Workbook()
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    green_fill = PatternFill("solid", fgColor="D5F5E3")
    grand_fill = PatternFill("solid", fgColor="F5B7B1")
    yellow_fill = PatternFill("solid", fgColor="FCF3CF")
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    center = Alignment(horizontal="center", vertical="center")

    # ---- Summary ----
    ws = wb.active
    ws.title = "Summary"
    headers = ["S.No", "Student Name", "Roll Number", "Grand Total"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin
    for i, s in enumerate(students, 1):
        r = i + 1
        ws.cell(row=r, column=1, value=i).border = thin
        ws.cell(row=r, column=2, value=s.get("name", "")).border = thin
        ws.cell(row=r, column=3, value=s.get("roll", "")).border = thin
        cell = ws.cell(row=r, column=4, value=s.get("grand_total", 0))
        cell.fill = grand_fill
        cell.border = thin
        cell.alignment = center
        cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 12

    # ---- Flat sheet (one row per student, all parts) ----
    ws2 = wb.create_sheet("Student Marks (Flat)")
    flat_headers = ["S.No", "Student Name", "Roll Number"]
    for q in range(1, 11):
        for p in "abcdefghij":
            flat_headers.append(f"{q}({p})")
        flat_headers.append(f"Q{q} Total")
    flat_headers.append("Grand Total")

    for col, h in enumerate(flat_headers, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin

    for i, s in enumerate(students, 1):
        r = i + 1
        col = 1
        ws2.cell(row=r, column=col, value=i).border = thin
        col += 1
        ws2.cell(row=r, column=col, value=s.get("name", "")).border = thin
        col += 1
        ws2.cell(row=r, column=col, value=s.get("roll", "")).border = thin
        col += 1

        for q in range(1, 11):
            part_start = col
            for p in "abcdefghij":
                val = s.get(f"q{q}{p}")
                # fallback to whole-question mark on part 'a' only if no parts
                if val is None and p == "a" and s.get(f"q{q}") is not None:
                    # if only qN was spoken without part, put it in Total via formula still
                    pass
                cell = ws2.cell(row=r, column=col, value=val if val is not None else "")
                cell.border = thin
                cell.alignment = center
                if val is not None:
                    cell.fill = yellow_fill
                col += 1
            # If no parts but qN exists, put value in Total column as number
            q_only = s.get(f"q{q}")
            has_any_part = any(s.get(f"q{q}{p}") is not None for p in "abcdefghij")
            if not has_any_part and q_only is not None:
                cell = ws2.cell(row=r, column=col, value=int(q_only))
            else:
                start_l = get_column_letter(part_start)
                end_l = get_column_letter(col - 1)
                cell = ws2.cell(row=r, column=col, value=f"=SUM({start_l}{r}:{end_l}{r})")
            cell.fill = green_fill
            cell.border = thin
            cell.alignment = center
            col += 1

        # Grand total formula
        total_cols = []
        c = 4
        for q in range(1, 11):
            total_cols.append(get_column_letter(c + 10))
            c += 11
        formula = "=" + "+".join(f"{tc}{r}" for tc in total_cols)
        cell = ws2.cell(row=r, column=col, value=formula)
        cell.fill = grand_fill
        cell.border = thin
        cell.alignment = center
        cell.font = Font(bold=True)

    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 16
    for col in range(4, len(flat_headers) + 1):
        ws2.column_dimensions[get_column_letter(col)].width = 7
    ws2.freeze_panes = "A2"

    # ---- Info ----
    ws3 = wb.create_sheet("Info")
    ws3["A1"] = "Voice Marks → Excel (Faculty Dictation App)"
    ws3["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws3["A3"] = f"Students: {len(students)}"
    ws3["A5"] = "Faculty spoke marks while checking; app parsed Name, Roll, Question marks."
    ws3["A6"] = "Yellow = spoken marks | Green = question total (formula) | Red = Grand Total (formula)"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ------------------------------------------------------------------
# Sidebar – how to speak
# ------------------------------------------------------------------
with st.sidebar:
    st.title("🎤 Voice Marks → Excel")
    st.markdown("---")
    st.markdown(
        """
**How faculty uses it (Android)**

1. Open this app on phone  
2. Tap the big text box  
3. Tap **mic on Gboard** (Google voice typing)  
4. Speak while checking the copy:

```
Name Rahul Sharma
Roll number 2K23CSUN001
Question 1 A 5 marks
Question 1 B 3 marks
Question 2 zero
Question 5 4 marks
Question 7 2 marks
```

5. Tap **Parse & Add Student**  
6. Repeat for next student  
7. Download Excel when done
        """
    )
    st.markdown("---")
    st.markdown("**Speaking tips**")
    st.caption(
        "• Say “question 1 A 5 marks” clearly\n"
        "• Use “zero” for 0\n"
        "• One student per dictation works best\n"
        "• You can edit the text before parsing"
    )
    st.markdown("---")
    st.info(
        "No special mic app needed. "
        "Wired earphone mic + Android Gboard voice typing is enough."
    )


# ------------------------------------------------------------------
# Main UI
# ------------------------------------------------------------------
st.title("🎤 Voice Marks → Excel")
st.caption(
    "Speak marks while checking copies (earphone mic on Android). "
    "App extracts Name, Roll Number & marks → Excel with auto-sum."
)

st.subheader("1. Dictate or paste the marks script")

example = (
    "Name Pratham, roll number 2K23UON0104, "
    "question 1 A 4 marks, question 1 B 4 marks, question 1 C 2 marks, "
    "question 2 zero, question 4 zero, question 5 4 marks, "
    "question 7 2 marks, question 10 4 marks"
)

transcript = st.text_area(
    "Speak here (Android: tap mic on keyboard) or paste transcript",
    value=st.session_state.last_transcript,
    height=140,
    placeholder=example,
    help="On Android Chrome/Firefox: tap the box → tap keyboard mic → speak.",
)

col_a, col_b, col_c = st.columns([1, 1, 1])
with col_a:
    parse_clicked = st.button("🔍 Parse & Add Student", type="primary", use_container_width=True)
with col_b:
    if st.button("📋 Load example", use_container_width=True):
        st.session_state.last_transcript = example
        st.rerun()
with col_c:
    if st.button("🗑️ Clear text", use_container_width=True):
        st.session_state.last_transcript = ""
        st.rerun()

if parse_clicked:
    if not transcript.strip():
        st.warning("Please speak or type the marks script first.")
    else:
        student = parse_transcript(transcript)
        st.session_state.last_transcript = transcript

        if not student["name"] and not student["roll"]:
            st.error(
                "Could not find Name or Roll Number. "
                "Try: “Name Rahul Sharma roll number 2K23CSUN001 …”"
            )
        else:
            # Avoid exact duplicate roll
            roll = student["roll"]
            st.session_state.students = [
                s for s in st.session_state.students if s.get("roll") != roll or not roll
            ]
            st.session_state.students.append(student)
            st.success(
                f"Added **{student['name'] or 'Unknown'}** "
                f"({student['roll'] or 'no roll'}) — "
                f"Grand Total: **{student['grand_total']}**"
            )
            st.session_state.last_transcript = ""
            st.rerun()

# Preview of last parse (if text present and not yet added)
if transcript.strip() and not parse_clicked:
    preview = parse_transcript(transcript)
    with st.expander("Live parse preview (tap Parse & Add to save)", expanded=True):
        c1, c2, c3 = st.columns(3)
        c1.metric("Name", preview["name"] or "—")
        c2.metric("Roll", preview["roll"] or "—")
        c3.metric("Grand Total", preview["grand_total"])
        # Show found marks
        found = []
        for q in range(1, 11):
            parts = []
            for p in "abcdefghij":
                v = preview.get(f"q{q}{p}")
                if v is not None:
                    parts.append(f"{p}={v}")
            if parts:
                found.append(f"Q{q}: {', '.join(parts)}")
            elif preview.get(f"q{q}") is not None:
                found.append(f"Q{q}: {preview[f'q{q}']}")
        if found:
            st.write(" · ".join(found))
        else:
            st.caption("No question marks detected yet — try “question 1 A 5 marks”")

# ------------------------------------------------------------------
# Students list
# ------------------------------------------------------------------
st.markdown("---")
st.subheader("2. Students added")

if st.session_state.students:
    rows = []
    for s in st.session_state.students:
        row = {
            "Name": s.get("name", ""),
            "Roll No": s.get("roll", ""),
            "Grand Total": s.get("grand_total", 0),
        }
        for q in range(1, 11):
            # show compact marks
            parts = []
            for p in "abcdefghij":
                v = s.get(f"q{q}{p}")
                if v is not None:
                    parts.append(f"{p}:{v}")
            if parts:
                row[f"Q{q}"] = " ".join(parts)
            elif s.get(f"q{q}") is not None:
                row[f"Q{q}"] = str(s[f"q{q}"])
            else:
                row[f"Q{q}"] = ""
        rows.append(row)

    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

    m1, m2, m3 = st.columns(3)
    totals = [s["grand_total"] for s in st.session_state.students]
    m1.metric("Students", len(totals))
    m2.metric("Highest", max(totals) if totals else 0)
    m3.metric("Average", round(sum(totals) / len(totals), 1) if totals else 0)

    excel_bytes = create_excel(st.session_state.students)
    st.download_button(
        label="📥 Download Excel (Summary + Flat marks with formulas)",
        data=excel_bytes,
        file_name=f"Voice_Marks_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    if st.button("🔄 Clear all students"):
        st.session_state.students = []
        st.rerun()
else:
    st.info("No students yet. Dictate marks for the first student and tap **Parse & Add Student**.")

st.markdown("---")
st.caption(
    "Android tip: use Gboard voice typing into the text box while checking. "
    "Say “Name … Roll number … Question 1 A 5 marks …” clearly. "
    "Excel totals are formulas — open in Excel/Sheets to recalculate if needed."
)
