import streamlit as st
from openai import OpenAI
from docx import Document
import re
import pandas as pd

# =========================
# OPENAI CLIENT
# =========================
client = OpenAI(api_key="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")

# =========================
# MARKING PROMPT
# =========================
MARKING_PROMPT = """

TASK NAME:
Catching Teller Crow Analytical Essay

TEACHER CONTEXT:
This is a mixed-ability Year 10 cohort. Grading reflects cohort standard rather than strict rubric interpretation.

You must:
- First determine true rubric level
- Then adjust to realistic classroom level
- Prioritise understanding, evidence, and structure
- Do not average criteria

-------------------------

GRADE CONSTRAINTS:

- Responses that describe or summarise without analysis must be capped at C+/B-
- No B+ or above without sustained analytical reasoning
- Quotes do not increase marks without explanation
- Weak or repetitive explanations remain in C range
- Weak language technique analysis limits access to the highest A range, but does NOT automatically prevent A if understanding and analysis are strong

INCOMPLETE RESPONSE RULE (CRITICAL):

- FIRST: Determine if the response is complete BEFORE grading quality

A response is INCOMPLETE if:
- It lacks multiple fully developed paragraphs
- It only answers part of the task
- It stops abruptly or appears cut off
- It presents only fragmentary or partial analysis

If a response includes only one developed paragraph → treat as INCOMPLETE

IMPORTANT:
A shorter response can still be COMPLETE if it fully addresses the task with clear and developed analysis.

IF INCOMPLETE:
- MUST NOT exceed D+ under any circumstances
- Quality of individual paragraphs DOES NOT matter
- Presence of quotes DOES NOT matter

CRITICAL:
Incomplete responses must be graded as D or below, even if parts are strong

MANDATORY DECISION STEP:

Before assigning any grade, the model MUST explicitly decide:

Is this response COMPLETE or INCOMPLETE?

- If INCOMPLETE → apply cap immediately (D+ or below)
- If COMPLETE → proceed to grading

The model must NOT skip this step

The Completeness Check must directly control the grade:
- If marked INCOMPLETE → grade must follow the D+ cap rule with no exceptions


-------------------------

GRADE BAND INTERPRETATION (AUSTRALIAN CONTEXT):

A (85–100):
Clearly above expected standard. Consistent analysis, relevant evidence, logical structure. Not perfection.

B (70–84):
At expected standard. Clear understanding, some analysis, may lack depth.

C (60–69):
Basic understanding. Limited, repetitive, or inconsistent analysis.

D (50–59):
Below standard. Minimal explanation, weak analysis.

E/F (<50):
Very limited or incomplete.

-------------------------

GRADE PRECISION RULE (+ / -):

- Use + and - to reflect position within a grade band

Guidelines:

+ (e.g. B+, C+):
- Upper end of the band
- Close to the next grade above
- Shows several strengths but lacks consistency for the next band

Mid (e.g. B, C):
- Securely within the band
- Meets expectations without strong indicators of moving up or down

- (e.g. B-, C-):
- Lower end of the band
- Just meets the criteria for the band
- Shows noticeable weaknesses or inconsistency

CRITICAL:
- Do NOT default to plain B or C
- The model must choose +, mid, or - based on quality within the band


-------------------------

REAL CLASSROOM ANCHORS:

A:
- Sustained analysis across the response
- Explains HOW and WHY consistently
- Clearly above standard

B:
- Clear understanding, some analysis
- Inconsistent or surface-level at times

C:
- Descriptive, retelling, basic explanation
- Quotes present but weak analysis

D:
- Minimal explanation, unclear ideas

CRITICAL RULE:
If a response could reasonably be written by a student who understands the text but does not analyse deeply → it must remain in C range.

-------------------------

CRITICAL DISTINCTIONS:

- Description + quotes ≠ analysis
- If HOW/WHY is mostly missing → C range
If techniques are not analysed in any meaningful way → max B
- If techniques are only mentioned → max C+

-------------------------

A-RANGE CLARIFICATION (STRICT):

- A responses are COMMON in a mixed-ability cohort and must be awarded when deserved
- A does NOT require perfection, sophistication, or advanced literary insight
- If a response shows consistent explanation of HOW and WHY ideas are constructed → it MUST be considered for A

CRITICAL:
- Do NOT default strong responses to B due to minor weaknesses
- If the response is clearly above standard overall → it must be graded in the A range

REAL CLASSROOM ADJUSTMENT:

- In a mixed-ability cohort, responses that show clear understanding, relevant evidence, and mostly consistent explanation may still reach A- even if analysis is not fully sustained in every paragraph
- Occasional lapses in depth should NOT automatically prevent A-range if overall quality is clearly above standard

A- responses may show strong understanding and mostly consistent analysis, even if some paragraphs are less developed or technique analysis is not deeply sophisticated.

CRITICAL:
A responses should be awarded when the response is clearly above cohort standard, even if minor weaknesses are present.

-------------------------

OVERALL GRADE RULE:

- Reflect dominant performance across criteria
- 2+ C criteria → max B-
- Multiple D criteria → max C

-------------------------

ANTI-MIDDLE-BIAS RULE:

- Do NOT default to B when uncertain
- If analysis is limited → must be C range
- If analysis is strong and consistent → must be A range
- B should ONLY be used when the response clearly sits between these levels

CRITICAL:
- Overuse of B is incorrect grading
- The model must actively choose between A, B, and C based on evidence — not safety

If unsure between B and C → default to C unless clear analytical evidence justifies B

-------------------------

GRADE ANCHOR GUIDANCE:

Typical C response:
- Identifies theme
- Uses quotes
- Explains in a basic, repetitive, or descriptive way
- Limited depth of analysis

Typical B response:
- Clear understanding of the text
- Uses relevant evidence
- Some analysis is present, but may not be sustained or deep

Typical A response:
- Clear and consistent analysis throughout
- Strong understanding of ideas and themes
- Evidence is explained and linked to meaning
- Does NOT require perfect or highly sophisticated analysis, but must go beyond basic explanation

-------------------------

ANONYMISATION RULE:
Ignore student names.

-------------------------

RUBRIC:

Identify theme
A: analyse + explain
B: explain themes
C: explain a theme
D: limited explanation
E: mention only

Interpretation
A: strong justification
B: justified
C: some elaboration
D: limited
E: statement only

Language techniques
A: multiple techniques
B: some techniques
C: one technique
D: attempt
E: none

Evidence
A: fluent range
B: embedded
C: present
D: attempted
E: none

Style
A: strong formal
B: structured formal
C: some formal
D: limited
E: very limited

-------------------------

STUDENT RESPONSE:
[PASTE STUDENT RESPONSE HERE — DO NOT EDIT]

-------------------------

OUTPUT FORMAT:

Completeness Check:
[Complete / Incomplete]

Overall Indicative Level:
[Band]

Overall Justification:
1–2 sentences

---

Criterion-by-criterion:

Criterion name:
True Rubric Level:
Adjusted Level:
Justification:

---

Strengths:
- point
- point
- point

Areas for Improvement:
- point
- point
- point
"""

# =========================
# CLEAN TEXT (ANONYMISATION)
# =========================
def clean_text(text):
    text = re.sub(r'(?i)name\s*:\s*.*', '', text)

    lines = text.split("\n")
    cleaned = []

    for i, line in enumerate(lines):
        if i < 3 and len(line.split()) <= 4:
            continue
        cleaned.append(line)

    return "\n".join(cleaned)

# =========================
# READ DOCX
# =========================
def read_docx(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs])

# =========================
# CALL MODEL
# =========================
def mark_response(text):
    prompt = MARKING_PROMPT.replace(
        "[PASTE STUDENT RESPONSE HERE — DO NOT EDIT]", text
    )

    response = client.responses.create(
        model="gpt-4o-mini",
        input=prompt
    )

    return response.output_text

# =========================
# EXTRACT GRADE (ROBUST)
# =========================
def extract_grade(text):
    import re

    # Try strict format first
    match = re.search(
        r"Overall Indicative Level:\s*\n?\s*([A-D][\+\-]?)",
        text
    )
    if match:
        return match.group(1)

    # Fallback: grab first valid grade anywhere
    match_alt = re.search(
        r"\b(A\+|A-|A|B\+|B-|B|C\+|C-|C|D\+|D-|D)\b",
        text
    )
    if match_alt:
        return match_alt.group(1)

    return "N/A"

# =========================
# UI
# =========================
st.title("FirstPass Marking Tool")

st.info(
    "Note: For best results, use first names or depersonalised files. "
    "The system removes identifying info for marking but keeps file mapping in exports."
)

uploaded_files = st.file_uploader(
    "Upload student .docx files",
    type=["docx"],
    accept_multiple_files=True
)

# =========================
# RUN MARKING
# =========================
if st.button("Run Marking"):

    results = []

    for i, file in enumerate(uploaded_files):
        with st.spinner(f"Processing Student {i+1}..."):
            text = clean_text(read_docx(file))
            result = mark_response(text)
            grade = extract_grade(result)

            results.append({
                "student": f"Student {i+1}",
                "original_file": file.name,
                "grade": grade,
                "result": result
            })

    st.success("Done!")

    # =========================
    # SUMMARY
    # =========================
    st.write("### Summary")
    for r in results:
        st.write(f"{r['student']} — {r['grade']}")

    # =========================
    # CSV EXPORT (WITH MAPPING)
    # =========================
    df = pd.DataFrame(results)[["student", "original_file", "grade"]]

    st.download_button(
        label="Download Grades CSV",
        data=df.to_csv(index=False),
        file_name="grades_with_mapping.csv",
        mime="text/csv"
    )

    # =========================
    # FULL RESULTS
    # =========================
    for r in results:
        st.subheader(r["student"])
        st.caption(f"Original file: {r['original_file']}")
        st.write(r["result"])