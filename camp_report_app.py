#!/usr/bin/env python3
"""
Camp Data Analysis Tool  —  Gemma 4 Edition
=====================================================================
Features:
  • Gemma 4 E4B Model Integration (Fast, Lightweight, Smart)
  • Restored robust AI prompts (Eliminates Parse-Errors)
  • Restored original Quant Charts (Visual Summary)
  • Restored premium Excel formatting (Column widths, text wrapping)
  • Dynamic Excel Dashboards with Data Validation Dropdowns
  • Mixed Methods Triangulation (Quant Shifts × Qual Tags)
  • Automated Visual Qual Dashboards (All 8 Questions Charted)
  • NaN / INF Error Handling Patched
"""

import hashlib, json, os, re, traceback
from datetime import datetime

import numpy  as np
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from xlsxwriter.utility import xl_col_to_name


# ═══════════════════════════════════════════════════════════════════
#  UTILITY
# ═══════════════════════════════════════════════════════════════════

def find_col(df, keywords):
    """First column whose name contains every keyword (case-insensitive)."""
    for col in df.columns:
        if all(k.lower() in col.lower() for k in keywords):
            return col
    return None

def student_id(email):
    """Consistent 10-char anonymised ID derived from email."""
    return hashlib.sha256(email.strip().lower().encode()).hexdigest()[:10].upper()

def safe_json(text):
    """Robustly extract a JSON array from model output."""
    for src in [text, re.sub(r"```(?:json)?", "", text)]:
        src = src.strip()
        try:
            return json.loads(src)
        except Exception:
            pass
        m = re.search(r'\[.*\]', src, re.DOTALL)
        if m:
            try:
                return json.loads(m.group())
            except Exception:
                pass
    return None


# ═══════════════════════════════════════════════════════════════════
#  SIMPLIFIED QUALITATIVE CODEBOOK (RESTORED FULL DESCRIPTIONS)
# ═══════════════════════════════════════════════════════════════════

QUAL_QUESTIONS = [

    {   # ─── Q1 ─────────────────────────────────────────────────
        "num": 1, "label": "Peer Assistance",
        "keywords": ["helped a classmate"],
        "fields": {
            "Q1_Direction": ["Gave-Help", "Received-Help", "Mutual", "Not-Mentioned"],
            "Q1_Type": ["Equipment/Gear", "Campcraft", "Emotional-Support", "Physical-Help", "Taught-a-Skill", "Unclear"],
            "Q1_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": """
Q1_Direction:
  "Gave-Help"     → the student helped someone else
  "Received-Help" → someone helped the student
  "Mutual"        → both helped each other
  "Not-Mentioned" → no clear example

Q1_Type:
  "Equipment/Gear"     → Trangia, bike, navigation
  "Campcraft"          → tent setup, packing, campsite
  "Emotional-Support"  → encouragement, motivation
  "Physical-Help"      → carrying heavy gear
  "Taught-a-Skill"     → explaining how to do something
  "Unclear"            → type of help is not clear
"""
    },

    {   # ─── Q2 ─────────────────────────────────────────────────
        "num": 2, "label": "Challenge & Resilience",
        "keywords": ["biggest", "challenge"],
        "fields": {
            "Q2_Challenge": ["Physical-Effort", "Fear-or-Anxiety", "Weather-or-Environment", "Social-Difficulty", "Equipment-or-Skills", "Not-Specified"],
            "Q2_Response": ["Pushed-Through-Alone", "Friends-or-Staff-Helped", "Adapted-Approach", "Just-Tried-It", "Not-Explained"],
            "Q2_Growth": ["Yes", "Partial", "No"],
            "Q2_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": """
Q2_Challenge: Physical-Effort, Fear-or-Anxiety, Weather-or-Environment, Social-Difficulty, Equipment-or-Skills, Not-Specified.
Q2_Response: Pushed-Through-Alone, Friends-or-Staff-Helped, Adapted-Approach, Just-Tried-It, Not-Explained.
Q2_Growth: "Yes" (explicit growth), "Partial" (implied), "No".
"""
    },

    {   # ─── Q3 ─────────────────────────────────────────────────
        "num": 3, "label": "Surprising Skill",
        "keywords": ["surprised yourself"],
        "fields": {
            "Q3_Activity": ["Mountain-Biking", "Sea-Kayaking", "Snorkelling", "Coasteering", "Hiking", "Camp-Cooking", "Swimming", "Other"],
            "Q3_Growth_Type": ["Beat-Own-Expectations", "Overcame-a-Fear", "Got-Noticeably-Better", "Discovered-Enjoyment"],
            "Q3_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": """
Q3_Growth_Type:
  "Beat-Own-Expectations" → expected to dislike/fail, but didn't
  "Overcame-a-Fear"       → pushed through fear
  "Got-Noticeably-Better" → visible skill improvement
  "Discovered-Enjoyment"  → found unexpected fun
"""
    },

    {   # ─── Q4 ─────────────────────────────────────────────────
        "num": 4, "label": "First Nations Culture",
        "keywords": ["first nations", "stood out"],
        "fields": {
            "Q4_Topic": ["Plants-Animals-Country", "Stories-and-History", "Practices-and-Tools", "Bush-Food-and-Knowledge", "Movement-and-Country", "Vague-or-Unclear"],
            "Q4_Depth": ["Recalled-a-Fact", "New-Perspective", "Personal-Connection", "Surface-Level"],
            "Q4_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": """
Q4_Depth:
  "Recalled-a-Fact"     → remembered info without deeper thought
  "New-Perspective"     → describes a shift in understanding
  "Personal-Connection" → emotional resonance or empathy
  "Surface-Level"       → minimal engagement
"""
    },

    {   # ─── Q5 ─────────────────────────────────────────────────
        "num": 5, "label": "Favourite Part",
        "keywords": ["absolute favourite"],
        "fields": {
            "Q5_What": ["A-Specific-Activity", "Social-Time", "Scenery-or-Nature", "Personal-Achievement", "Food-or-Cooking", "Free-Time", "Everything"],
            "Q5_Why": ["Thrill-and-Fun", "Friendship", "Achievement", "Beauty-or-Peace", "Freedom-or-Choice", "Not-Explained"],
            "Q5_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": "Code WHAT their favourite part was and WHY it was their favourite."
    },

    {   # ─── Q6 ─────────────────────────────────────────────────
        "num": 6, "label": "Improvement Suggestions",
        "keywords": ["camp director"],
        "fields": {
            "Q6_Suggestion": ["More-Time-or-Activities", "Better-Food", "Better-Gear-or-Comfort", "Schedule-Change", "Happy-With-Everything", "Warmer-Weather-or-Season", "Other"],
            "Q6_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": "Code their primary suggestion for the camp director."
    },

    {   # ─── Q7 ─────────────────────────────────────────────────
        "num": 7, "label": "Advice for Year 7s",
        "keywords": ["year 7 student"],
        "fields": {
            "Q7_Advice": ["Mindset-and-Attitude", "Pack-and-Gear", "Food-and-Snacks", "Be-Social", "Comfort-and-Sleep", "Safety-and-Health", "Just-Try-Everything", "Not-Specified"],
            "Q7_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": "Code the main piece of advice given to younger students."
    },

    {   # ─── Q8 ─────────────────────────────────────────────────
        "num": 8, "label": "7-Day Camp Preparation",
        "keywords": ["better prepare"],
        "fields": {
            "Q8_Prep_Focus": ["Get-Fitter", "Better-Packing", "Practice-Camp-Skills", "Mental-Preparation", "More-School-Support", "Build-On-This-Camp", "Better-Food-Planning", "Not-Specified"],
            "Q8_Growth": ["Yes", "Partial", "No"],
            "Q8_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": "Code what they think is most important to prepare, and if there is evidence of forward-thinking growth."
    },
]

ALL_QUAL_FIELDS = [f for qc in QUAL_QUESTIONS for f in qc["fields"]]
SENTIMENT_FIELDS = [f for f in ALL_QUAL_FIELDS if f.endswith("_Sentiment")]
GROWTH_FIELDS    = ["Q2_Growth", "Q8_Growth"]
SENTIMENT_SCORE = {"Positive": 1.0, "Mixed": 0.25, "Neutral": 0.0, "Negative": -1.0}


# ═══════════════════════════════════════════════════════════════════
#  MLX INFERENCE & ROBUST AI PROMPTS
# ═══════════════════════════════════════════════════════════════════

def call_mlx(model, tokenizer, system_msg, user_msg, max_tokens=2500):
    from mlx_lm import generate
    messages = [{"role": "system", "content": system_msg},
                {"role": "user",   "content": user_msg}]
    if hasattr(tokenizer, "apply_chat_template"):
        prompt = tokenizer.apply_chat_template(
            messages, tokenize=False, add_generation_prompt=True)
    else:
        prompt = f"### System:\n{system_msg}\n\n### User:\n{user_msg}\n\n### Assistant:\n"
    return generate(model, tokenizer, prompt=prompt, max_tokens=max_tokens, verbose=False)


SYSTEM_CODER = (
    "You are coding student survey responses from a school camp. "
    "Return ONLY a valid JSON array — no explanation, no markdown, no preamble. "
    "Use ONLY the listed values for each field. Never invent new values. "
    "If a response is completely nonsensical or off-topic, use the 'Not-Specified', "
    "'Unclear', or 'No' tags provided in the guide."
)

def build_coding_prompt(qc, batch):
    """Restored full verbose prompt to prevent Parse-Errors."""
    fields_spec = ""
    for field, values in qc["fields"].items():
        fields_spec += f'\n  "{field}": one of {json.dumps(values)}'

    items = "\n".join(
        f'{{"id":{item["id"]},"text":{json.dumps(item["text"])}}}'
        for item in batch
    )

    example_fields = ", ".join(f'"{f}": "..."' for f in qc["fields"])

    return f"""Code each response to this camp survey question:
"{qc['label']}"

FIELDS TO CODE FOR EACH RESPONSE:{fields_spec}

CODING GUIDE:
{qc['guide']}

STUDENT RESPONSES:
{items}

Return ONLY a JSON array. Each object must have "id" plus the coded fields.
Example format: [{{"id": 0, {example_fields}}}]"""


def code_question(model, tokenizer, qc, post_df, email_list, status_label, root, BATCH=30):
    col = find_col(post_df, qc["keywords"])
    fields = list(qc["fields"].keys())
    
    email_resp = dict(zip(post_df["Email address"], post_df.get(col, pd.Series(dtype=str))))

    items = []
    for idx, email in enumerate(email_list):
        resp = str(email_resp.get(email, "")).strip()
        items.append({"id": idx, "text": resp, "empty": len(resp) <= 3})

    coded = {}
    to_code = [it for it in items if not it["empty"]]

    total_batches = max(1, (len(to_code) + BATCH - 1) // BATCH)
    for bn, start in enumerate(range(0, len(to_code), BATCH), 1):
        batch = to_code[start : start + BATCH]
        status_label.config(
            text=f"Status: Coding Q{qc['num']} — batch {bn}/{total_batches}...",
            fg="#D97706"
        )
        root.update()

        prompt = build_coding_prompt(qc, batch)
        raw    = call_mlx(model, tokenizer, SYSTEM_CODER, prompt, max_tokens=2500)
        parsed = safe_json(raw)

        if isinstance(parsed, list):
            for row in parsed:
                if isinstance(row, dict) and "id" in row:
                    idx = row["id"]
                    coded[idx] = {f: str(row.get(f, "Parse-Error")).strip() for f in fields}

        for it in batch:
            if it["id"] not in coded:
                coded[it["id"]] = {f: "Parse-Error" for f in fields}

    for it in items:
        if it["empty"]:
            coded[it["id"]] = {f: "No-Response" for f in fields}

    return coded


SYSTEM_NARRATOR = (
    "You are an educational analyst writing a concise summary for a school camp report. "
    "Write in clear paragraphs. Be specific and evidence-based."
)

def generate_narrative(model, tokenizer, qc, responses, status_label, root):
    status_label.config(text=f"Status: Writing Q{qc['num']} narrative & extracting quotes...", fg="#D97706")
    root.update()

    sample = [r for r in responses if len(r.strip()) > 3][:50]
    if not sample:
        return "Insufficient responses to generate a summary."

    resp_text = "\n".join(f"- {r}" for r in sample)
    prompt = f"""Summarise these student responses about: "{qc['label']}"

Write exactly 2 short paragraphs:
  1. The most common themes and overall student sentiment.
  2. What these responses reveal about student development or growth.

Then, below your summary, include exactly 3 to 5 highly valuable, reflective, verbatim quotes from the students that capture the essence of their experience. Format the quotes cleanly with bullet points and quotation marks.

Student responses (n={len(sample)}):
{resp_text}"""

    return call_mlx(model, tokenizer, SYSTEM_NARRATOR, prompt, max_tokens=600).strip()


# ═══════════════════════════════════════════════════════════════════
#  DATASET BUILDERS
# ═══════════════════════════════════════════════════════════════════

def build_coded_df(post_df, email_list, names, surnames, classes, locations,
                   model, tokenizer, status_label, root, year):
    coded_df = pd.DataFrame({
        "Year":       year,
        "Student_ID": [student_id(e) for e in email_list],
        "First_Name": names,
        "Surname":    surnames,
        "Class":      classes,
        "Location":   locations,
    })

    narratives = []

    for qc in QUAL_QUESTIONS:
        qnum = qc["num"]
        col  = find_col(post_df, qc["keywords"])

        email_resp = {}
        if col:
            email_resp = dict(zip(post_df["Email address"], post_df[col]))
        raw_vals = [str(email_resp.get(e, "")).strip() for e in email_list]
        coded_df[f"Q{qnum}_Response"] = raw_vals

        if not col:
            for f in qc["fields"]:
                coded_df[f] = "Column-Not-Found"
            narratives.append({
                "Question": f"Q{qnum}: {qc['label']}",
                "n": 0,
                "Summary": "Column not found in post-camp CSV.",
            })
            continue

        coded = code_question(model, tokenizer, qc, post_df, email_list, status_label, root)
        for field in qc["fields"]:
            coded_df[field] = [coded.get(i, {}).get(field, "No-Response") for i in range(len(email_list))]

        narrative_text = generate_narrative(model, tokenizer, qc, raw_vals, status_label, root)
        valid_n = sum(1 for r in raw_vals if len(r) > 3)
        narratives.append({
            "Question": f"Q{qnum}: {qc['label']}",
            "n": valid_n,
            "Summary": narrative_text,
        })

    return coded_df, narratives


def build_summary_metadata():
    rows = []
    for qc in QUAL_QUESTIONS:
        qnum = qc["num"]
        rows.append(("SECTION", f"Q{qnum} — {qc['label']}", "", ""))

        for field, valid_vals in qc["fields"].items():
            rows.append(("FIELD", field, "", ""))
            for val in valid_vals:
                rows.append(("DATA", field, val, ""))
            rows.append(("BLANK", "", "", ""))
        rows.append(("BLANK", "", "", ""))
    return rows


def build_longitudinal_df(coded_df, merged_df, metrics_processed, year):
    long_df = coded_df[["Year","Student_ID","First_Name","Surname","Class","Location"]].copy()

    for metric in metrics_processed:
        col_shift = metric
        if col_shift in merged_df.columns:
            long_df[f"{metric}_Shift"] = merged_df[col_shift].values
        for suffix in ("_pre", "_post"):
            candidates = [c for c in merged_df.columns if metric.lower().replace(" ","_") in c.lower() and c.endswith(suffix)]
            if candidates:
                long_df[f"{metric}{suffix.capitalize()}"] = merged_df[candidates[0]].values

    for qc in QUAL_QUESTIONS:
        for field in qc["fields"]:
            if field in coded_df.columns:
                long_df[field] = coded_df[field].values

    def sent_num(val): return SENTIMENT_SCORE.get(str(val), np.nan)
    sent_cols = [f for f in SENTIMENT_FIELDS if f in long_df.columns]
    if sent_cols:
        long_df["Sentiment_Score"] = long_df[sent_cols].map(sent_num).mean(axis=1).round(2)
    else:
        long_df["Sentiment_Score"] = np.nan

    def growth_num(val): return 1 if val == "Yes" else (0.5 if val == "Partial" else 0)
    g_cols = [f for f in GROWTH_FIELDS if f in long_df.columns]
    if g_cols:
        long_df["Growth_Score"] = long_df[g_cols].map(growth_num).sum(axis=1).round(1)
    else:
        long_df["Growth_Score"] = np.nan

    return long_df


def build_mixed_methods(long_df):
    """Triangulates Qual Tags with Quant Shifts"""
    tables = []
    pairs = [
        ("Teamwork Shift by Peer Help Type", "Q1_Type", "Teamwork"),
        ("Teamwork Shift by Help Direction", "Q1_Direction", "Teamwork"),
        ("Drive Shift by Resilience Strategy", "Q2_Response", "Drive"),
        ("Drive Shift by Growth Mindset", "Q2_Growth", "Drive"),
        ("Cultural Shift by Engagement Depth", "Q4_Depth", "Aboriginal Culture"),
        ("Future Readiness (Hike) by Prep Growth", "Q8_Growth", "Overnight Hike")
    ]
    for title, qual, quant in pairs:
        q_col = f"{quant}_Shift"
        if qual in long_df.columns and q_col in long_df.columns:
            grp = long_df.groupby(qual)[q_col].agg(Average_Shift='mean', Student_Count='count').reset_index().round(2)
            grp.insert(0, "Analysis", title)
            grp.rename(columns={qual: "Qualitative Tag", "Average_Shift": "Avg Shift", "Student_Count": "n"}, inplace=True)
            tables.append(grp)
    return tables


def build_qual_chart_data(long_df):
    charts = {}
    if "Q1_Type" in long_df.columns: charts["Q1"] = long_df["Q1_Type"].value_counts().reset_index()
    if "Q2_Challenge" in long_df.columns and "Q2_Growth" in long_df.columns:
        charts["Q2"] = pd.crosstab(long_df["Q2_Challenge"], long_df["Q2_Growth"]).reindex(columns=["Yes", "Partial", "No"], fill_value=0).reset_index()
    if "Q3_Activity" in long_df.columns and "Q3_Growth_Type" in long_df.columns:
        charts["Q3"] = pd.crosstab(long_df["Q3_Activity"], long_df["Q3_Growth_Type"]).reset_index()
    if "Q4_Topic" in long_df.columns and "Q4_Depth" in long_df.columns:
        charts["Q4"] = pd.crosstab(long_df["Q4_Topic"], long_df["Q4_Depth"]).reindex(columns=["New-Perspective", "Personal-Connection", "Recalled-a-Fact", "Surface-Level"], fill_value=0).reset_index()
    if "Q5_What" in long_df.columns: charts["Q5"] = long_df["Q5_What"].value_counts().reset_index()
    if "Q6_Suggestion" in long_df.columns: charts["Q6"] = long_df["Q6_Suggestion"].value_counts().reset_index()
    if "Q7_Advice" in long_df.columns: charts["Q7"] = long_df["Q7_Advice"].value_counts().reset_index()
    if "Q8_Prep_Focus" in long_df.columns and "Q8_Growth" in long_df.columns:
        charts["Q8"] = pd.crosstab(long_df["Q8_Prep_Focus"], long_df["Q8_Growth"]).reindex(columns=["Yes", "Partial", "No"], fill_value=0).reset_index()
    return charts


# ═══════════════════════════════════════════════════════════════════
#  EXCEL WRITER (RESTORED FORMATTING)
# ═══════════════════════════════════════════════════════════════════

def write_excel(output_path, tab1_df, avg_df, dist_df, breakdown_df,
                coded_df, summary_metadata, long_df, narratives, metrics_processed, mm_tables, qual_charts):

    writer = pd.ExcelWriter(output_path, engine="xlsxwriter")
    workbook = writer.book
    workbook.nan_inf_to_errors = True  # FIX: Prevent crash on Infinity or NaN

    F = {
        "head":      workbook.add_format({"bold":True,"bg_color":"#D9D9D9","border":1}),
        "head_grn":  workbook.add_format({"bold":True,"bg_color":"#E2EFDA","border":1,"text_wrap":True}),
        "head_blu":  workbook.add_format({"bold":True,"bg_color":"#DDEEFF","border":1}),
        "skip":      workbook.add_format({"font_color":"#AAAAAA","italic":True}),
        "err":       workbook.add_format({"font_color":"#FF6600","italic":True}),
        "g_shift":   workbook.add_format({"bg_color":"#C6EFCE","font_color":"#006100","num_format":"+0;-0;0"}),
        "r_shift":   workbook.add_format({"bg_color":"#FFC7CE","font_color":"#9C0006","num_format":"+0;-0;0"}),
        "z_shift":   workbook.add_format({"num_format":"0"}),
        "num":       workbook.add_format({"num_format":"0.00"}),
        "g_dec":     workbook.add_format({"bg_color":"#C6EFCE","font_color":"#006100","num_format":"+0.00;-0.00;0.00"}),
        "r_dec":     workbook.add_format({"bg_color":"#FFC7CE","font_color":"#9C0006","num_format":"+0.00;-0.00;0.00"}),
        "z_dec":     workbook.add_format({"num_format":"0.00"}),
        "pct":       workbook.add_format({"num_format":"0%"}),
        "pct1":      workbook.add_format({"num_format":"0.0%"}),
        "wrap":      workbook.add_format({"text_wrap":True,"valign":"top"}),
        "section":   workbook.add_format({"bold":True,"bg_color":"#375623","font_color":"#FFFFFF","font_size":11}),
        "field_hdr": workbook.add_format({"bold":True,"bg_color":"#E2EFDA","font_color":"#375623","italic":True}),
        "field_hdr_n": workbook.add_format({"bold":True,"bg_color":"#E2EFDA","font_color":"#375623","italic":True,"num_format":'"n = "0'}),
        "sent_pos":  workbook.add_format({"bg_color":"#C6EFCE","font_color":"#006100"}),
        "sent_neg":  workbook.add_format({"bg_color":"#FFC7CE","font_color":"#9C0006"}),
        "sent_mix":  workbook.add_format({"bg_color":"#FFEB9C","font_color":"#7D4E00"}),
        "sent_neu":  workbook.add_format({"font_color":"#555555"}),
        "grow_yes":  workbook.add_format({"bg_color":"#C6EFCE","font_color":"#006100","bold":True}),
        "grow_par":  workbook.add_format({"bg_color":"#FFEB9C","font_color":"#7D4E00"}),
        "grow_no":   workbook.add_format({"bg_color":"#F2F2F2","font_color":"#AAAAAA"}),
        "long_head": workbook.add_format({"bold":True,"bg_color":"#1F3864","font_color":"#FFFFFF","border":1,"text_wrap":True}),
        "mm_head":   workbook.add_format({"bold":True,"bg_color":"#4472C4","font_color":"#FFFFFF","border":1}),
    }

    def decimal_grid(ws, df, start_col):
        for r in range(len(df)):
            for c in range(start_col, len(df.columns)):
                v = df.iloc[r, c]
                if pd.isna(v):   ws.write(r+1, c, "—", F["skip"])
                elif v >  0.001: ws.write_number(r+1, c, v, F["g_dec"])
                elif v < -0.001: ws.write_number(r+1, c, v, F["r_dec"])
                else:            ws.write_number(r+1, c, v, F["z_dec"])

    def sent_fmt(val):
        if val == "Positive": return F["sent_pos"]
        if val == "Negative": return F["sent_neg"]
        if val == "Mixed":    return F["sent_mix"]
        return F["sent_neu"]

    def grow_fmt(val):
        if val == "Yes":     return F["grow_yes"]
        if val == "Partial": return F["grow_par"]
        return F["grow_no"]

    # ── Write DataFrames to sheets (with na_rep fix) ──────────────
    tab1_df.to_excel(writer,      sheet_name="Individual Scores", index=False, na_rep="—")
    avg_df.to_excel(writer,       sheet_name="Group Averages",    index=False, na_rep="—")
    breakdown_df.to_excel(writer, sheet_name="Group Breakdown",   index=False, na_rep="—")
    coded_df.to_excel(writer,     sheet_name="Qual Responses",    index=False, na_rep="—")
    long_df.to_excel(writer,      sheet_name="Longitudinal Data", index=False, na_rep="—")
    dist_df.to_excel(writer,      sheet_name="_ChartData",        index=False, na_rep="0")

    # ──────────────────────────────────────────────────────────────
    #  Visual Summary (Restored Original Quant Charts)
    # ──────────────────────────────────────────────────────────────
    ws_charts = workbook.add_worksheet("Visual Summary")
    ws_charts.hide_gridlines(2)
    writer.sheets["Visual Summary"] = ws_charts
    ws_charts.set_first_sheet()
    ws_charts.activate()

    # Chart 1: Pre/Post Averages
    c1 = workbook.add_chart({"type": "bar"})
    for name, col_idx, colour in [("Pre-Camp Average", 1, "#8FAADC"), ("Post-Camp Average", 2, "#A9D18E")]:
        c1.add_series({
            "name": name,
            "categories": ["Group Averages", 1, 0, len(avg_df), 0],
            "values":     ["Group Averages", 1, col_idx, len(avg_df), col_idx],
            "fill": {"color": colour},
            "data_labels": {"value": True, "num_format": "0.0"},
        })
    c1.set_title({"name": "Pre-Camp vs Post-Camp Averages", "name_font": {"size": 14, "bold": True}})
    c1.set_x_axis({"name": "Average Score (out of 10)", "max": 10, "major_gridlines": {"visible": True}})
    c1.set_y_axis({"reverse": True})
    c1.set_size({"width": 800, "height": 450})
    ws_charts.insert_chart("B2", c1)

    # Chart 2: Impact Distribution
    c2 = workbook.add_chart({"type": "bar", "subtype": "percent_stacked"})
    for name, col_idx, colour in [("Declined", 3, "#FFC7CE"), ("Stayed the Same", 2, "#D9D9D9"), ("Improved", 1, "#C6EFCE")]:
        c2.add_series({
            "name": name,
            "categories": ["_ChartData", 1, 0, len(dist_df), 0],
            "values":     ["_ChartData", 1, col_idx, len(dist_df), col_idx],
            "fill": {"color": colour},
            "data_labels": {"value": True, "num_format": "0%"},
        })
    c2.set_title({"name": "Student Impact: Distribution of Change", "name_font": {"size": 14, "bold": True}})
    c2.set_x_axis({"name": "% of Students"})
    c2.set_y_axis({"reverse": True})
    c2.set_size({"width": 800, "height": 450})
    ws_charts.insert_chart("B26", c2)

    # ──────────────────────────────────────────────────────────────
    #  Format: Individual Scores
    # ──────────────────────────────────────────────────────────────
    ws = writer.sheets["Individual Scores"]
    for i, col in enumerate(tab1_df.columns): ws.write(0, i, col, F["head"])
    ws.set_column("A:F", 15); ws.set_column(6, len(tab1_df.columns)-1, 16)
    for r in range(len(tab1_df)):
        for c, m in enumerate(metrics_processed, start=6):
            v = tab1_df.iloc[r, c]
            if pd.isna(v):  ws.write(r+1, c, "—", F["skip"])
            elif v > 0:     ws.write_number(r+1, c, v, F["g_shift"])
            elif v < 0:     ws.write_number(r+1, c, v, F["r_shift"])
            else:           ws.write_number(r+1, c, v, F["z_shift"])

    # ──────────────────────────────────────────────────────────────
    #  Format: Group Averages & Breakdown
    # ──────────────────────────────────────────────────────────────
    ws = writer.sheets["Group Averages"]
    for i, col in enumerate(avg_df.columns): ws.write(0, i, col, F["head"])
    ws.set_column("A:A", 25); ws.set_column("B:D", 18)
    decimal_grid(ws, avg_df, 1)

    ws = writer.sheets["Group Breakdown"]
    for i, col in enumerate(breakdown_df.columns): ws.write(0, i, col, F["head"])
    ws.set_column("A:B", 22); ws.set_column(2, len(breakdown_df.columns)-1, 16)
    decimal_grid(ws, breakdown_df, 2)

    # ──────────────────────────────────────────────────────────────
    #  Qual Responses  (Restored Formatting + Dropdowns)
    # ──────────────────────────────────────────────────────────────
    ws = writer.sheets["Qual Responses"]
    for i, col in enumerate(coded_df.columns):
        ws.write(0, i, col, F["head_grn"])
    for i, col in enumerate(coded_df.columns):
        if "_Response" in col:
            ws.set_column(i, i, 50, F["wrap"])
        elif col in ("First_Name","Surname","Class","Location","Year","Student_ID"):
            ws.set_column(i, i, 16)
        else:
            ws.set_column(i, i, 22)
    ws.freeze_panes(1, 4)

    for r in range(len(coded_df)):
        for c, col in enumerate(coded_df.columns):
            val = coded_df.iloc[r, c]
            val_str = "" if pd.isna(val) else str(val)
            if val_str in ("No-Response", "Column-Not-Found", ""):
                ws.write(r+1, c, val_str, F["skip"])
            elif val_str == "Parse-Error":
                ws.write(r+1, c, val_str, F["err"])
            elif col.endswith("_Sentiment"):
                ws.write(r+1, c, val_str, sent_fmt(val_str))
            elif col in GROWTH_FIELDS:
                ws.write(r+1, c, val_str, grow_fmt(val_str))
            else:
                ws.write(r+1, c, val_str)

    for i, col_name in enumerate(coded_df.columns):
        options = []
        for qc in QUAL_QUESTIONS:
            if col_name in qc["fields"]:
                options = qc["fields"][col_name]
                break
        if options:
            c_let = xl_col_to_name(i)
            dropdown = options + ["No-Response", "Parse-Error"]
            ws.data_validation(f"{c_let}2:{c_let}5000", {'validate': 'list', 'source': dropdown})

    # ──────────────────────────────────────────────────────────────
    #  Qual Summary  (Dynamic Formulas)
    # ──────────────────────────────────────────────────────────────
    ws_sum = workbook.add_worksheet("Qual Summary")
    writer.sheets["Qual Summary"] = ws_sum
    ws_sum.set_column("A:A", 38)
    ws_sum.set_column("B:B", 12)
    ws_sum.set_column("C:C", 12)
    ws_sum.hide_gridlines(2)

    col_map = {name: xl_col_to_name(i) for i, name in enumerate(coded_df.columns)}

    r = 0
    last_field_row = 0
    for item in summary_metadata:
        row_type = item[0]
        if row_type == "SECTION":
            title = item[1]
            ws_sum.write(r, 0, title, F["section"])
            ws_sum.write(r, 1, "",    F["section"])
            ws_sum.write(r, 2, "",    F["section"])
        elif row_type == "FIELD":
            field_name = item[1]
            ws_sum.write(r, 0, field_name, F["field_hdr"])
            opt_count = sum(1 for x in summary_metadata if x[0] == "DATA" and x[1] == field_name)
            if opt_count > 0:
                ws_sum.write_formula(r, 1, f"=SUM(B{r+2}:B{r+1+opt_count})", F["field_hdr_n"])
            else:
                ws_sum.write(r, 1, 0, F["field_hdr_n"])
            ws_sum.write(r, 2, "", F["field_hdr"])
            last_field_row = r 
        elif row_type == "DATA":
            field_name = item[1]
            val = item[2]
            ws_sum.write(r, 0, val)
            if field_name in col_map:
                c_let = col_map[field_name]
                ws_sum.write_formula(r, 1, f'=COUNTIF(\'Qual Responses\'!${c_let}:${c_let}, "{val}")')
                ws_sum.write_formula(r, 2, f'=IF(B{last_field_row+1}>0, B{r+1}/B{last_field_row+1}, 0)', F["pct1"])
            else:
                ws_sum.write(r, 1, 0)
                ws_sum.write(r, 2, 0, F["pct1"])
        r += 1

    # ──────────────────────────────────────────────────────────────
    #  Qual Highlights (Narratives + Quotes)
    # ──────────────────────────────────────────────────────────────
    ws_high = workbook.add_worksheet("Qual Highlights")
    writer.sheets["Qual Highlights"] = ws_high
    ws_high.set_column("A:A", 25)
    ws_high.set_column("B:B", 110, F["wrap"])
    ws_high.hide_gridlines(2)

    row_idx = 1
    for nar in narratives:
        ws_high.write(row_idx, 0, nar["Question"], F["section"])
        ws_high.write(row_idx, 1, f"n = {nar['n']}", F["section"])
        row_idx += 1
        
        ws_high.write(row_idx, 0, "Summary & Quotes", F["field_hdr"])
        ws_high.write(row_idx, 1, nar["Summary"], F["wrap"])
        
        lines = len(nar["Summary"].split('\n')) + (len(nar["Summary"]) // 80)
        ws_high.set_row(row_idx, max(75, lines * 16))
        row_idx += 3

    # ──────────────────────────────────────────────────────────────
    #  Mixed Methods Insights
    # ──────────────────────────────────────────────────────────────
    ws_mm = workbook.add_worksheet("Mixed Methods Insights")
    writer.sheets["Mixed Methods Insights"] = ws_mm
    ws_mm.set_column("A:A", 35); ws_mm.set_column("B:D", 15)
    row_cursor = 1
    for df in mm_tables:
        for i, col in enumerate(df.columns): ws_mm.write(row_cursor, i, col, F["mm_head"])
        for r_idx in range(len(df)):
            for c_idx in range(len(df.columns)):
                val = df.iloc[r_idx, c_idx]
                if pd.isna(val):
                    ws_mm.write(row_cursor + 1 + r_idx, c_idx, "—", F["skip"])
                else:
                    fmt = F["g_shift"] if c_idx == 2 and isinstance(val, (int,float)) and val > 0 else F["r_shift"] if c_idx == 2 and isinstance(val, (int,float)) and val < 0 else None
                    ws_mm.write(row_cursor + 1 + r_idx, c_idx, val, fmt)
        row_cursor += len(df) + 3

    # ──────────────────────────────────────────────────────────────
    #  Longitudinal Data
    # ──────────────────────────────────────────────────────────────
    ws = writer.sheets["Longitudinal Data"]
    for i, col in enumerate(long_df.columns):
        ws.write(0, i, col, F["long_head"])
    ws.set_column("A:A", 8)   
    ws.set_column("B:B", 14)  
    ws.set_column("C:E", 16)  
    ws.set_column("F:F", 16)  
    ws.set_column(6, len(long_df.columns)-1, 20)
    ws.freeze_panes(1, 2)

    for r in range(len(long_df)):
        for c, col in enumerate(long_df.columns):
            v = long_df.iloc[r, c]
            if pd.isna(v):
                ws.write(r+1, c, "")
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.write_number(r+1, c, v, F["num"] if "." in str(v) else None)
            elif str(v) in ("No-Response","Column-Not-Found",""):
                ws.write(r+1, c, str(v), F["skip"])
            elif col.endswith("_Sentiment"):
                ws.write(r+1, c, str(v), sent_fmt(str(v)))
            elif col in GROWTH_FIELDS:
                ws.write(r+1, c, str(v), grow_fmt(str(v)))
            else:
                ws.write(r+1, c, str(v))

    # ──────────────────────────────────────────────────────────────
    #  Qual Dashboard (All 8 Charts)
    # ──────────────────────────────────────────────────────────────
    ws_qchart = workbook.add_worksheet("_QualChartData")
    ws_qchart.set_column("B:D", None, F["pct"])
    ws_qchart.hide()
    
    ws_dash = workbook.add_worksheet("Qual Dashboard")
    writer.sheets["Qual Dashboard"] = ws_dash
    ws_dash.hide_gridlines(2)

    if "Q1" in qual_charts:
        df1 = qual_charts["Q1"]
        df1.to_excel(writer, sheet_name="_QualChartData", startcol=0, index=False)
        c1 = workbook.add_chart({'type': 'doughnut'})
        c1.add_series({'categories': ['_QualChartData', 1, 0, len(df1), 0], 'values': ['_QualChartData', 1, 1, len(df1), 1], 'data_labels': {'percentage': True}})
        c1.set_title({'name': 'Q1: Types of Peer Assistance'})
        c1.set_size({'width': 500, 'height': 300})
        ws_dash.insert_chart('B2', c1)

    if "Q2" in qual_charts:
        df2 = qual_charts["Q2"]
        df2.to_excel(writer, sheet_name="_QualChartData", startcol=3, index=False)
        c2 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
        for i, name in enumerate(["Yes", "Partial", "No"]):
            c2.add_series({'name': name, 'categories': ['_QualChartData', 1, 3, len(df2), 3], 'values': ['_QualChartData', 1, 4+i, len(df2), 4+i], 'fill': {'color': ['#A9D18E', '#FFD966', '#D9D9D9'][i]}})
        c2.set_title({'name': 'Q2: Resilience Growth by Challenge Type'})
        c2.set_size({'width': 500, 'height': 300})
        ws_dash.insert_chart('J2', c2)

    if "Q3" in qual_charts:
        df3 = qual_charts["Q3"]
        df3.to_excel(writer, sheet_name="_QualChartData", startcol=8, index=False)
        c3 = workbook.add_chart({'type': 'column'})
        for i, name in enumerate(df3.columns[1:]): 
            c3.add_series({'name': name, 'categories': ['_QualChartData', 1, 8, len(df3), 8], 'values': ['_QualChartData', 1, 9+i, len(df3), 9+i]})
        c3.set_title({'name': 'Q3: Surprising Skills & Growth Types'})
        c3.set_size({'width': 500, 'height': 300})
        ws_dash.insert_chart('B18', c3)

    if "Q4" in qual_charts:
        df4 = qual_charts["Q4"]
        df4.to_excel(writer, sheet_name="_QualChartData", startcol=14, index=False)
        c4 = workbook.add_chart({'type': 'bar', 'subtype': 'stacked'})
        for i, name in enumerate(["New-Perspective", "Personal-Connection", "Recalled-a-Fact", "Surface-Level"]):
            c4.add_series({'name': name, 'categories': ['_QualChartData', 1, 14, len(df4), 14], 'values': ['_QualChartData', 1, 15+i, len(df4), 15+i]})
        c4.set_title({'name': 'Q4: First Nations - Depth of Internalization'})
        c4.set_size({'width': 500, 'height': 300})
        ws_dash.insert_chart('J18', c4)

    if "Q5" in qual_charts:
        df5 = qual_charts["Q5"]
        df5.to_excel(writer, sheet_name="_QualChartData", startcol=20, index=False)
        c5 = workbook.add_chart({'type': 'doughnut'})
        c5.add_series({'categories': ['_QualChartData', 1, 20, len(df5), 20], 'values': ['_QualChartData', 1, 21, len(df5), 21], 'data_labels': {'percentage': True}})
        c5.set_title({'name': 'Q5: Favourite Part of Camp'})
        c5.set_size({'width': 500, 'height': 300})
        ws_dash.insert_chart('B34', c5)

    if "Q6" in qual_charts:
        df6 = qual_charts["Q6"].sort_values(by=qual_charts["Q6"].columns[1], ascending=True) 
        df6.to_excel(writer, sheet_name="_QualChartData", startcol=23, index=False)
        c6 = workbook.add_chart({'type': 'bar'})
        c6.add_series({'categories': ['_QualChartData', 1, 23, len(df6), 23], 'values': ['_QualChartData', 1, 24, len(df6), 24], 'data_labels': {'value': True}, 'fill': {'color': '#5B9BD5'}})
        c6.set_title({'name': 'Q6: Top Improvement Suggestions'})
        c6.set_legend({'none': True})
        c6.set_size({'width': 500, 'height': 300})
        ws_dash.insert_chart('J34', c6)

    if "Q7" in qual_charts:
        df7 = qual_charts["Q7"]
        df7.to_excel(writer, sheet_name="_QualChartData", startcol=26, index=False)
        c7 = workbook.add_chart({'type': 'radar', 'subtype': 'filled'})
        c7.add_series({'categories': ['_QualChartData', 1, 26, len(df7), 26], 'values': ['_QualChartData', 1, 27, len(df7), 27], 'fill': {'color': '#ED7D31', 'transparency': 50}})
        c7.set_title({'name': 'Q7: Cohort Wisdom Profile (Advice for Y7s)'})
        c7.set_legend({'none': True})
        c7.set_size({'width': 500, 'height': 300})
        ws_dash.insert_chart('B50', c7)

    if "Q8" in qual_charts:
        df8 = qual_charts["Q8"]
        df8.to_excel(writer, sheet_name="_QualChartData", startcol=29, index=False)
        c8 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
        for i, name in enumerate(["Yes", "Partial", "No"]):
            c8.add_series({'name': name, 'categories': ['_QualChartData', 1, 29, len(df8), 29], 'values': ['_QualChartData', 1, 30+i, len(df8), 30+i], 'fill': {'color': ['#A9D18E', '#FFD966', '#D9D9D9'][i]}})
        c8.set_title({'name': 'Q8: 7-Day Prep Focus vs. Growth Mindset'})
        c8.set_size({'width': 500, 'height': 300})
        ws_dash.insert_chart('J50', c8)

    writer.close()


# ═══════════════════════════════════════════════════════════════════
#  QUANTITATIVE PROCESSING
# ═══════════════════════════════════════════════════════════════════

METRIC_MAP = {
    "Camping Skills": (["confident", "camping"], ["confident", "camping"]),
    "Sleeping Outdoors": (["sleeping", "outdoors", "bugs"], ["handle", "sleeping", "outdoors"]),
    "Swimming": (["confident", "swimming", "deep"], ["confident", "swimming", "deep"]),
    "Biking Comfort": (["comfortable", "riding", "bike"], ["comfortable", "bike", "now"]),
    "Overnight Hike": (["excited", "overnight", "hike"], ["excited", "overnight", "hike", "future"]),
    "Sea Kayaking": (["excited", "sea kayaking"], ["excited", "sea kayaking", "again"]),
    "Coasteering": (["excited", "coasteering"], ["excited", "coasteering", "again"]),
    "Mountain Biking": (["excited", "mountain bike", "flowy"], ["excited", "mountain biking", "flowy"]),
    "Snorkelling": (["excited", "snorkelling"], ["excited", "snorkelling", "again"]),
    "Teamwork": (["confident", "group", "work together"],["during camp","group","work together"]),
    "Autonomy": (["organising", "own gear"], ["during camp", "manage", "own gear"]),
    "Drive": (["physical activity", "tiring"], ["things got tiring", "focus", "energy"]),
    "Aboriginal Culture":(["Thinking back", "Aboriginal culture"],["Now that camp is finished","Aboriginal culture"]),
}

def process_quantitative(students_df, pre_df, post_df):
    merged = pd.merge(students_df[["First name", "Surname", "Email"]], pre_df, left_on="Email", right_on="Email address", how="left")
    merged = pd.merge(merged, post_df, left_on="Email", right_on="Email address", how="left", suffixes=("_pre", "_post"))
    
    loc_raw, class_raw = find_col(post_df, ["location"]), find_col(post_df, ["class"])
    loc_col = loc_raw + "_post" if loc_raw and loc_raw + "_post" in merged.columns else loc_raw
    class_col = class_raw + "_post" if class_raw and class_raw + "_post" in merged.columns else class_raw

    tab1_data = {
        "First Name": merged["First name"], 
        "Surname": merged["Surname"], 
        "Email": merged["Email"],
        "Class": merged[class_col].fillna("Unknown") if class_col else "Unknown",
        "Location": merged[loc_col].fillna("Unknown") if loc_col else "Unknown",
        "Form Status": merged.apply(
            lambda r: "Both" if pd.notna(r.get("Timestamp_pre")) and pd.notna(r.get("Timestamp_post"))
            else "Incomplete", axis=1)
    }

    avgs, dists, metrics = [], [], []
    for name, (pk, qk) in METRIC_MAP.items():
        pc, qc = find_col(pre_df, pk), find_col(post_df, qk)
        if not pc or not qc: continue
        mp, mq = pc + "_pre" if pc + "_pre" in merged.columns else pc, qc + "_post" if qc + "_post" in merged.columns else qc
        if mp not in merged.columns or mq not in merged.columns: continue
        
        pre_s, post_s = pd.to_numeric(merged[mp], errors="coerce"), pd.to_numeric(merged[mq], errors="coerce")
        diff = post_s - pre_s
        tab1_data[name] = diff
        metrics.append(name)
        mask = pd.notna(pre_s) & pd.notna(post_s)
        vp, vq, vd = pre_s[mask], post_s[mask], diff[mask]
        
        avgs.append({"Metric": name, "Pre-Camp Avg": vp.mean(), "Post-Camp Avg": vq.mean(), "Avg Shift": vq.mean() - vp.mean()})
        dists.append({"Metric": name, "Improved": len(vd[vd>0])/len(vd) if len(vd) else 0, "Stayed Same": len(vd[vd==0])/len(vd) if len(vd) else 0, "Declined": len(vd[vd<0])/len(vd) if len(vd) else 0})

    tab1_df = pd.DataFrame(tab1_data)
    unique_locs, unique_classes = [l for l in tab1_df["Location"].unique() if str(l).lower() not in ("unknown","nan")], [c for c in tab1_df["Class"].unique() if str(c).lower() not in ("unknown","nan")]
    
    bd_rows = []
    for cls in sorted(unique_classes):
        sub = tab1_df[tab1_df["Class"] == cls]
        row = {"Class": cls, "Location": "ALL LOCATIONS"}
        for m in metrics: row[m] = sub[m].mean()
        bd_rows.append(row)
        for loc in sorted(unique_locs):
            sl = sub[sub["Location"] == loc]
            if not sl.empty:
                row2 = {"Class": f"  ↳ {loc}", "Location": loc}
                for m in metrics: row2[m] = sl[m].mean()
                bd_rows.append(row2)

    return tab1_df, pd.DataFrame(avgs), pd.DataFrame(dists), pd.DataFrame(bd_rows), merged, metrics


# ═══════════════════════════════════════════════════════════════════
#  MAIN REPORT GENERATOR
# ═══════════════════════════════════════════════════════════════════

def generate_report(student_path, pre_path, post_path, status_label, root):
    try:
        status_label.config(text="Status: Loading data…", fg="blue"); root.update()
        students, pre_df, post_df = pd.read_csv(student_path), pd.read_csv(pre_path), pd.read_csv(post_path)
        for df in [students, pre_df, post_df]: df.columns = df.columns.str.strip()
        students["Email"], pre_df["Email address"], post_df["Email address"] = students["Email"].astype(str).str.strip().str.lower(), pre_df["Email address"].astype(str).str.strip().str.lower(), post_df["Email address"].astype(str).str.strip().str.lower()
        
        attend_col = find_col(post_df, ["attend", "camp"])
        if attend_col: post_df = post_df[post_df[attend_col].astype(str).str.lower().str.contains("yes", na=False)]
        pre_df, post_df = pre_df.sort_values("Timestamp").drop_duplicates("Email address", keep="last"), post_df.sort_values("Timestamp").drop_duplicates("Email address", keep="last")

        status_label.config(text="Status: Processing quantitative data…", fg="blue"); root.update()
        tab1_df, avg_df, dist_df, breakdown_df, merged_df, metrics_processed = process_quantitative(students, pre_df, post_df)

        status_label.config(text="Status: Loading AI model (~2.8 GB) - First run takes a few mins...", fg="#D97706"); root.update()
        year, qual_ok = datetime.now().year, False
        
        try:
            from mlx_lm import load
            model, tokenizer = load("mlx-community/gemma-4-e4b-it-4bit")
            email_list, names, surnames = merged_df["Email"].str.lower().tolist(), merged_df["First name"].tolist(), merged_df["Surname"].tolist()
            loc_col, class_col = find_col(post_df, ["location"]), find_col(post_df, ["class"])
            loc_map, class_map = dict(zip(post_df["Email address"], post_df[loc_col])) if loc_col else {}, dict(zip(post_df["Email address"], post_df[class_col])) if class_col else {}
            locations, classes = [str(loc_map.get(e, "Unknown")) for e in email_list], [str(class_map.get(e, "Unknown")) for e in email_list]

            coded_df, narratives = build_coded_df(post_df, email_list, names, surnames, classes, locations, model, tokenizer, status_label, root, year)
            long_df = build_longitudinal_df(coded_df, tab1_df, metrics_processed, year)
            summary_metadata = build_summary_metadata()
            
            mm_tables = build_mixed_methods(long_df)
            qual_charts = build_qual_chart_data(long_df)
            qual_ok = True

        except Exception as ai_err:
            err = f"AI error: {ai_err}"
            status_label.config(text="Status: AI failed — quantitative only", fg="red"); root.update()
            coded_df, long_df = pd.DataFrame([{"Error": err}]), pd.DataFrame([{"Note": err}])
            summary_metadata, narratives = [("SECTION", "AI Error", "", ""), ("DATA", "Error", err, "")], [{"Question": "AI Status", "n": 0, "Summary": err}]
            mm_tables, qual_charts = [], {}

        status_label.config(text="Status: Writing Excel report…", fg="blue"); root.update()
        output = "Camp_Analysis_Report.xlsx"
        write_excel(output, tab1_df, avg_df, dist_df, breakdown_df, coded_df, summary_metadata, long_df, narratives, metrics_processed, mm_tables, qual_charts)

        status_label.config(text="Status: Complete ✓", fg="#006100")
        messagebox.showinfo("Done", f"{'Full qualitative analysis & Dashboards complete.' if qual_ok else '⚠ AI failed.'}\n\nFile: {os.path.abspath(output)}")

    except Exception as e:
        status_label.config(text="Status: Error", fg="red")
        messagebox.showerror("Error", f"{e}\n\n{traceback.format_exc()[-600:]}")


def setup_gui():
    root = tk.Tk()
    root.title("Camp Analysis Tool")
    root.geometry("580x460"); root.resizable(False, False)
    files = {"student": "", "pre": "", "post": ""}
    def pick(key, entry):
        p = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if p: entry.delete(0, tk.END); entry.insert(0, p); files[key] = p
    def row(label, key):
        f = tk.Frame(root); f.pack(pady=7, padx=25, fill="x")
        tk.Label(f, text=label, font=("Arial", 10, "bold")).pack(anchor="w")
        e = tk.Entry(f); e.pack(side="left", fill="x", expand=True, padx=(0, 10))
        tk.Button(f, text="Browse", width=10, command=lambda: pick(key, e)).pack(side="right")

    tk.Label(root, text="Camp Report Generator", font=("Arial", 17, "bold"), pady=12).pack()
    tk.Label(root, text="Gemma 4 E4B | 8 Dashboards | Triangulation", font=("Arial", 9), fg="#555555").pack()
    tk.Frame(root, height=1, bg="#CCCCCC").pack(fill="x", padx=25, pady=8)
    row("1.  Student List CSV", "student"); row("2.  Pre-Camp Survey CSV", "pre"); row("3.  Post-Camp Survey CSV", "post")
    tk.Frame(root, height=1, bg="#CCCCCC").pack(fill="x", padx=25, pady=10)
    status = tk.Label(root, text="Status: Waiting for files…", font=("Arial", 10, "italic"), fg="gray")
    status.pack(pady=(0, 6))
    tk.Button(root, text="▶  GENERATE REPORT", bg="#2E75B6", fg="white", font=("Arial", 13, "bold"), command=lambda: generate_report(files["student"], files["pre"], files["post"], status, root)).pack(pady=6, fill="x", padx=25)
    root.mainloop()

if __name__ == "__main__":
    setup_gui()