#!/usr/bin/env python3
"""
Camp Data Analysis Tool  —  Speed-Optimised Edition (Rigorous NLP & Purging)
=====================================================================
Features:
  • Ghost-Row Purging: Drops non-responders to ensure pristine data.
  • Isolated Experience Sentiment: Prevents constructive feedback from dragging down scores.
  • Tone Analysis: Codes Q6/Q7 for Tone (Constructive, Critical, Practical).
  • Challenge Attitude: Codes Q2 for resilience rather than penalizing hardship.
  • On-disk coding cache: same cohort data = instant re-run.
  • Thinking mode disabled for narratives (removes wasted CoT tokens).
  • Tighter batch sizes and dynamic_max for more reliable JSON output.
  • All Original Dashboards, Highlights, and Premium Formatting Preserved.
"""

import gc, hashlib, json, os, pickle, re, traceback
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
    # Strip Gemma thinking blocks if present
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()

    candidates = [
        text,
        re.sub(r"```(?:json)?", "", text).replace("```", ""),  # strip all code fences
    ]
    for src in candidates:
        src = src.strip()
        try:
            result = json.loads(src)
            if isinstance(result, list):
                return result
        except Exception:
            pass
        # Find the outermost [...] block
        m = re.search(r'\[.*\]', src, re.DOTALL)
        if m:
            try:
                result = json.loads(m.group())
                if isinstance(result, list):
                    return result
            except Exception:
                pass

    preview = text[:300].replace('\n', ' ')
    print(f"\n⚠️  safe_json could not parse model output. Raw preview:\n  {preview}\n")
    return None


# ═══════════════════════════════════════════════════════════════════
#  ON-DISK CODING CACHE
# ═══════════════════════════════════════════════════════════════════

CACHE_DIR     = os.path.expanduser("~/Library/Application Support/Camp_Analysis")
CACHE_FILE    = os.path.join(CACHE_DIR, "coding_cache.pkl")
CACHE_VERSION = "v5-strict-rigorous"

def load_coding_cache():
    try:
        with open(CACHE_FILE, "rb") as f:
            data = pickle.load(f)
        if not isinstance(data, dict) or data.get("__version__") != CACHE_VERSION:
            return {"__version__": CACHE_VERSION}
        return data
    except Exception:
        return {"__version__": CACHE_VERSION}

def save_coding_cache(cache):
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        cache["__version__"] = CACHE_VERSION
        with open(CACHE_FILE, "wb") as f:
            pickle.dump(cache, f)
    except Exception as e:
        print(f"Cache save warning: {e}")

def data_fingerprint(email_list, post_df):
    content = "|".join(sorted(str(e) for e in email_list)) + post_df.to_csv(index=False)
    return hashlib.sha256(content.encode()).hexdigest()


# ═══════════════════════════════════════════════════════════════════
#  REFINED QUALITATIVE CODEBOOK (TONE vs SENTIMENT)
# ═══════════════════════════════════════════════════════════════════

QUAL_QUESTIONS = [
    {
        "num": 1, "label": "Peer Assistance",
        "keywords": ["helped a classmate"],
        "fields": {
            "Q1_Direction": ["Gave-Help", "Received-Help", "Mutual", "Not-Mentioned"],
            "Q1_Type": ["Equipment/Gear", "Campcraft", "Emotional-Support", "Physical-Help", "Taught-a-Skill", "Unclear"],
            "Q1_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": "Code the direction of help and the type of help given."
    },
    {
        "num": 2, "label": "Challenge & Resilience",
        "keywords": ["biggest", "challenge"],
        "fields": {
            "Q2_Challenge": ["Physical-Effort", "Fear-or-Anxiety", "Weather-or-Environment", "Social-Difficulty", "Equipment-or-Skills", "Not-Specified"],
            "Q2_Response": ["Pushed-Through-Alone", "Friends-or-Staff-Helped", "Adapted-Approach", "Just-Tried-It", "Not-Explained"],
            "Q2_Growth": ["Yes", "Partial", "No"],
            "Q2_Attitude": ["Resilient-and-Proud", "Matter-of-Fact", "Overwhelmed-or-Negative"],
        },
        "guide": "Code what the challenge was, how they responded, and their attitude looking back at the hardship."
    },
    {
        "num": 3, "label": "Surprising Skill",
        "keywords": ["surprised yourself"],
        "fields": {
            "Q3_Activity": ["Mountain-Biking", "Sea-Kayaking", "Snorkelling", "Coasteering", "Hiking", "Camp-Cooking", "Swimming", "Other"],
            "Q3_Growth_Type": ["Beat-Own-Expectations", "Overcame-a-Fear", "Got-Noticeably-Better", "Discovered-Enjoyment"],
            "Q3_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": "Code the activity and the type of personal growth."
    },
    {
        "num": 4, "label": "First Nations Culture",
        "keywords": ["first nations", "stood out"],
        "fields": {
            "Q4_Topic": ["Plants-Animals-Country", "Stories-and-History", "Practices-and-Tools", "Bush-Food-and-Knowledge", "Movement-and-Country", "Vague-or-Unclear"],
            "Q4_Depth": ["Recalled-a-Fact", "New-Perspective", "Personal-Connection", "Surface-Level"],
            "Q4_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": "Code the topic remembered and their depth of understanding."
    },
    {
        "num": 5, "label": "Favourite Part",
        "keywords": ["absolute favourite"],
        "fields": {
            "Q5_What": ["A-Specific-Activity", "Social-Time", "Scenery-or-Nature", "Personal-Achievement", "Food-or-Cooking", "Free-Time", "Everything"],
            "Q5_Why": ["Thrill-and-Fun", "Friendship", "Achievement", "Beauty-or-Peace", "Freedom-or-Choice", "Not-Explained"],
            "Q5_Sentiment": ["Positive", "Neutral", "Negative", "Mixed"],
        },
        "guide": "Code WHAT their favourite part was and WHY."
    },
    {
        "num": 6, "label": "Improvement Suggestions",
        "keywords": ["camp director"],
        "fields": {
            "Q6_Suggestion": ["More-Time-or-Activities", "Better-Food", "Better-Gear-or-Comfort", "Schedule-Change", "Happy-With-Everything", "Warmer-Weather-or-Season", "Other"],
            "Q6_Tone": ["Constructive-Suggestion", "Critical-Complaint", "Praising", "Neutral-Observation"],
        },
        "guide": "Code their primary suggestion and the Tone of their feedback (e.g. constructive ideas vs complaining)."
    },
    {
        "num": 7, "label": "Advice for Year 7s",
        "keywords": ["year 7 student"],
        "fields": {
            "Q7_Advice": ["Mindset-and-Attitude", "Pack-and-Gear", "Food-and-Snacks", "Be-Social", "Comfort-and-Sleep", "Safety-and-Health", "Just-Try-Everything", "Not-Specified"],
            "Q7_Tone": ["Encouraging", "Cautionary-or-Warning", "Practical-and-Factual"],
        },
        "guide": "Code the main piece of advice given and the overall tone used to deliver it."
    },
    {
        "num": 8, "label": "7-Day Camp Preparation",
        "keywords": ["better prepare"],
        "fields": {
            "Q8_Prep_Focus": ["Get-Fitter", "Better-Packing", "Practice-Camp-Skills", "Mental-Preparation", "More-School-Support", "Build-On-This-Camp", "Better-Food-Planning", "Not-Specified"],
            "Q8_Growth": ["Yes", "Partial", "No"],
        },
        "guide": "Code what they think is most important to prepare. Sentiment is not needed here."
    },
]

ALL_QUAL_FIELDS = [f for qc in QUAL_QUESTIONS for f in qc["fields"]]
EXPERIENCE_SENTIMENT_FIELDS = ["Q1_Sentiment", "Q3_Sentiment", "Q4_Sentiment", "Q5_Sentiment"]
GROWTH_FIELDS    = ["Q2_Growth", "Q8_Growth"]
SENTIMENT_SCORE = {"Positive": 1.0, "Mixed": 0.25, "Neutral": 0.0, "Negative": -1.0}


# ═══════════════════════════════════════════════════════════════════
#  MLX INFERENCE
# ═══════════════════════════════════════════════════════════════════

def call_mlx(model, tokenizer, system_msg, user_msg, max_tokens=2500, thinking=False):
    from mlx_lm import generate
    messages = [{"role": "system", "content": system_msg},
                {"role": "user",   "content": user_msg}]
    if hasattr(tokenizer, "apply_chat_template"):
        try:
            prompt = tokenizer.apply_chat_template(
                messages, tokenize=False, add_generation_prompt=True,
                enable_thinking=thinking)
        except TypeError:
            prompt = tokenizer.apply_chat_template(
                messages, tokenize=False, add_generation_prompt=True)
    else:
        prompt = f"### System:\n{system_msg}\n\n### User:\n{user_msg}\n\n### Assistant:\n"
    return generate(model, tokenizer, prompt=prompt, max_tokens=max_tokens, verbose=False)


# ═══════════════════════════════════════════════════════════════════
#  PER-QUESTION BATCH CODING
# ═══════════════════════════════════════════════════════════════════

SYSTEM_CODER = (
    "You are coding student survey responses from a school camp. "
    "Return ONLY a valid JSON array — no explanation, no markdown, no preamble. "
    "Use ONLY the listed values for each field. Never invent new values. "
    "If a response is completely nonsensical or off-topic, use the 'Not-Specified', "
    "'Unclear', or 'No' tags provided in the guide."
)

def build_coding_prompt(qc, batch):
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


def code_question(coding_model, coding_tokenizer, qc, post_df, email_list,
                  status_label, root, BATCH=40):
    col    = find_col(post_df, qc["keywords"])
    fields = list(qc["fields"].keys())

    email_resp = dict(zip(post_df["Email address"],
                          post_df.get(col, pd.Series(dtype=str))))
    items = []
    for idx, email in enumerate(email_list):
        resp = str(email_resp.get(email, "")).strip()
        items.append({"id": idx, "text": resp, "empty": len(resp) <= 3})

    coded   = {}
    to_code = [it for it in items if not it["empty"]]
    total_batches = max(1, (len(to_code) + BATCH - 1) // BATCH)

    for bn, start in enumerate(range(0, len(to_code), BATCH), 1):
        batch = to_code[start:start + BATCH]
        status_label.config(
            text=f"Status: Coding Q{qc['num']} ({qc['label']}) — "
                 f"batch {bn}/{total_batches}...",
            fg="#D97706")
        root.update()

        prompt = build_coding_prompt(qc, batch)
        dynamic_max = max(512, int(len(batch) * (len(fields) * 8 + 15) * 1.2) + 64)
        raw    = call_mlx(coding_model, coding_tokenizer, SYSTEM_CODER,
                          prompt, max_tokens=dynamic_max, thinking=False)
        parsed = safe_json(raw)

        if isinstance(parsed, list):
            for row in parsed:
                if isinstance(row, dict) and "id" in row:
                    coded[row["id"]] = {
                        f: str(row.get(f, "Parse-Error")).strip()
                        for f in fields
                    }

        for it in batch:
            if it["id"] not in coded:
                coded[it["id"]] = {f: "Parse-Error" for f in fields}

    for it in items:
        if it["empty"]:
            coded[it["id"]] = {f: "No-Response" for f in fields}

    return coded


# ═══════════════════════════════════════════════════════════════════
#  NARRATIVE GENERATION  (Gemma 4 E4B, thinking disabled)
# ═══════════════════════════════════════════════════════════════════

SYSTEM_NARRATOR = (
    "You are an educational analyst writing a concise summary for a school camp report. "
    "Write in clear paragraphs. Be specific and evidence-based."
)

def generate_narrative(narrative_model, narrative_tokenizer, qc, responses,
                       status_label, root):
    status_label.config(
        text=f"Status: Writing Q{qc['num']} narrative & extracting quotes...",
        fg="#D97706")
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

    return call_mlx(narrative_model, narrative_tokenizer, SYSTEM_NARRATOR,
                    prompt, max_tokens=500, thinking=False).strip()


# ═══════════════════════════════════════════════════════════════════
#  DATASET BUILDERS
# ═══════════════════════════════════════════════════════════════════

def build_coded_df(post_df, email_list, names, surnames, classes, locations,
                   coding_model, coding_tokenizer,
                   narrative_model, narrative_tokenizer,
                   status_label, root, year):

    cache       = load_coding_cache()
    fingerprint = data_fingerprint(email_list, post_df)

    if fingerprint in cache:
        status_label.config(
            text="Status: Cache hit — skipping coding (data unchanged) ✓",
            fg="#006100")
        root.update()
        all_coded, q_resps = cache[fingerprint]
    else:
        q_resps   = {}
        all_coded = {i: {} for i in range(len(email_list))}

        for qc in QUAL_QUESTIONS:
            qnum = qc["num"]
            col  = find_col(post_df, qc["keywords"])

            if col:
                q_resps[qnum] = dict(zip(post_df["Email address"], post_df[col]))
                q_coded = code_question(
                    coding_model, coding_tokenizer,
                    qc, post_df, email_list, status_label, root)
                for idx in range(len(email_list)):
                    all_coded[idx].update(q_coded.get(idx, {
                        f: "Parse-Error" for f in qc["fields"]
                    }))
            else:
                q_resps[qnum] = {}
                for idx in range(len(email_list)):
                    for f in qc["fields"]:
                        all_coded[idx][f] = "Column-Not-Found"

        cache[fingerprint] = (all_coded, q_resps)
        save_coding_cache(cache)

    coded_df = _assemble_coded_df(all_coded, q_resps, email_list, names, surnames, classes, locations, year)
    
    narratives = []
    for qc in QUAL_QUESTIONS:
        qnum = qc["num"]
        raw_vals = [str(q_resps.get(qnum, {}).get(e, "")).strip() for e in email_list]
        if not raw_vals or all(r == "" for r in raw_vals):
            narratives.append({
                "Question": f"Q{qnum}: {qc['label']}",
                "n": 0,
                "Summary": "Column not found in post-camp CSV.",
            })
            continue

        narrative_text = generate_narrative(
            narrative_model, narrative_tokenizer,
            qc, raw_vals, status_label, root)
        valid_n = sum(1 for r in raw_vals if len(r) > 3)
        narratives.append({
            "Question": f"Q{qnum}: {qc['label']}",
            "n": valid_n,
            "Summary": narrative_text,
        })

    return coded_df, narratives


def _assemble_coded_df(all_coded, q_resps, email_list, names, surnames, classes, locations, year):
    coded_df = pd.DataFrame({
        "Year":       year,
        "Student_ID": [student_id(e) for e in email_list],
        "First_Name": names,
        "Surname":    surnames,
        "Class":      classes,
        "Location":   locations,
    })
    for qc in QUAL_QUESTIONS:
        qnum = qc["num"]
        coded_df[f"Q{qnum}_Response"] = [
            str(q_resps.get(qnum, {}).get(e, "")).strip() for e in email_list
        ]
        for field in qc["fields"]:
            coded_df[field] = [
                all_coded.get(i, {}).get(field, "No-Response") for i in range(len(email_list))
            ]
    return coded_df


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
    sent_cols = [f for f in EXPERIENCE_SENTIMENT_FIELDS if f in long_df.columns]
    if sent_cols:
        long_df["Avg_Experience_Sentiment"] = long_df[sent_cols].map(sent_num).mean(axis=1).round(2)
    else:
        long_df["Avg_Experience_Sentiment"] = np.nan

    def growth_num(val): return 1 if val == "Yes" else (0.5 if val == "Partial" else 0)
    g_cols = [f for f in GROWTH_FIELDS if f in long_df.columns]
    if g_cols:
        long_df["Growth_Index"] = long_df[g_cols].map(growth_num).sum(axis=1).round(1)
    else:
        long_df["Growth_Index"] = np.nan

    return long_df


def build_mixed_methods(long_df):
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
    if "Q1_Type" in long_df.columns and "Q1_Direction" in long_df.columns:
        charts["Q1"] = pd.crosstab(long_df["Q1_Type"], long_df["Q1_Direction"]).reset_index()
    if "Q2_Challenge" in long_df.columns and "Q2_Attitude" in long_df.columns:
        charts["Q2"] = pd.crosstab(long_df["Q2_Challenge"], long_df["Q2_Attitude"]).reset_index()
    if "Q3_Activity" in long_df.columns and "Q3_Growth_Type" in long_df.columns:
        charts["Q3"] = pd.crosstab(long_df["Q3_Activity"], long_df["Q3_Growth_Type"]).reset_index()
    if "Q4_Topic" in long_df.columns and "Q4_Depth" in long_df.columns:
        charts["Q4"] = pd.crosstab(long_df["Q4_Topic"], long_df["Q4_Depth"]).reset_index()
    if "Q5_What" in long_df.columns and "Q5_Sentiment" in long_df.columns:
        charts["Q5"] = pd.crosstab(long_df["Q5_What"], long_df["Q5_Sentiment"]).reset_index()
    if "Q6_Suggestion" in long_df.columns and "Q6_Tone" in long_df.columns: 
        charts["Q6"] = pd.crosstab(long_df["Q6_Suggestion"], long_df["Q6_Tone"]).reset_index()
    if "Q7_Advice" in long_df.columns and "Q7_Tone" in long_df.columns: 
        charts["Q7"] = pd.crosstab(long_df["Q7_Advice"], long_df["Q7_Tone"]).reset_index()
    if "Q8_Prep_Focus" in long_df.columns and "Q8_Growth" in long_df.columns:
        charts["Q8"] = pd.crosstab(long_df["Q8_Prep_Focus"], long_df["Q8_Growth"]).reindex(columns=["Yes", "Partial", "No"], fill_value=0).reset_index()
    if "Avg_Experience_Sentiment" in long_df.columns and "Class" in long_df.columns:
        charts["Q9"] = long_df.groupby("Class")["Avg_Experience_Sentiment"].mean().reset_index().round(2)
    return charts


# ═══════════════════════════════════════════════════════════════════
#  EXCEL WRITER (PREMIUM FORMATTING MAINTAINED)
# ═══════════════════════════════════════════════════════════════════

def write_excel(output_path, tab1_df, avg_df, dist_df, breakdown_df,
                coded_df, summary_metadata, long_df, narratives, metrics_processed, mm_tables, qual_charts):

    writer = pd.ExcelWriter(output_path, engine="xlsxwriter")
    workbook = writer.book
    workbook.nan_inf_to_errors = True 

    F = {
        "head":      workbook.add_format({"bold":True,"bg_color":"#D9D9D9","border":1}),
        "head_grn":  workbook.add_format({"bold":True,"bg_color":"#E2EFDA","border":1,"text_wrap":True}),
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

    # ── Write DataFrames to sheets
    tab1_df.to_excel(writer,      sheet_name="Individual Scores", index=False, na_rep="—")
    avg_df.to_excel(writer,       sheet_name="Group Averages",    index=False, na_rep="—")
    breakdown_df.to_excel(writer, sheet_name="Group Breakdown",   index=False, na_rep="—")
    coded_df.to_excel(writer,     sheet_name="Qual Responses",    index=False, na_rep="—")
    long_df.to_excel(writer,      sheet_name="Longitudinal Data", index=False, na_rep="—")
    dist_df.to_excel(writer,      sheet_name="_ChartData",        index=False, na_rep="0")

    # ──────────────────────────────────────────────────────────────
    #  Visual Summary
    # ──────────────────────────────────────────────────────────────
    ws_charts = workbook.add_worksheet("Visual Summary")
    ws_charts.hide_gridlines(2)
    writer.sheets["Visual Summary"] = ws_charts
    ws_charts.set_first_sheet()
    ws_charts.activate()

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
    #  Qual Responses  (Dropdowns preserved)
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
    #  Qual Summary (Dynamic Formulas preserved)
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
    #  Qual Highlights (Narratives + Quotes preserved)
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
    #  Mixed Methods Insights preserved
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
    #  Longitudinal Data (Format tracking preserved)
    # ──────────────────────────────────────────────────────────────
    ws = writer.sheets["Longitudinal Data"]
    for i, col in enumerate(long_df.columns):
        ws.write(0, i, col, F["long_head"])
    ws.set_column("A:A", 8); ws.set_column("B:B", 14); ws.set_column("C:E", 16); ws.set_column("F:F", 16)  
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
    #  Qual Dashboard (Upgraded to 9-Chart Cross-Tab Matrix)
    # ──────────────────────────────────────────────────────────────
    ws_qchart = workbook.add_worksheet("_QualChartData")
    ws_qchart.hide()
    
    ws_dash = workbook.add_worksheet("Qual Dashboard")
    writer.sheets["Qual Dashboard"] = ws_dash
    ws_dash.hide_gridlines(2)

    col_cursor = 0
    def insert_dash_chart(df, chart_type, title, pos_cell, subtype=None):
        nonlocal col_cursor
        df.to_excel(writer, sheet_name="_QualChartData", startcol=col_cursor, index=False)
        opts = {'type': chart_type}
        if subtype: opts['subtype'] = subtype
        c = workbook.add_chart(opts)
        
        for i, name in enumerate(df.columns[1:]):
            c.add_series({
                'name': name,
                'categories': ['_QualChartData', 1, col_cursor, len(df), col_cursor],
                'values': ['_QualChartData', 1, col_cursor+1+i, len(df), col_cursor+1+i]
            })
        c.set_title({'name': title, 'name_font': {'size': 12}})
        c.set_size({'width': 480, 'height': 280})
        ws_dash.insert_chart(pos_cell, c)
        col_cursor += len(df.columns) + 1

    if "Q1" in qual_charts: insert_dash_chart(qual_charts["Q1"], 'column', 'Q1: Help Type by Direction', 'B2', subtype='stacked')
    if "Q2" in qual_charts: insert_dash_chart(qual_charts["Q2"], 'column', 'Q2: Hardship Challenge vs Student Attitude', 'J2', subtype='percent_stacked')
    if "Q3" in qual_charts: insert_dash_chart(qual_charts["Q3"], 'bar', 'Q3: Surprising Skills Discovered', 'R2', subtype='stacked')
    
    if "Q4" in qual_charts: insert_dash_chart(qual_charts["Q4"], 'bar', 'Q4: First Nations Depth', 'B17', subtype='stacked')
    if "Q5" in qual_charts: insert_dash_chart(qual_charts["Q5"], 'column', 'Q5: Favourite Part by Sentiment', 'J17', subtype='stacked')
    if "Q6" in qual_charts: insert_dash_chart(qual_charts["Q6"], 'bar', 'Q6: Feedback Tone (Constructive vs Critical)', 'R17', subtype='stacked')
    
    if "Q7" in qual_charts: insert_dash_chart(qual_charts["Q7"], 'bar', 'Q7: Advice Delivery Style', 'B32', subtype='stacked')
    if "Q8" in qual_charts: insert_dash_chart(qual_charts["Q8"], 'column', 'Q8: Future Prep vs Growth Mindset', 'J32', subtype='percent_stacked')
    if "Q9" in qual_charts: insert_dash_chart(qual_charts["Q9"], 'column', 'Cohort Health: True Experience Sentiment by Class', 'R32')

    writer.close()


# ═══════════════════════════════════════════════════════════════════
#  QUANTITATIVE PROCESSING (STRICT PURGING INCLUDED)
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
    # STRICT FILTER: Only process students who actually completed the post survey.
    post_emails = post_df["Email address"].dropna().unique()
    students_df = students_df[students_df["Email"].isin(post_emails)]
    
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
        status_label.config(text="Status: Loading & Purging Data…", fg="blue"); root.update()
        students, pre_df, post_df = pd.read_csv(student_path), pd.read_csv(pre_path), pd.read_csv(post_path)
        for df in [students, pre_df, post_df]: df.columns = df.columns.str.strip()
        students["Email"], pre_df["Email address"], post_df["Email address"] = students["Email"].astype(str).str.strip().str.lower(), pre_df["Email address"].astype(str).str.strip().str.lower(), post_df["Email address"].astype(str).str.strip().str.lower()
        
        attend_col = find_col(post_df, ["attend", "camp"])
        if attend_col: post_df = post_df[post_df[attend_col].astype(str).str.lower().str.contains("yes", na=False)]
        
        pre_df = pre_df.sort_values("Timestamp").drop_duplicates("Email address", keep="last")
        post_df = post_df.sort_values("Timestamp").drop_duplicates("Email address", keep="last")

        status_label.config(text="Status: Processing quantitative data…", fg="blue"); root.update()
        tab1_df, avg_df, dist_df, breakdown_df, merged_df, metrics_processed = process_quantitative(students, pre_df, post_df)

        status_label.config(text="Status: Loading Gemma 4 E4B Model (~2.8 GB)...", fg="#D97706"); root.update()
        year = datetime.now().year
        
        try:
            from mlx_lm import load
            
            # Use the local model
            model, tokenizer = load("FakeRockert543/gemma-4-e4b-it-MLX-4bit")
            
            email_list, names, surnames = merged_df["Email"].str.lower().tolist(), merged_df["First name"].tolist(), merged_df["Surname"].tolist()
            loc_col, class_col = find_col(post_df, ["location"]), find_col(post_df, ["class"])
            loc_map = dict(zip(post_df["Email address"], post_df[loc_col])) if loc_col else {}
            class_map = dict(zip(post_df["Email address"], post_df[class_col])) if class_col else {}
            locations = [str(loc_map.get(e, "Unknown")) for e in email_list]
            classes = [str(class_map.get(e, "Unknown")) for e in email_list]

            coded_df, narratives = build_coded_df(
                post_df, email_list, names, surnames, classes, locations,
                model, tokenizer, model, tokenizer,
                status_label, root, year)

            del model, tokenizer
            gc.collect()

            long_df = build_longitudinal_df(coded_df, merged_df, metrics_processed, year)
            summary_metadata = build_summary_metadata()
            mm_tables = build_mixed_methods(long_df)
            qual_charts = build_qual_chart_data(long_df)

        except Exception as ai_err:
            err = f"AI error: {ai_err}"
            status_label.config(text="Status: AI failed — quantitative only", fg="red"); root.update()
            coded_df, long_df = pd.DataFrame([{"Error": err}]), pd.DataFrame([{"Note": err}])
            summary_metadata, narratives = [("SECTION", "AI Error", "", ""), ("DATA", "Error", err, "")], [{"Question": "AI Status", "n": 0, "Summary": err}]
            mm_tables, qual_charts = [], {}

        status_label.config(text="Status: Writing 9-Chart Excel Report…", fg="blue"); root.update()
        output = "Camp_Analysis_Report.xlsx"
        write_excel(output, tab1_df, avg_df, dist_df, breakdown_df, coded_df, summary_metadata, long_df, narratives, metrics_processed, mm_tables, qual_charts)

        status_label.config(text="Status: Complete ✓", fg="#006100")
        messagebox.showinfo("Done", f"Full rigorous qualitative analysis & Dashboards complete.\n\nFile: {os.path.abspath(output)}")

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
    tk.Label(root, text="Gemma 4 E4B | Cache + Fast Settings | Rigorous 9 Dashboards", font=("Arial", 9), fg="#555555").pack()
    tk.Frame(root, height=1, bg="#CCCCCC").pack(fill="x", padx=25, pady=8)
    row("1.  Student List CSV", "student"); row("2.  Pre-Camp Survey CSV", "pre"); row("3.  Post-Camp Survey CSV", "post")
    tk.Frame(root, height=1, bg="#CCCCCC").pack(fill="x", padx=25, pady=10)
    status = tk.Label(root, text="Status: Waiting for files…", font=("Arial", 10, "italic"), fg="gray")
    status.pack(pady=(0, 6))
    tk.Button(root, text="▶  GENERATE REPORT", bg="#2E75B6", fg="white", font=("Arial", 13, "bold"), command=lambda: generate_report(files["student"], files["pre"], files["post"], status, root)).pack(pady=6, fill="x", padx=25)
    root.mainloop()

if __name__ == "__main__":
    setup_gui()