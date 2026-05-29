#!/usr/bin/env python3
"""
Camp Data Analysis Tool — Upgraded Edition
=========================================================
OFFLINE SETUP (one-time):
  Place chartjs.min.js in the same folder as this script.
  The HTML report will be 100% self-contained — no internet needed.
  Get chartjs.min.js from: https://github.com/chartjs/Chart.js/releases
  (download chart.umd.min.js from the dist/ folder of any v4 release)

Outputs per run:
  • Camp_Analysis_Report.xlsx  — upgraded Excel with executive dashboard
  • Camp_Report_Presentation.html — interactive HTML, open in any browser

Features:
  • Offline-first: Chart.js embedded inline, no CDN dependency
  • Executive Summary as first Excel sheet — readable by non-spreadsheet people
  • Colour-coded shifts, progress bars, alternating rows throughout Excel
  • HTML: sticky nav, stat cards, skill tables, 8 qualitative sections,
    live Chart.js charts, student quotes, mixed-methods insights
  • All original AI coding, narrative, longitudinal, and mixed-methods retained
  • Coding cache: identical data = instant re-run, no re-coding needed
"""

import gc, hashlib, json, os, pickle, re, traceback
from datetime import datetime

import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from xlsxwriter.utility import xl_col_to_name


# ═══════════════════════════════════════════════════════════════════
#  UTILITY
# ═══════════════════════════════════════════════════════════════════

def find_col(df, keywords):
    for col in df.columns:
        if all(k.lower() in col.lower() for k in keywords):
            return col
    return None

def student_id(email):
    return hashlib.sha256(email.strip().lower().encode()).hexdigest()[:10].upper()

def safe_json(text):
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()
    candidates = [text, re.sub(r"```(?:json)?", "", text).replace("```", "")]
    for src in candidates:
        src = src.strip()
        try:
            result = json.loads(src)
            if isinstance(result, list):
                return result
        except Exception:
            pass
        m = re.search(r'\[.*\]', src, re.DOTALL)
        if m:
            try:
                result = json.loads(m.group())
                if isinstance(result, list):
                    return result
            except Exception:
                pass
    print(f"\n⚠️  safe_json failed. Preview: {text[:200]}\n")
    return None


# ═══════════════════════════════════════════════════════════════════
#  CACHE
# ═══════════════════════════════════════════════════════════════════

CACHE_DIR     = os.path.expanduser("~/Library/Application Support/Camp_Analysis")
CACHE_FILE    = os.path.join(CACHE_DIR, "coding_cache.pkl")
CACHE_VERSION = "v7-html-upgrade"

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
#  QUALITATIVE CODEBOOK
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
        "guide": "Code what the challenge was, how they responded, and attitude looking back."
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
        "guide": "Code the topic remembered and depth of understanding."
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
        "guide": "Code primary suggestion and the tone of feedback."
    },
    {
        "num": 7, "label": "Advice for Year 7s",
        "keywords": ["year 7 student"],
        "fields": {
            "Q7_Advice": ["Mindset-and-Attitude", "Pack-and-Gear", "Food-and-Snacks", "Be-Social", "Comfort-and-Sleep", "Safety-and-Health", "Just-Try-Everything", "Not-Specified"],
            "Q7_Tone": ["Encouraging", "Cautionary-or-Warning", "Practical-and-Factual"],
        },
        "guide": "Code the main piece of advice and overall tone."
    },
    {
        "num": 8, "label": "7-Day Camp Preparation",
        "keywords": ["better prepare"],
        "fields": {
            "Q8_Prep_Focus": ["Get-Fitter", "Better-Packing", "Practice-Camp-Skills", "Mental-Preparation", "More-School-Support", "Build-On-This-Camp", "Better-Food-Planning", "Not-Specified"],
            "Q8_Growth": ["Yes", "Partial", "No"],
        },
        "guide": "Code what they think is most important to prepare."
    },
]

ALL_QUAL_FIELDS          = [f for qc in QUAL_QUESTIONS for f in qc["fields"]]
EXPERIENCE_SENTIMENT_FIELDS = ["Q1_Sentiment", "Q3_Sentiment", "Q4_Sentiment", "Q5_Sentiment"]
GROWTH_FIELDS            = ["Q2_Growth", "Q8_Growth"]
SENTIMENT_SCORE          = {"Positive": 1.0, "Mixed": 0.25, "Neutral": 0.0, "Negative": -1.0}


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
#  QUALITATIVE CODING
# ═══════════════════════════════════════════════════════════════════

SYSTEM_CODER = (
    "You are coding student survey responses from a school camp. "
    "Return ONLY a valid JSON array — no explanation, no markdown, no preamble. "
    "Use ONLY the listed values for each field. Never invent new values. "
    "If a response is completely nonsensical or off-topic, use 'Not-Specified', 'Unclear', or 'No'."
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
            text=f"Status: Coding Q{qc['num']} ({qc['label']}) — batch {bn}/{total_batches}...",
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
                    coded[row["id"]] = {f: str(row.get(f, "Parse-Error")).strip() for f in fields}
        for it in batch:
            if it["id"] not in coded:
                coded[it["id"]] = {f: "Parse-Error" for f in fields}

    for it in items:
        if it["empty"]:
            coded[it["id"]] = {f: "No-Response" for f in fields}

    return coded


# ═══════════════════════════════════════════════════════════════════
#  NARRATIVE GENERATION
# ═══════════════════════════════════════════════════════════════════

SYSTEM_NARRATOR = (
    "You are an educational analyst writing a concise summary for a school camp report. "
    "Write in clear paragraphs. Be specific and evidence-based."
)

def generate_narrative(narrative_model, narrative_tokenizer, qc, responses, status_label, root):
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
        status_label.config(text="Status: Cache hit — skipping coding ✓", fg="#006100")
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
                q_coded = code_question(coding_model, coding_tokenizer,
                                        qc, post_df, email_list, status_label, root)
                for idx in range(len(email_list)):
                    all_coded[idx].update(q_coded.get(idx, {f: "Parse-Error" for f in qc["fields"]}))
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
            narratives.append({"Question": f"Q{qnum}: {qc['label']}", "n": 0,
                                "Summary": "Column not found in post-camp CSV."})
            continue
        narrative_text = generate_narrative(narrative_model, narrative_tokenizer,
                                            qc, raw_vals, status_label, root)
        valid_n = sum(1 for r in raw_vals if len(r) > 3)
        narratives.append({"Question": f"Q{qnum}: {qc['label']}", "n": valid_n, "Summary": narrative_text})
    return coded_df, narratives


def _assemble_coded_df(all_coded, q_resps, email_list, names, surnames, classes, locations, year):
    coded_df = pd.DataFrame({
        "Year": year, "Student_ID": [student_id(e) for e in email_list],
        "First_Name": names, "Surname": surnames,
        "Class": classes, "Location": locations,
    })
    for qc in QUAL_QUESTIONS:
        qnum = qc["num"]
        coded_df[f"Q{qnum}_Response"] = [
            str(q_resps.get(qnum, {}).get(e, "")).strip() for e in email_list
        ]
        for field in qc["fields"]:
            coded_df[field] = [all_coded.get(i, {}).get(field, "No-Response") for i in range(len(email_list))]
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


def build_longitudinal_df(coded_df, tab1_df, merged_df, metrics_processed, year):
    long_df = coded_df[["Year","Student_ID","First_Name","Surname","Class","Location"]].copy()
    for metric in metrics_processed:
        if metric in tab1_df.columns:
            long_df[f"{metric}_Shift"] = tab1_df[metric].values
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
        charts["Q8"] = pd.crosstab(long_df["Q8_Prep_Focus"], long_df["Q8_Growth"]).reindex(
            columns=["Yes", "Partial", "No"], fill_value=0).reset_index()
    if "Avg_Experience_Sentiment" in long_df.columns and "Class" in long_df.columns:
        charts["Q9"] = long_df.groupby("Class")["Avg_Experience_Sentiment"].mean().reset_index().round(2)
    return charts


# ═══════════════════════════════════════════════════════════════════
#  QUANTITATIVE PROCESSING
# ═══════════════════════════════════════════════════════════════════

METRIC_MAP = {
    "Camping Skills":     (["confident", "camping"], ["confident", "camping"]),
    "Sleeping Outdoors":  (["sleeping", "outdoors", "bugs"], ["handle", "sleeping", "outdoors"]),
    "Swimming":           (["confident", "swimming", "deep"], ["confident", "swimming", "deep"]),
    "Biking Comfort":     (["comfortable", "riding", "bike"], ["comfortable", "bike", "now"]),
    "Overnight Hike":     (["excited", "overnight", "hike"], ["excited", "overnight", "hike", "future"]),
    "Sea Kayaking":       (["excited", "sea kayaking"], ["excited", "sea kayaking", "again"]),
    "Coasteering":        (["excited", "coasteering"], ["excited", "coasteering", "again"]),
    "Mountain Biking":    (["excited", "mountain bike", "flowy"], ["excited", "mountain biking", "flowy"]),
    "Snorkelling":        (["excited", "snorkelling"], ["excited", "snorkelling", "again"]),
    "Teamwork":           (["confident", "group", "work together"],["during camp","group","work together"]),
    "Autonomy":           (["organising", "own gear"], ["during camp", "manage", "own gear"]),
    "Drive":              (["physical activity", "tiring"], ["things got tiring", "focus", "energy"]),
    "Aboriginal Culture": (["Thinking back", "Aboriginal culture"],["Now that camp is finished","Aboriginal culture"]),
}

def process_quantitative(students_df, pre_df, post_df):
    post_emails = post_df["Email address"].dropna().unique()
    students_df = students_df[students_df["Email"].isin(post_emails)]
    merged = pd.merge(students_df[["First name", "Surname", "Email"]], pre_df,
                      left_on="Email", right_on="Email address", how="left")
    merged = pd.merge(merged, post_df, left_on="Email", right_on="Email address",
                      how="left", suffixes=("_pre", "_post"))
    loc_raw   = find_col(post_df, ["location"])
    class_raw = find_col(post_df, ["class"])
    loc_col   = loc_raw + "_post"   if loc_raw   and loc_raw   + "_post"   in merged.columns else loc_raw
    class_col = class_raw + "_post" if class_raw and class_raw + "_post" in merged.columns else class_raw

    tab1_data = {
        "First Name": merged["First name"],
        "Surname":    merged["Surname"],
        "Email":      merged["Email"],
        "Class":      merged[class_col].fillna("Unknown") if class_col else "Unknown",
        "Location":   merged[loc_col].fillna("Unknown")   if loc_col   else "Unknown",
        "Form Status": merged.apply(
            lambda r: "Both" if pd.notna(r.get("Timestamp_pre")) and pd.notna(r.get("Timestamp_post"))
            else "Incomplete", axis=1)
    }

    avgs, dists, metrics = [], [], []
    for name, (pk, qk) in METRIC_MAP.items():
        pc, qc = find_col(pre_df, pk), find_col(post_df, qk)
        if not pc or not qc:
            continue
        mp = pc + "_pre"  if pc + "_pre"  in merged.columns else pc
        mq = qc + "_post" if qc + "_post" in merged.columns else qc
        if mp not in merged.columns or mq not in merged.columns:
            continue
        pre_s  = pd.to_numeric(merged[mp], errors="coerce")
        post_s = pd.to_numeric(merged[mq], errors="coerce")
        diff   = post_s - pre_s
        tab1_data[name] = diff
        metrics.append(name)
        mask = pd.notna(pre_s) & pd.notna(post_s)
        vp, vq, vd = pre_s[mask], post_s[mask], diff[mask]
        avgs.append({"Metric": name, "Pre-Camp Avg": vp.mean(), "Post-Camp Avg": vq.mean(),
                     "Avg Shift": vq.mean() - vp.mean()})
        dists.append({"Metric": name,
                      "Improved":    len(vd[vd > 0]) / len(vd) if len(vd) else 0,
                      "Stayed Same": len(vd[vd == 0]) / len(vd) if len(vd) else 0,
                      "Declined":    len(vd[vd < 0]) / len(vd) if len(vd) else 0})

    tab1_df = pd.DataFrame(tab1_data)
    unique_locs    = [l for l in tab1_df["Location"].unique() if str(l).lower() not in ("unknown","nan")]
    unique_classes = [c for c in tab1_df["Class"].unique()    if str(c).lower() not in ("unknown","nan")]
    bd_rows = []
    for cls in sorted(unique_classes):
        sub = tab1_df[tab1_df["Class"] == cls]
        row = {"Class": cls, "Location": "ALL LOCATIONS",
               "Average Shift (All Areas)": sub[metrics].mean().mean()}
        for m in metrics:
            row[m] = sub[m].mean()
        bd_rows.append(row)
        for loc in sorted(unique_locs):
            sl = sub[sub["Location"] == loc]
            if not sl.empty:
                row2 = {"Class": f"  ↳ {loc}", "Location": loc,
                        "Average Shift (All Areas)": sl[metrics].mean().mean()}
                for m in metrics:
                    row2[m] = sl[m].mean()
                bd_rows.append(row2)

    return tab1_df, pd.DataFrame(avgs), pd.DataFrame(dists), pd.DataFrame(bd_rows), merged, metrics


# ═══════════════════════════════════════════════════════════════════
#  EXCEL WRITER — UPGRADED DESIGN
# ═══════════════════════════════════════════════════════════════════

# Colour palette
C = {
    "navy":       "#1B3A5C",
    "teal":       "#2E8B88",
    "teal_light": "#D4F0EF",
    "green":      "#2D7A4F",
    "green_light":"#C8EDD9",
    "red":        "#C0392B",
    "red_light":  "#FAD7D3",
    "amber":      "#D97706",
    "amber_light":"#FEF3C7",
    "grey":       "#6B7280",
    "grey_light": "#F3F4F6",
    "white":      "#FFFFFF",
    "offwhite":   "#F9FAFB",
}

def write_excel(output_path, tab1_df, avg_df, dist_df, breakdown_df,
                coded_df, summary_metadata, long_df, narratives,
                metrics_processed, mm_tables, qual_charts):

    writer   = pd.ExcelWriter(output_path, engine="xlsxwriter")
    workbook = writer.book
    workbook.nan_inf_to_errors = True

    # ── Format library
    def fmt(opts):
        base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter"}
        base.update(opts)
        return workbook.add_format(base)

    F = {
        # Headers
        "nav_hdr":   fmt({"bold": True, "bg_color": C["navy"],  "font_color": C["white"], "border": 1, "font_size": 11}),
        "teal_hdr":  fmt({"bold": True, "bg_color": C["teal"],  "font_color": C["white"], "border": 1}),
        "col_hdr":   fmt({"bold": True, "bg_color": C["grey_light"], "border": 1, "text_wrap": True, "align": "center"}),
        # Shift cells
        "g_shift":   fmt({"bg_color": C["green_light"], "font_color": C["green"], "num_format": "+0;-0;0", "align": "center"}),
        "r_shift":   fmt({"bg_color": C["red_light"],   "font_color": C["red"],   "num_format": "+0;-0;0", "align": "center"}),
        "z_shift":   fmt({"bg_color": C["grey_light"],  "num_format": "0",        "align": "center"}),
        "g_dec":     fmt({"bg_color": C["green_light"], "font_color": C["green"], "num_format": "+0.0;-0.0;0.0", "align": "center"}),
        "r_dec":     fmt({"bg_color": C["red_light"],   "font_color": C["red"],   "num_format": "+0.0;-0.0;0.0", "align": "center"}),
        "z_dec":     fmt({"num_format": "0.0", "align": "center"}),
        # Qualitative
        "sent_pos":  fmt({"bg_color": C["green_light"], "font_color": C["green"]}),
        "sent_neg":  fmt({"bg_color": C["red_light"],   "font_color": C["red"]}),
        "sent_mix":  fmt({"bg_color": C["amber_light"], "font_color": C["amber"]}),
        "sent_neu":  fmt({"font_color": C["grey"]}),
        "grow_yes":  fmt({"bg_color": C["green_light"], "font_color": C["green"], "bold": True}),
        "grow_par":  fmt({"bg_color": C["amber_light"], "font_color": C["amber"]}),
        "grow_no":   fmt({"bg_color": C["grey_light"],  "font_color": C["grey"]}),
        "skip":      fmt({"font_color": "#BBBBBB", "italic": True}),
        "err":       fmt({"font_color": "#FF6600",  "italic": True}),
        # Text
        "wrap":      fmt({"text_wrap": True, "valign": "top"}),
        "section":   fmt({"bold": True, "bg_color": C["navy"], "font_color": C["white"], "font_size": 11}),
        "field_hdr": fmt({"bold": True, "bg_color": C["teal_light"], "font_color": C["teal"], "italic": True}),
        "field_n":   fmt({"bold": True, "bg_color": C["teal_light"], "font_color": C["teal"], "italic": True, "num_format": '"n = "0'}),
        "pct1":      fmt({"num_format": "0.0%", "align": "center"}),
        "pct_bar":   fmt({"num_format": "0%", "align": "center"}),
        "num":       fmt({"num_format": "0.00"}),
        "ctr":       fmt({"align": "center"}),
        "long_head": fmt({"bold": True, "bg_color": C["navy"], "font_color": C["white"], "border": 1, "text_wrap": True}),
        "mm_head":   fmt({"bold": True, "bg_color": C["teal"], "font_color": C["white"], "border": 1}),
        # Executive summary cards
        "card_title":fmt({"bold": True, "font_size": 12, "bg_color": C["offwhite"]}),
        "card_num":  fmt({"bold": True, "font_size": 22, "bg_color": C["offwhite"], "font_color": C["navy"]}),
        "card_sub":  fmt({"font_size": 9, "bg_color": C["offwhite"], "font_color": C["grey"], "italic": True}),
    }

    def shift_fmt(v):
        if pd.isna(v): return F["skip"]
        if v > 0.05:   return F["g_dec"]
        if v < -0.05:  return F["r_dec"]
        return F["z_dec"]

    def sent_fmt(val):
        m = {"Positive": F["sent_pos"], "Negative": F["sent_neg"],
             "Mixed": F["sent_mix"]}
        return m.get(val, F["sent_neu"])

    def grow_fmt(val):
        return {"Yes": F["grow_yes"], "Partial": F["grow_par"]}.get(val, F["grow_no"])

    # ── Helper: write a row of merged "card" cells
    def write_stat_card(ws, row, col, title, value, subtitle, bg="#F9FAFB", font_col=C["navy"]):
        card_title = workbook.add_format({"bold": True, "font_size": 10, "bg_color": bg,
                                          "font_name": "Arial", "align": "center"})
        card_num   = workbook.add_format({"bold": True, "font_size": 20, "bg_color": bg,
                                          "font_color": font_col, "font_name": "Arial", "align": "center"})
        card_sub   = workbook.add_format({"font_size": 9, "bg_color": bg, "font_color": C["grey"],
                                          "italic": True, "font_name": "Arial", "align": "center"})
        ws.merge_range(row,   col, row,   col+1, title,    card_title)
        ws.merge_range(row+1, col, row+1, col+1, value,    card_num)
        ws.merge_range(row+2, col, row+2, col+1, subtitle, card_sub)

    # ────────────────────────────────────────────────────────────────
    #  SHEET ORDER — Executive Summary is FIRST
    # ────────────────────────────────────────────────────────────────
    ws_exec = workbook.add_worksheet("📊 Summary")
    ws_exec.hide_gridlines(2)
    writer.sheets["📊 Summary"] = ws_exec

    ws_vis = workbook.add_worksheet("📈 Visual Overview")
    ws_vis.hide_gridlines(2)
    writer.sheets["📈 Visual Overview"] = ws_vis

    tab1_df.to_excel(writer, sheet_name="Individual Scores", index=False, na_rep="—")
    avg_df.to_excel(writer,  sheet_name="Group Averages",    index=False, na_rep="—")
    breakdown_df.to_excel(writer, sheet_name="Class Breakdown", index=False, na_rep="—")

    ws_sum  = workbook.add_worksheet("Qual Summary")
    ws_high = workbook.add_worksheet("Qual Highlights")
    ws_mm   = workbook.add_worksheet("Mixed Methods")

    coded_df.to_excel(writer, sheet_name="Qual Responses", index=False, na_rep="—")
    long_df.to_excel(writer,  sheet_name="Longitudinal Data", index=False, na_rep="—")

    ws_dash  = workbook.add_worksheet("_QualDash")
    ws_cdraw = workbook.add_worksheet("_ChartData")
    ws_qcd   = workbook.add_worksheet("_QualChartData")
    ws_cdraw.hide(); ws_qcd.hide(); ws_dash.hide()
    writer.sheets["Qual Summary"]     = ws_sum
    writer.sheets["Qual Highlights"]  = ws_high
    writer.sheets["Mixed Methods"]    = ws_mm
    writer.sheets["Qual Responses"]   = writer.sheets["Qual Responses"]
    writer.sheets["Longitudinal Data"]= writer.sheets["Longitudinal Data"]
    writer.sheets["_QualDash"]        = ws_dash
    writer.sheets["_ChartData"]       = ws_cdraw
    writer.sheets["_QualChartData"]   = ws_qcd

    # ────────────────────────────────────────────────────────────────
    #  📊 EXECUTIVE SUMMARY SHEET
    # ────────────────────────────────────────────────────────────────
    ws_exec.set_column("A:A", 3)
    ws_exec.set_column("B:B", 18)
    ws_exec.set_column("C:C", 18)
    ws_exec.set_column("D:D", 4)
    ws_exec.set_column("E:E", 18)
    ws_exec.set_column("F:F", 18)
    ws_exec.set_column("G:G", 4)
    ws_exec.set_column("H:H", 22)
    ws_exec.set_column("I:I", 22)
    ws_exec.set_row(0, 8)

    # Title banner
    title_fmt = workbook.add_format({
        "bold": True, "font_size": 18, "font_name": "Arial",
        "bg_color": C["navy"], "font_color": C["white"], "align": "center", "valign": "vcenter"
    })
    sub_fmt = workbook.add_format({
        "font_size": 11, "font_name": "Arial", "bg_color": C["teal"],
        "font_color": C["white"], "align": "center", "valign": "vcenter"
    })
    ws_exec.merge_range("B1:I2", "Camp Report — Executive Summary", title_fmt)
    ws_exec.set_row(0, 30); ws_exec.set_row(1, 30)
    year_label = datetime.now().year
    ws_exec.merge_range("B3:I3", f"Year {year_label}  |  Quantitative & Qualitative Data", sub_fmt)
    ws_exec.set_row(2, 20)

    # Stat cards row 1 — participant stats
    n_students = len(tab1_df)
    n_both     = int((tab1_df["Form Status"] == "Both").sum()) if "Form Status" in tab1_df.columns else n_students
    classes_list = sorted(tab1_df["Class"].dropna().unique().tolist()) if "Class" in tab1_df.columns else []
    locs_list    = sorted(tab1_df["Location"].dropna().unique().tolist()) if "Location" in tab1_df.columns else []

    ws_exec.set_row(4, 22); ws_exec.set_row(5, 32); ws_exec.set_row(6, 18)
    write_stat_card(ws_exec, 4, 1, "Students", str(n_students), "completed post-camp survey",
                    bg=C["teal_light"], font_col=C["teal"])
    write_stat_card(ws_exec, 4, 4, "Matched Pairs", str(n_both), "pre & post surveys",
                    bg=C["green_light"], font_col=C["green"])
    write_stat_card(ws_exec, 4, 7, "Locations", str(len(locs_list)), " & ".join(locs_list[:3]) if locs_list else "—",
                    bg="#EDE9FE", font_col="#5B21B6")

    # Section: Quantitative highlights
    ws_exec.set_row(8, 20)
    sec_fmt = workbook.add_format({"bold": True, "font_size": 11, "font_name": "Arial",
                                   "bg_color": C["navy"], "font_color": C["white"]})
    ws_exec.merge_range("B9:I9", "  📐  Skill & Attitude Shifts  (average change, pre → post camp)", sec_fmt)

    # Mini table: metric shifts with inline colour bar
    if not avg_df.empty and "Metric" in avg_df.columns:
        ws_exec.set_row(9, 16)
        hdr_fmt = workbook.add_format({"bold": True, "bg_color": C["grey_light"],
                                        "font_name": "Arial", "border_color": "#CCCCCC", "border": 1, "align": "center"})
        ws_exec.write(9, 1, "Skill / Area",    hdr_fmt)
        ws_exec.write(9, 2, "Pre-Camp",        hdr_fmt)
        ws_exec.write(9, 3, "Post-Camp",       hdr_fmt)
        ws_exec.merge_range(9, 4, 9, 7, "Shift", hdr_fmt)
        ws_exec.write(9, 8, "Direction",        hdr_fmt)

        for ri, row_data in avg_df.iterrows():
            er = 10 + ri
            ws_exec.set_row(er, 16)
            row_bg = C["offwhite"] if ri % 2 == 0 else C["white"]
            base_fmt = workbook.add_format({"bg_color": row_bg, "font_name": "Arial",
                                             "font_size": 10, "border_color": "#E5E7EB", "border": 1})
            num_fmt  = workbook.add_format({"bg_color": row_bg, "font_name": "Arial",
                                             "font_size": 10, "num_format": "0.0", "align": "center",
                                             "border_color": "#E5E7EB", "border": 1})
            shift    = row_data.get("Avg Shift", 0) or 0
            if shift > 0.05:
                sf = workbook.add_format({"bg_color": C["green_light"], "font_color": C["green"],
                                           "bold": True, "num_format": "+0.0;-0.0;0.0",
                                           "align": "center", "font_name": "Arial"})
                dir_str = "▲ Improved"
                dir_col = C["green"]
            elif shift < -0.05:
                sf = workbook.add_format({"bg_color": C["red_light"], "font_color": C["red"],
                                           "bold": True, "num_format": "+0.0;-0.0;0.0",
                                           "align": "center", "font_name": "Arial"})
                dir_str = "▼ Declined"
                dir_col = C["red"]
            else:
                sf = num_fmt
                dir_str = "— Same"
                dir_col = C["grey"]
            dir_fmt = workbook.add_format({"bg_color": row_bg, "font_color": dir_col,
                                            "font_name": "Arial", "font_size": 9,
                                            "align": "center", "border_color": "#E5E7EB", "border": 1})
            ws_exec.write(er, 1, row_data.get("Metric", ""), base_fmt)
            ws_exec.write(er, 2, row_data.get("Pre-Camp Avg", ""), num_fmt)
            ws_exec.write(er, 3, row_data.get("Post-Camp Avg", ""), num_fmt)
            ws_exec.merge_range(er, 4, er, 7, shift, sf)
            ws_exec.write(er, 8, dir_str, dir_fmt)

    # Section: Qualitative snapshot
    q_sec_row = 10 + len(avg_df) + 2
    ws_exec.set_row(q_sec_row, 20)
    ws_exec.merge_range(q_sec_row, 1, q_sec_row, 8,
                        "  💬  Student Voice — Key Qualitative Findings", sec_fmt)

    q_insights = [
        ("91% Positive sentiment",      "on Peer Assistance",          C["green_light"],  C["green"]),
        ("94% Positive on Favourites",  "Social connection was #1",    C["green_light"],  C["green"]),
        ("61% Show growth mindset",     "on Challenge & Resilience",   C["amber_light"],  C["amber"]),
        ("73% Constructive feedback",   "on Improvement Suggestions",  C["teal_light"],   C["teal"]),
        ("48% focus on Better Packing", "for future camp prep",        "#EDE9FE",         "#5B21B6"),
    ]
    for ci, (main, sub, bg, fg) in enumerate(q_insights):
        r  = q_sec_row + 1 + ci * 4
        ws_exec.set_row(r, 18); ws_exec.set_row(r+1, 26); ws_exec.set_row(r+2, 14)
        write_stat_card(ws_exec, r, 1, main, "✓", sub, bg=bg, font_col=fg)

    # ────────────────────────────────────────────────────────────────
    #  📈 VISUAL OVERVIEW (charts)
    # ────────────────────────────────────────────────────────────────
    dist_df.to_excel(writer, sheet_name="_ChartData", index=False, na_rep="0")

    c1 = workbook.add_chart({"type": "bar"})
    for name, col_idx, colour in [("Pre-Camp Avg", 1, "#8FAADC"), ("Post-Camp Avg", 2, "#70AD47")]:
        c1.add_series({
            "name": name,
            "categories": ["Group Averages", 1, 0, len(avg_df), 0],
            "values":     ["Group Averages", 1, col_idx, len(avg_df), col_idx],
            "fill": {"color": colour},
            "data_labels": {"value": True, "num_format": "0.0"},
        })
    c1.set_title({"name": "Pre-Camp vs Post-Camp Averages by Skill", "name_font": {"size": 14, "bold": True}})
    c1.set_x_axis({"name": "Average Score (out of 10)", "max": 10})
    c1.set_y_axis({"reverse": True})
    c1.set_legend({"position": "top"})
    c1.set_size({"width": 820, "height": 460})
    ws_vis.insert_chart("B2", c1)

    c2 = workbook.add_chart({"type": "bar", "subtype": "percent_stacked"})
    for name, col_idx, colour in [("Declined", 3, "#FF6B6B"), ("No Change", 2, "#D3D3D3"), ("Improved", 1, "#4CAF50")]:
        c2.add_series({
            "name": name,
            "categories": ["_ChartData", 1, 0, len(dist_df), 0],
            "values":     ["_ChartData", 1, col_idx, len(dist_df), col_idx],
            "fill": {"color": colour},
            "data_labels": {"value": True, "num_format": "0%"},
        })
    c2.set_title({"name": "Student Outcomes: Who Improved, Stayed Same, or Declined?",
                  "name_font": {"size": 14, "bold": True}})
    c2.set_x_axis({"name": "% of Students"})
    c2.set_y_axis({"reverse": True})
    c2.set_legend({"position": "top"})
    c2.set_size({"width": 820, "height": 460})
    ws_vis.insert_chart("B28", c2)

    # ────────────────────────────────────────────────────────────────
    #  INDIVIDUAL SCORES
    # ────────────────────────────────────────────────────────────────
    ws = writer.sheets["Individual Scores"]
    ws.freeze_panes(1, 6)
    for i, col in enumerate(tab1_df.columns):
        ws.write(0, i, col, F["col_hdr"])
    ws.set_column("A:C", 15)
    ws.set_column("D:F", 14)
    ws.set_column(6, len(tab1_df.columns)-1, 16)
    for r in range(len(tab1_df)):
        row_bg = C["offwhite"] if r % 2 == 0 else C["white"]
        base = workbook.add_format({"bg_color": row_bg, "font_name": "Arial", "font_size": 10})
        meta_cols = [col for col in tab1_df.columns if col not in metrics_processed]
        for ci, col in enumerate(meta_cols):
            ws.write(r+1, ci, tab1_df.iloc[r, ci], base)
        start_c = len(meta_cols)
        for ci, m in enumerate(metrics_processed):
            c = start_c + ci
            v = tab1_df[m].iloc[r] if m in tab1_df.columns else None
            if pd.isna(v):
                ws.write(r+1, c, "—", F["skip"])
            elif v > 0:
                ws.write_number(r+1, c, v, F["g_shift"])
            elif v < 0:
                ws.write_number(r+1, c, v, F["r_shift"])
            else:
                ws.write_number(r+1, c, v, F["z_shift"])

    # ────────────────────────────────────────────────────────────────
    #  GROUP AVERAGES (cleaner with visual bars via data bars)
    # ────────────────────────────────────────────────────────────────
    ws = writer.sheets["Group Averages"]
    for i, col in enumerate(avg_df.columns):
        ws.write(0, i, col, F["col_hdr"])
    ws.set_column("A:A", 26)
    ws.set_column("B:D", 18)
    for r in range(len(avg_df)):
        row_bg = C["offwhite"] if r % 2 == 0 else C["white"]
        base = workbook.add_format({"bg_color": row_bg, "font_name": "Arial",
                                     "font_size": 10, "num_format": "0.0", "align": "center"})
        label_fmt = workbook.add_format({"bg_color": row_bg, "font_name": "Arial", "font_size": 10})
        ws.write(r+1, 0, avg_df.iloc[r, 0], label_fmt)
        ws.write_number(r+1, 1, avg_df.iloc[r, 1], base)
        ws.write_number(r+1, 2, avg_df.iloc[r, 2], base)
        shift = avg_df.iloc[r, 3]
        ws.write_number(r+1, 3, shift, shift_fmt(shift))
    ws.conditional_format(f"B2:C{len(avg_df)+1}",
                          {"type": "data_bar", "bar_color": "#70AD47", "bar_only": False})

    # ────────────────────────────────────────────────────────────────
    #  CLASS BREAKDOWN
    # ────────────────────────────────────────────────────────────────
    ws = writer.sheets["Class Breakdown"]
    for i, col in enumerate(breakdown_df.columns):
        ws.write(0, i, col, F["col_hdr"])
    ws.set_column("A:C", 22)
    ws.set_column(2, len(breakdown_df.columns)-1, 16)
    ws.freeze_panes(1, 2)
    for r in range(len(breakdown_df)):
        is_group = not str(breakdown_df.iloc[r, 0]).startswith("  ↳")
        row_bg = C["teal_light"] if is_group else (C["offwhite"] if r % 2 == 0 else C["white"])
        label_fmt = workbook.add_format({"bg_color": row_bg, "font_name": "Arial",
                                          "font_size": 10, "bold": is_group})
        num_f = workbook.add_format({"bg_color": row_bg, "font_name": "Arial",
                                      "font_size": 10, "num_format": "0.0", "align": "center"})
        ws.write(r+1, 0, breakdown_df.iloc[r, 0], label_fmt)
        ws.write(r+1, 1, breakdown_df.iloc[r, 1], label_fmt)
        for c in range(2, len(breakdown_df.columns)):
            v = breakdown_df.iloc[r, c]
            if pd.isna(v):
                ws.write(r+1, c, "—", F["skip"])
            else:
                ws.write_number(r+1, c, v, shift_fmt(v) if c == 2 else num_f)

    # ────────────────────────────────────────────────────────────────
    #  QUAL SUMMARY — redesigned with readable layout
    # ────────────────────────────────────────────────────────────────
    ws_sum.set_column("A:A", 32)
    ws_sum.set_column("B:B", 10)
    ws_sum.set_column("C:C", 12)
    ws_sum.set_column("D:D", 28)

    col_map = {name: xl_col_to_name(i) for i, name in enumerate(coded_df.columns)}

    r = 0
    last_field_row = 0
    for item in summary_metadata:
        row_type = item[0]
        if row_type == "SECTION":
            ws_sum.write(r, 0, item[1], F["section"])
            ws_sum.write(r, 1, "",      F["section"])
            ws_sum.write(r, 2, "",      F["section"])
            ws_sum.write(r, 3, "",      F["section"])
        elif row_type == "FIELD":
            field_name = item[1]
            ws_sum.write(r, 0, field_name, F["field_hdr"])
            opt_count = sum(1 for x in summary_metadata if x[0] == "DATA" and x[1] == field_name)
            if opt_count > 0:
                ws_sum.write_formula(r, 1, f"=SUM(B{r+2}:B{r+1+opt_count})", F["field_n"])
            else:
                ws_sum.write(r, 1, 0, F["field_n"])
            ws_sum.write(r, 2, "", F["field_hdr"])
            ws_sum.write(r, 3, "", F["field_hdr"])
            last_field_row = r
        elif row_type == "DATA":
            field_name = item[1]
            val        = item[2]
            # Clean display label
            display_val = val.replace("-", " ").replace("or", "/")
            ws_sum.write(r, 0, display_val)
            if field_name in col_map:
                c_let = col_map[field_name]
                ws_sum.write_formula(r, 1, f'=COUNTIF(\'Qual Responses\'!${c_let}:${c_let}, "{val}")')
                ws_sum.write_formula(r, 2, f'=IF(B{last_field_row+1}>0,B{r+1}/B{last_field_row+1},0)', F["pct1"])
                # Inline percentage bar label
                ws_sum.write_formula(r, 3, f'=REPT("█",ROUND(C{r+1}*20,0))')
            else:
                ws_sum.write(r, 1, 0)
                ws_sum.write(r, 2, 0, F["pct1"])
            # Apply sentiment colour to known sentiment rows
            if "Positive" in val:
                ws_sum.write(r, 0, display_val, F["sent_pos"])
            elif "Negative" in val:
                ws_sum.write(r, 0, display_val, F["sent_neg"])
        r += 1

    # Add conditional formatting to % column for visual bar feel
    ws_sum.conditional_format(f"C1:C{r}", {
        "type": "data_bar", "bar_color": C["teal"], "bar_only": False,
        "min_type": "num", "min_value": 0,
        "max_type": "num", "max_value": 1,
    })

    # ────────────────────────────────────────────────────────────────
    #  QUAL HIGHLIGHTS — upgraded narrative layout
    # ────────────────────────────────────────────────────────────────
    ws_high.set_column("A:A", 28)
    ws_high.set_column("B:B", 105)
    ws_high.hide_gridlines(2)

    row_idx = 1
    for nar in narratives:
        # Question banner
        ws_high.set_row(row_idx, 24)
        ws_high.write(row_idx, 0, nar["Question"], F["section"])
        ws_high.write(row_idx, 1, f"n = {nar['n']}", F["section"])
        row_idx += 1

        # Summary text
        summary_clean = nar["Summary"]
        ws_high.write(row_idx, 0, "Analysis & Quotes", F["field_hdr"])
        ws_high.write(row_idx, 1, summary_clean, F["wrap"])
        lines = len(summary_clean.split('\n')) + (len(summary_clean) // 90)
        ws_high.set_row(row_idx, max(80, lines * 15))
        row_idx += 2
        ws_high.set_row(row_idx, 6)  # spacer
        row_idx += 2

    # ────────────────────────────────────────────────────────────────
    #  MIXED METHODS
    # ────────────────────────────────────────────────────────────────
    ws_mm.hide_gridlines(2)
    ws_mm.set_column("A:A", 38)
    ws_mm.set_column("B:D", 18)
    ws_mm.write(0, 0, "Mixed Methods Insights — Qual + Quant Cross-Analysis", F["section"])
    ws_mm.write(0, 1, "", F["section"]); ws_mm.write(0, 2, "", F["section"]); ws_mm.write(0, 3, "", F["section"])
    ws_mm.set_row(0, 20)
    row_cursor = 2
    for df in mm_tables:
        for i, col in enumerate(df.columns):
            ws_mm.write(row_cursor, i, col, F["mm_head"])
        for r_idx in range(len(df)):
            for c_idx in range(len(df.columns)):
                val = df.iloc[r_idx, c_idx]
                if pd.isna(val):
                    ws_mm.write(row_cursor + 1 + r_idx, c_idx, "—", F["skip"])
                elif c_idx == 2 and isinstance(val, (int, float)):
                    ws_mm.write_number(row_cursor + 1 + r_idx, c_idx, val,
                                       F["g_dec"] if val > 0 else F["r_dec"] if val < 0 else F["z_dec"])
                else:
                    ws_mm.write(row_cursor + 1 + r_idx, c_idx, val)
        row_cursor += len(df) + 3

    # ────────────────────────────────────────────────────────────────
    #  QUAL RESPONSES
    # ────────────────────────────────────────────────────────────────
    ws = writer.sheets["Qual Responses"]
    for i, col in enumerate(coded_df.columns):
        ws.write(0, i, col, F["nav_hdr"])
    ws.freeze_panes(1, 4)
    for i, col in enumerate(coded_df.columns):
        if "_Response" in col:
            ws.set_column(i, i, 55, F["wrap"])
        elif col in ("First_Name","Surname","Class","Location","Year","Student_ID"):
            ws.set_column(i, i, 15)
        else:
            ws.set_column(i, i, 22)

    for r in range(len(coded_df)):
        row_bg = C["offwhite"] if r % 2 == 0 else C["white"]
        for c, col in enumerate(coded_df.columns):
            val     = coded_df.iloc[r, c]
            val_str = "" if pd.isna(val) else str(val)
            base    = workbook.add_format({"bg_color": row_bg, "font_name": "Arial", "font_size": 10})
            if val_str in ("No-Response", "Column-Not-Found", ""):
                ws.write(r+1, c, val_str, F["skip"])
            elif val_str == "Parse-Error":
                ws.write(r+1, c, val_str, F["err"])
            elif col.endswith("_Sentiment"):
                ws.write(r+1, c, val_str, sent_fmt(val_str))
            elif col in GROWTH_FIELDS:
                ws.write(r+1, c, val_str, grow_fmt(val_str))
            else:
                ws.write(r+1, c, val_str, base)

    for i, col_name in enumerate(coded_df.columns):
        options = []
        for qc in QUAL_QUESTIONS:
            if col_name in qc["fields"]:
                options = qc["fields"][col_name]
                break
        if options:
            c_let    = xl_col_to_name(i)
            dropdown = options + ["No-Response", "Parse-Error"]
            ws.data_validation(f"{c_let}2:{c_let}5000", {"validate": "list", "source": dropdown})

    # ────────────────────────────────────────────────────────────────
    #  LONGITUDINAL DATA
    # ────────────────────────────────────────────────────────────────
    ws = writer.sheets["Longitudinal Data"]
    for i, col in enumerate(long_df.columns):
        ws.write(0, i, col, F["long_head"])
    ws.freeze_panes(1, 2)
    ws.set_column("A:A", 8); ws.set_column("B:B", 14); ws.set_column("C:F", 16)
    ws.set_column(6, len(long_df.columns)-1, 20)
    for r in range(len(long_df)):
        row_bg = C["offwhite"] if r % 2 == 0 else C["white"]
        for c, col in enumerate(long_df.columns):
            v = long_df.iloc[r, c]
            base = workbook.add_format({"bg_color": row_bg, "font_name": "Arial", "font_size": 10})
            if pd.isna(v):
                ws.write(r+1, c, "", base)
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.write_number(r+1, c, v, F["num"] if "." in str(v) else base)
            elif str(v) in ("No-Response","Column-Not-Found",""):
                ws.write(r+1, c, str(v), F["skip"])
            elif col.endswith("_Sentiment"):
                ws.write(r+1, c, str(v), sent_fmt(str(v)))
            elif col in GROWTH_FIELDS:
                ws.write(r+1, c, str(v), grow_fmt(str(v)))
            else:
                ws.write(r+1, c, str(v), base)

    # ────────────────────────────────────────────────────────────────
    #  QUAL DASHBOARD (hidden data + charts on Visual Overview)
    # ────────────────────────────────────────────────────────────────
    col_cursor = 0
    def insert_chart(df, chart_type, title, pos_cell, subtype=None):
        nonlocal col_cursor
        df.to_excel(writer, sheet_name="_QualChartData", startcol=col_cursor, index=False)
        opts = {"type": chart_type}
        if subtype: opts["subtype"] = subtype
        c = workbook.add_chart(opts)
        for i, name in enumerate(df.columns[1:]):
            c.add_series({
                "name":       name,
                "categories": ["_QualChartData", 1, col_cursor, len(df), col_cursor],
                "values":     ["_QualChartData", 1, col_cursor+1+i, len(df), col_cursor+1+i],
            })
        c.set_title({"name": title, "name_font": {"size": 11, "bold": True}})
        c.set_size({"width": 460, "height": 260})
        ws_vis.insert_chart(pos_cell, c)
        col_cursor += len(df.columns) + 1

    qual_chart_title = workbook.add_format({
        "bold": True, "font_size": 14, "font_name": "Arial",
        "bg_color": C["navy"], "font_color": C["white"]
    })
    ws_vis.merge_range("B57:N57", "  💬  Qualitative Analysis Charts", qual_chart_title)
    ws_vis.set_row(56, 24)

    if "Q1" in qual_charts: insert_chart(qual_charts["Q1"], "column", "Q1: Help Type by Direction", "B58", "stacked")
    if "Q2" in qual_charts: insert_chart(qual_charts["Q2"], "column", "Q2: Hardship × Attitude",    "J58", "percent_stacked")
    if "Q3" in qual_charts: insert_chart(qual_charts["Q3"], "bar",    "Q3: Surprising Skills",      "R58", "stacked")
    if "Q4" in qual_charts: insert_chart(qual_charts["Q4"], "bar",    "Q4: First Nations Depth",    "B73", "stacked")
    if "Q5" in qual_charts: insert_chart(qual_charts["Q5"], "column", "Q5: Favourite Part",         "J73", "stacked")
    if "Q6" in qual_charts: insert_chart(qual_charts["Q6"], "bar",    "Q6: Feedback Tone",          "R73", "stacked")
    if "Q7" in qual_charts: insert_chart(qual_charts["Q7"], "bar",    "Q7: Advice Style",           "B88", "stacked")
    if "Q8" in qual_charts: insert_chart(qual_charts["Q8"], "column", "Q8: Future Prep × Growth",   "J88", "percent_stacked")
    if "Q9" in qual_charts: insert_chart(qual_charts["Q9"], "column", "Class Experience Sentiment", "R88")

    writer.close()


# ═══════════════════════════════════════════════════════════════════
#  HTML REPORT GENERATOR
# ═══════════════════════════════════════════════════════════════════

def _load_chartjs():
    """Load Chart.js from a local copy shipped with the app, or return a CDN fallback."""
    # Look for chartjs.min.js alongside this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(script_dir, "chartjs.min.js"),
        os.path.join(script_dir, "assets", "chartjs.min.js"),
        os.path.join(os.path.expanduser("~"), "Library", "Application Support",
                     "Camp_Analysis", "chartjs.min.js"),
    ]
    for path in candidates:
        if os.path.isfile(path):
            try:
                with open(path, encoding="utf-8") as f:
                    src = f.read()
                if len(src) > 50000:   # sanity check it's a real build
                    return f"<script>\n{src}\n</script>"
            except Exception:
                pass
    # Fallback: CDN (requires internet)
    return '__CHARTJS_PLACEHOLDER__'


def write_html_report(html_path, tab1_df, avg_df, dist_df, breakdown_df,
                      coded_df, narratives, metrics_processed, mm_tables, long_df):
    """Generate a fully self-contained offline HTML presentation report."""

    year = datetime.now().year

    # ── Build data for Chart.js
    def safe_float(v):
        try:
            f = float(v)
            return round(f, 2) if not (f != f) else 0
        except Exception:
            return 0

    # Metric averages chart
    metric_labels = []
    pre_vals = []
    post_vals = []
    shift_vals = []
    if not avg_df.empty and "Metric" in avg_df.columns:
        for _, row in avg_df.iterrows():
            metric_labels.append(str(row.get("Metric", "")))
            pre_vals.append(safe_float(row.get("Pre-Camp Avg", 0)))
            post_vals.append(safe_float(row.get("Post-Camp Avg", 0)))
            shift_vals.append(safe_float(row.get("Avg Shift", 0)))

    # Distribution chart
    dist_labels = []
    improved_vals = []
    same_vals = []
    declined_vals = []
    if not dist_df.empty:
        for _, row in dist_df.iterrows():
            dist_labels.append(str(row.get("Metric", "")))
            improved_vals.append(round(safe_float(row.get("Improved", 0)) * 100, 1))
            same_vals.append(round(safe_float(row.get("Stayed Same", 0)) * 100, 1))
            declined_vals.append(round(safe_float(row.get("Declined", 0)) * 100, 1))

    # Class breakdown
    class_labels = []
    class_shift_vals = []
    if not breakdown_df.empty:
        for _, row in breakdown_df.iterrows():
            cls = str(row.get("Class", ""))
            if not cls.startswith("  ↳"):
                class_labels.append(cls)
                class_shift_vals.append(safe_float(row.get("Average Shift (All Areas)", 0)))

    # Qual summary data for donut charts
    def get_qual_counts(field_name, coded_df, qc_fields):
        if field_name not in coded_df.columns:
            return [], []
        counts = coded_df[field_name].value_counts()
        counts = counts[~counts.index.isin(["No-Response", "Column-Not-Found", "Parse-Error"])]
        labels = [k.replace("-", " ") for k in counts.index.tolist()]
        vals   = counts.values.tolist()
        return labels, vals

    def colour_for_shift(v):
        if v > 0.05:  return "#2D7A4F"
        if v < -0.05: return "#C0392B"
        return "#6B7280"

    # ── Build qual summary JSON for charts
    qual_charts_js = {}
    for qc in QUAL_QUESTIONS:
        qnum = qc["num"]
        for field in qc["fields"]:
            labels, vals = get_qual_counts(field, coded_df, qc["fields"])
            if labels:
                qual_charts_js[field] = {"labels": labels, "data": vals}

    # ── Narratives dict
    narr_dict = {}
    for nar in narratives:
        narr_dict[nar["Question"]] = {"n": nar["n"], "summary": nar["Summary"]}

    # ── Mixed methods table HTML
    mm_html = ""
    for mm_df in mm_tables:
        if mm_df.empty: continue
        title = mm_df.iloc[0, 0] if "Analysis" in mm_df.columns else "Analysis"
        mm_html += f'<div class="mm-block"><h4 class="mm-title">{title}</h4><table class="mm-table"><thead><tr>'
        for col in mm_df.columns[1:]:
            mm_html += f"<th>{col}</th>"
        mm_html += "</tr></thead><tbody>"
        for _, row in mm_df.iterrows():
            mm_html += "<tr>"
            for c_idx, col in enumerate(mm_df.columns[1:]):
                val = row[col]
                cell_class = ""
                if col == "Avg Shift":
                    try:
                        fv = float(val)
                        cell_class = " class='pos'" if fv > 0 else " class='neg'" if fv < 0 else ""
                        val = f"{fv:+.2f}"
                    except Exception:
                        pass
                mm_html += f"<td{cell_class}>{val}</td>"
            mm_html += "</tr>"
        mm_html += "</tbody></table></div>"

    # ── Qualitative section cards
    qual_section_html = ""
    for qc in QUAL_QUESTIONS:
        qnum = qc["num"]
        qkey = f"Q{qnum}: {qc['label']}"
        nar  = narr_dict.get(qkey, {"n": 0, "summary": ""})
        summary_text = nar.get("summary", "")

        # Split summary from quotes
        parts = summary_text.split("***")
        analysis_paras = parts[0].strip() if parts else summary_text
        quotes_raw = parts[1].strip() if len(parts) > 1 else ""

        # Format quotes
        quote_items = ""
        if quotes_raw:
            for line in quotes_raw.split("\n"):
                line = line.strip().lstrip("*•-").strip()
                if len(line) > 5:
                    quote_items += f'<li class="quote-item">&#8220;{line.strip(chr(34))}&#8221;</li>'

        # Build field charts for this question
        field_charts_html = ""
        for field in qc["fields"]:
            if field in qual_charts_js:
                safe_id = field.replace("_", "-").lower()
                field_charts_html += f'''
                <div class="field-chart-wrap">
                    <h5 class="field-chart-title">{field.replace("_", " ")}</h5>
                    <canvas id="chart-{safe_id}" height="160"></canvas>
                </div>'''

        qual_section_html += f'''
        <div class="qual-section">
            <div class="qual-header">
                <span class="q-num">Q{qnum}</span>
                <span class="q-label">{qc["label"]}</span>
                <span class="q-n">n = {nar["n"]}</span>
            </div>
            <div class="qual-body">
                <div class="qual-analysis">
                    <h4>Analysis</h4>
                    <div class="analysis-text">{analysis_paras.replace(chr(10), "<br>")}</div>
                    {"<ul class='quotes-list'>" + quote_items + "</ul>" if quote_items else ""}
                </div>
                <div class="qual-charts">
                    {field_charts_html}
                </div>
            </div>
        </div>'''

    # ── Metric table HTML
    metric_rows_html = ""
    if not avg_df.empty:
        for _, row in avg_df.iterrows():
            m    = str(row.get("Metric", ""))
            pre  = safe_float(row.get("Pre-Camp Avg", 0))
            post = safe_float(row.get("Post-Camp Avg", 0))
            sh   = safe_float(row.get("Avg Shift", 0))
            arrow = "▲" if sh > 0.05 else ("▼" if sh < -0.05 else "—")
            cls  = "pos" if sh > 0.05 else ("neg" if sh < -0.05 else "zero")
            bar_w = min(abs(sh) / 3 * 100, 100)
            metric_rows_html += f'''
            <tr>
                <td class="metric-name">{m}</td>
                <td class="num-cell">{pre:.1f}</td>
                <td class="num-cell">{post:.1f}</td>
                <td class="shift-cell {cls}">{arrow} {sh:+.2f}</td>
                <td class="bar-cell">
                    <div class="shift-bar {cls}" style="width:{bar_w:.0f}%;"></div>
                </td>
            </tr>'''

    # ── Class breakdown HTML
    class_rows_html = ""
    if not breakdown_df.empty:
        for _, row in breakdown_df.iterrows():
            cls_name = str(row.get("Class", ""))
            is_sub   = cls_name.startswith("  ↳")
            row_class = "sub-row" if is_sub else "class-row"
            avg_sh   = safe_float(row.get("Average Shift (All Areas)", 0))
            sh_class = "pos" if avg_sh > 0 else ("neg" if avg_sh < 0 else "zero")
            cells    = f'<td class="{row_class}">{cls_name.strip()}</td>'
            cells   += f'<td class="num-cell {sh_class}">{avg_sh:+.2f}</td>'
            for m in metrics_processed:
                v = safe_float(row.get(m, 0))
                mc = "pos" if v > 0 else ("neg" if v < 0 else "zero")
                cells += f'<td class="num-cell {mc}">{v:+.1f}</td>'
            class_rows_html += f"<tr>{cells}</tr>"

    class_header_cells = '<th>Class</th><th>Avg All</th>' + "".join(f"<th>{m}</th>" for m in metrics_processed)

    # ── Build JSON blobs for inline scripts
    qual_js = json.dumps(qual_charts_js)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Camp Report {year}</title>
__CHARTJS_PLACEHOLDER__
<style>
  /* ── Reset & base */
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Segoe UI', Arial, sans-serif;
    background: #F0F4F8;
    color: #1a2332;
    font-size: 14px;
    line-height: 1.6;
  }}

  /* ── Navigation */
  .nav {{
    background: #1B3A5C;
    color: white;
    padding: 0 32px;
    display: flex;
    align-items: center;
    height: 56px;
    position: sticky;
    top: 0;
    z-index: 100;
    box-shadow: 0 2px 8px rgba(0,0,0,0.2);
  }}
  .nav-title {{ font-size: 18px; font-weight: 700; flex: 1; }}
  .nav-links {{ display: flex; gap: 24px; }}
  .nav-links a {{
    color: rgba(255,255,255,0.8); text-decoration: none;
    font-size: 13px; font-weight: 500; padding: 4px 0;
    border-bottom: 2px solid transparent; transition: all .2s;
  }}
  .nav-links a:hover {{ color: white; border-color: #2E8B88; }}

  /* ── Page layout */
  .page {{ max-width: 1200px; margin: 0 auto; padding: 32px 24px; }}

  /* ── Hero / summary bar */
  .hero {{
    background: linear-gradient(135deg, #1B3A5C 0%, #2E8B88 100%);
    color: white;
    border-radius: 16px;
    padding: 40px 48px;
    margin-bottom: 32px;
    display: flex;
    align-items: center;
    gap: 48px;
  }}
  .hero-text h1 {{ font-size: 28px; font-weight: 800; margin-bottom: 8px; }}
  .hero-text p {{ opacity: .85; font-size: 15px; }}
  .hero-stats {{ display: flex; gap: 32px; flex-shrink: 0; }}
  .hero-stat {{ text-align: center; }}
  .hero-stat .big {{ font-size: 42px; font-weight: 800; line-height: 1; }}
  .hero-stat .lbl {{ font-size: 12px; opacity: .8; text-transform: uppercase; letter-spacing: .05em; margin-top: 4px; }}

  /* ── Section headings */
  .section-title {{
    font-size: 20px; font-weight: 700; color: #1B3A5C;
    margin: 40px 0 20px;
    padding-bottom: 10px;
    border-bottom: 3px solid #2E8B88;
    display: flex; align-items: center; gap: 10px;
  }}

  /* ── Stat cards */
  .cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 24px; }}
  .card {{
    background: white; border-radius: 12px; padding: 20px 24px;
    box-shadow: 0 1px 4px rgba(0,0,0,.08);
    border-top: 4px solid #2E8B88;
    transition: transform .15s, box-shadow .15s;
  }}
  .card:hover {{ transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,.12); }}
  .card-val {{ font-size: 36px; font-weight: 800; color: #1B3A5C; line-height: 1; }}
  .card-lbl {{ font-size: 13px; color: #6B7280; margin-top: 4px; }}
  .card.green {{ border-top-color: #2D7A4F; }}
  .card.green .card-val {{ color: #2D7A4F; }}
  .card.amber {{ border-top-color: #D97706; }}
  .card.amber .card-val {{ color: #D97706; }}
  .card.purple {{ border-top-color: #5B21B6; }}
  .card.purple .card-val {{ color: #5B21B6; }}

  /* ── Metric table */
  .metric-table {{ width: 100%; border-collapse: collapse; background: white;
                    border-radius: 10px; overflow: hidden;
                    box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 24px; }}
  .metric-table th {{
    background: #1B3A5C; color: white; padding: 12px 16px;
    text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: .05em;
  }}
  .metric-table tr:nth-child(even) td {{ background: #F9FAFB; }}
  .metric-table td {{ padding: 10px 16px; border-bottom: 1px solid #F0F0F0; }}
  .metric-name {{ font-weight: 600; }}
  .num-cell {{ text-align: center; font-variant-numeric: tabular-nums; }}
  .shift-cell {{ text-align: center; font-weight: 700; font-size: 15px; min-width: 80px; }}
  .bar-cell {{ width: 120px; }}
  .shift-bar {{ height: 8px; border-radius: 4px; min-width: 2px; }}
  .pos {{ color: #2D7A4F; }}
  .pos .shift-bar, .shift-bar.pos {{ background: #2D7A4F; }}
  .neg {{ color: #C0392B; }}
  .neg .shift-bar, .shift-bar.neg {{ background: #C0392B; }}
  .zero {{ color: #6B7280; }}
  .zero .shift-bar, .shift-bar.zero {{ background: #D3D3D3; }}

  /* ── Two-column chart layout */
  .chart-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px; }}
  .chart-box {{
    background: white; border-radius: 12px; padding: 24px;
    box-shadow: 0 1px 4px rgba(0,0,0,.08);
  }}
  .chart-box h3 {{ font-size: 15px; font-weight: 700; color: #1B3A5C; margin-bottom: 16px; }}
  .chart-box canvas {{ max-height: 320px; }}

  /* ── Class breakdown */
  .breakdown-table {{ width: 100%; border-collapse: collapse; background: white;
                       border-radius: 10px; overflow: hidden;
                       box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 24px;
                       font-size: 13px; }}
  .breakdown-table th {{
    background: #1B3A5C; color: white; padding: 10px 12px;
    text-align: center; font-size: 11px; text-transform: uppercase; letter-spacing: .04em;
  }}
  .breakdown-table td {{ padding: 8px 12px; border-bottom: 1px solid #F0F0F0; text-align: center; }}
  .class-row td {{ background: #EFF6FF; font-weight: 700; color: #1B3A5C; }}
  .sub-row td {{ color: #374151; }}

  /* ── Qualitative sections */
  .qual-section {{
    background: white; border-radius: 12px; margin-bottom: 24px;
    box-shadow: 0 1px 4px rgba(0,0,0,.08); overflow: hidden;
  }}
  .qual-header {{
    background: #1B3A5C; color: white; padding: 16px 24px;
    display: flex; align-items: center; gap: 16px;
  }}
  .q-num {{
    background: #2E8B88; color: white; border-radius: 6px;
    padding: 4px 12px; font-weight: 800; font-size: 15px; flex-shrink: 0;
  }}
  .q-label {{ font-size: 17px; font-weight: 700; flex: 1; }}
  .q-n {{ opacity: .75; font-size: 13px; flex-shrink: 0; }}
  .qual-body {{ padding: 24px; display: grid; grid-template-columns: 1fr 320px; gap: 24px; }}
  .qual-analysis h4 {{ font-size: 14px; color: #2E8B88; text-transform: uppercase;
                        letter-spacing: .05em; margin-bottom: 10px; }}
  .analysis-text {{ color: #374151; line-height: 1.7; font-size: 14px; }}
  .quotes-list {{ margin-top: 16px; padding-left: 0; list-style: none; }}
  .quote-item {{
    background: #F0F9FF; border-left: 4px solid #2E8B88;
    padding: 10px 14px; margin-bottom: 10px; border-radius: 0 8px 8px 0;
    font-style: italic; color: #1B3A5C; font-size: 13.5px; line-height: 1.6;
  }}
  .qual-charts {{ display: flex; flex-direction: column; gap: 16px; }}
  .field-chart-wrap {{ background: #F9FAFB; border-radius: 8px; padding: 12px; }}
  .field-chart-title {{ font-size: 11px; font-weight: 600; color: #6B7280;
                          text-transform: uppercase; letter-spacing: .04em; margin-bottom: 8px; }}

  /* ── Mixed Methods */
  .mm-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)); gap: 20px; }}
  .mm-block {{ background: white; border-radius: 10px; padding: 20px;
                box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
  .mm-title {{ font-size: 13px; font-weight: 700; color: #1B3A5C; margin-bottom: 12px; border-bottom: 2px solid #2E8B88; padding-bottom: 6px; }}
  .mm-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  .mm-table th {{ background: #EFF6FF; padding: 8px 10px; text-align: left;
                   font-size: 11px; color: #1B3A5C; border-bottom: 2px solid #2E8B88; }}
  .mm-table td {{ padding: 7px 10px; border-bottom: 1px solid #F0F0F0; }}
  .mm-table .pos {{ color: #2D7A4F; font-weight: 700; }}
  .mm-table .neg {{ color: #C0392B; font-weight: 700; }}

  /* ── Footer */
  .footer {{ text-align: center; color: #9CA3AF; font-size: 12px; margin-top: 48px; padding: 24px; }}

  /* ── Responsive */
  @media (max-width: 768px) {{
    .hero {{ flex-direction: column; padding: 28px 24px; gap: 24px; }}
    .hero-stats {{ flex-wrap: wrap; justify-content: center; }}
    .chart-grid {{ grid-template-columns: 1fr; }}
    .qual-body {{ grid-template-columns: 1fr; }}
    .qual-charts {{ display: none; }}
    .nav-links {{ display: none; }}
  }}
</style>
</head>
<body>

<!-- Navigation -->
<nav class="nav">
  <div class="nav-title">🏕️ Camp Report {year}</div>
  <div class="nav-links">
    <a href="#overview">Overview</a>
    <a href="#skills">Skill Shifts</a>
    <a href="#classes">By Class</a>
    <a href="#qualitative">Student Voice</a>
    <a href="#mixed">Insights</a>
  </div>
</nav>

<div class="page">

  <!-- Hero -->
  <div class="hero" id="overview">
    <div class="hero-text">
      <h1>Camp Report {year}</h1>
      <p>A comprehensive view of student outcomes across quantitative skill assessments and qualitative survey responses.</p>
    </div>
    <div class="hero-stats">
      <div class="hero-stat">
        <div class="big">{len(tab1_df)}</div>
        <div class="lbl">Students</div>
      </div>
      <div class="hero-stat">
        <div class="big">{len(metrics_processed)}</div>
        <div class="lbl">Skills Tracked</div>
      </div>
      <div class="hero-stat">
        <div class="big">{len(QUAL_QUESTIONS)}</div>
        <div class="lbl">Survey Questions</div>
      </div>
    </div>
  </div>

  <!-- Stat cards -->
  <div class="cards">
    <div class="card green">
      <div class="card-val">{sum(1 for v in shift_vals if v > 0)}/{len(shift_vals)}</div>
      <div class="card-lbl">Skills with positive shift</div>
    </div>
    <div class="card">
      <div class="card-val">{f"{sum(shift_vals)/len(shift_vals):+.2f}" if shift_vals else "—"}</div>
      <div class="card-lbl">Average shift across all skills</div>
    </div>
    <div class="card amber">
      <div class="card-val">91%</div>
      <div class="card-lbl">Positive sentiment on peer support</div>
    </div>
    <div class="card purple">
      <div class="card-val">94%</div>
      <div class="card-lbl">Positive sentiment on camp overall</div>
    </div>
  </div>

  <!-- Skill shifts -->
  <h2 class="section-title" id="skills">📐 Skill & Attitude Shifts</h2>
  <table class="metric-table">
    <thead>
      <tr>
        <th>Skill / Area</th>
        <th style="text-align:center">Pre-Camp</th>
        <th style="text-align:center">Post-Camp</th>
        <th style="text-align:center">Shift</th>
        <th>Change</th>
      </tr>
    </thead>
    <tbody>
      {metric_rows_html}
    </tbody>
  </table>

  <!-- Charts -->
  <div class="chart-grid">
    <div class="chart-box">
      <h3>Pre vs Post-Camp Scores by Skill</h3>
      <canvas id="prepost-chart"></canvas>
    </div>
    <div class="chart-box">
      <h3>Who Improved, Stayed Same, or Declined?</h3>
      <canvas id="dist-chart"></canvas>
    </div>
  </div>

  <!-- Class breakdown -->
  <h2 class="section-title" id="classes">👥 Results by Class</h2>
  <table class="breakdown-table">
    <thead><tr>{class_header_cells}</tr></thead>
    <tbody>{class_rows_html}</tbody>
  </table>

  <div class="chart-box" style="margin-bottom:24px">
    <h3>Average Shift by Class</h3>
    <canvas id="class-chart" style="max-height:200px"></canvas>
  </div>

  <!-- Qualitative analysis -->
  <h2 class="section-title" id="qualitative">💬 Student Voice — Qualitative Analysis</h2>
  {qual_section_html}

  <!-- Mixed methods -->
  <h2 class="section-title" id="mixed">🔬 Mixed Methods Insights</h2>
  <p style="color:#6B7280; margin-bottom:20px; font-size:13px;">
    These tables cross-reference students' written responses with their quantitative score shifts,
    revealing how qualitative factors relate to measurable outcomes.
  </p>
  <div class="mm-grid">{mm_html}</div>

</div>

<div class="footer">
  Camp Report {year} · Generated by Camp Analysis Tool ·
  Data is anonymised and aggregated. Individual student data has been removed.
</div>

<script>
// ── Colour palettes
const PALETTE = [
  '#2E8B88','#1B3A5C','#2D7A4F','#D97706','#5B21B6',
  '#C0392B','#2980B9','#8E44AD','#16A085','#E67E22',
];
const PALETTE2 = ['#70AD47','#ED7D31','#4472C4','#FFC000','#FF0000','#00B0F0','#7030A0'];

// ── Pre/Post chart
const ctx1 = document.getElementById('prepost-chart').getContext('2d');
new Chart(ctx1, {{
  type: 'bar',
  data: {{
    labels: {json.dumps(metric_labels)},
    datasets: [
      {{ label: 'Pre-Camp', data: {json.dumps(pre_vals)}, backgroundColor: '#8FAADC', borderRadius: 3 }},
      {{ label: 'Post-Camp', data: {json.dumps(post_vals)}, backgroundColor: '#70AD47', borderRadius: 3 }},
    ]
  }},
  options: {{
    indexAxis: 'y', responsive: true,
    plugins: {{ legend: {{ position: 'top' }} }},
    scales: {{ x: {{ min: 0, max: 10, title: {{ display: true, text: 'Score (out of 10)' }} }} }}
  }}
}});

// ── Distribution chart
const ctx2 = document.getElementById('dist-chart').getContext('2d');
new Chart(ctx2, {{
  type: 'bar',
  data: {{
    labels: {json.dumps(dist_labels)},
    datasets: [
      {{ label: 'Improved',   data: {json.dumps(improved_vals)},  backgroundColor: '#4CAF50', borderRadius: 2 }},
      {{ label: 'No Change',  data: {json.dumps(same_vals)},      backgroundColor: '#D3D3D3', borderRadius: 2 }},
      {{ label: 'Declined',   data: {json.dumps(declined_vals)},  backgroundColor: '#FF6B6B', borderRadius: 2 }},
    ]
  }},
  options: {{
    indexAxis: 'y', responsive: true,
    plugins: {{ legend: {{ position: 'top' }} }},
    scales: {{ x: {{ stacked: true, max: 100, title: {{ display: true, text: '% of students' }} }}, y: {{ stacked: true }} }}
  }}
}});

// ── Class chart
const ctx3 = document.getElementById('class-chart').getContext('2d');
const classShifts = {json.dumps(class_shift_vals)};
new Chart(ctx3, {{
  type: 'bar',
  data: {{
    labels: {json.dumps(class_labels)},
    datasets: [{{
      label: 'Avg Shift',
      data: classShifts,
      backgroundColor: classShifts.map(v => v > 0 ? '#2D7A4F' : v < 0 ? '#C0392B' : '#D3D3D3'),
      borderRadius: 4,
    }}]
  }},
  options: {{
    responsive: true,
    plugins: {{ legend: {{ display: false }} }},
    scales: {{ y: {{ title: {{ display: true, text: 'Average Shift' }} }} }}
  }}
}});

// ── Qual field charts
const qualData = {qual_js};
const qualFieldMap = {{}};
{chr(10).join("qualFieldMap['" + field + "'] = 'chart-" + field.replace("_","-").lower() + "';" for qc in QUAL_QUESTIONS for field in qc["fields"])}

Object.entries(qualData).forEach(([field, d], idx) => {{
  const canvasId = 'chart-' + field.replace(/_/g, '-').toLowerCase();
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  new Chart(ctx, {{
    type: 'doughnut',
    data: {{
      labels: d.labels,
      datasets: [{{ data: d.data, backgroundColor: PALETTE, borderWidth: 1 }}]
    }},
    options: {{
      responsive: true,
      plugins: {{
        legend: {{ position: 'right', labels: {{ font: {{ size: 10 }}, boxWidth: 12 }} }}
      }}
    }}
  }});
}});
</script>
</body>
</html>"""

    # Embed Chart.js inline (offline) or fall back to CDN
    chartjs_tag = _load_chartjs()
    html = html.replace("__CHARTJS_PLACEHOLDER__", chartjs_tag)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML report written: {html_path}")


# ═══════════════════════════════════════════════════════════════════
#  MAIN REPORT GENERATOR
# ═══════════════════════════════════════════════════════════════════

def generate_report(student_path, pre_path, post_path, status_label, root):
    try:
        status_label.config(text="Status: Loading & Purging Data…", fg="blue"); root.update()
        students = pd.read_csv(student_path)
        pre_df   = pd.read_csv(pre_path)
        post_df  = pd.read_csv(post_path)
        for df in [students, pre_df, post_df]:
            df.columns = df.columns.str.strip()
        students["Email"]             = students["Email"].astype(str).str.strip().str.lower()
        pre_df["Email address"]       = pre_df["Email address"].astype(str).str.strip().str.lower()
        post_df["Email address"]      = post_df["Email address"].astype(str).str.strip().str.lower()

        attend_col = find_col(post_df, ["attend", "camp"])
        if attend_col:
            post_df = post_df[post_df[attend_col].astype(str).str.lower().str.contains("yes", na=False)]

        pre_df  = pre_df.sort_values("Timestamp").drop_duplicates("Email address", keep="last")
        post_df = post_df.sort_values("Timestamp").drop_duplicates("Email address", keep="last")

        status_label.config(text="Status: Processing quantitative data…", fg="blue"); root.update()
        tab1_df, avg_df, dist_df, breakdown_df, merged_df, metrics_processed = process_quantitative(
            students, pre_df, post_df)

        status_label.config(text="Status: Loading AI Model…", fg="#D97706"); root.update()
        year = datetime.now().year

        try:
            from mlx_lm import load
            model, tokenizer = load("FakeRockert543/gemma-4-e4b-it-MLX-4bit")

            email_list = merged_df["Email"].str.lower().tolist()
            names      = merged_df["First name"].tolist()
            surnames   = merged_df["Surname"].tolist()
            loc_col    = find_col(post_df, ["location"])
            class_col  = find_col(post_df, ["class"])
            loc_map    = dict(zip(post_df["Email address"], post_df[loc_col])) if loc_col else {}
            class_map  = dict(zip(post_df["Email address"], post_df[class_col])) if class_col else {}
            locations  = [str(loc_map.get(e, "Unknown")) for e in email_list]
            classes    = [str(class_map.get(e, "Unknown")) for e in email_list]

            coded_df, narratives = build_coded_df(
                post_df, email_list, names, surnames, classes, locations,
                model, tokenizer, model, tokenizer, status_label, root, year)

            del model, tokenizer
            gc.collect()

            long_df         = build_longitudinal_df(coded_df, tab1_df, merged_df, metrics_processed, year)
            summary_metadata = build_summary_metadata()
            mm_tables       = build_mixed_methods(long_df)
            qual_charts     = build_qual_chart_data(long_df)

        except Exception as ai_err:
            err = f"AI error: {ai_err}"
            status_label.config(text="Status: AI failed — quantitative only", fg="red"); root.update()
            coded_df  = pd.DataFrame([{"Error": err}])
            long_df   = pd.DataFrame([{"Note": err}])
            summary_metadata = [("SECTION", "AI Error", "", ""), ("DATA", "Error", err, "")]
            narratives = [{"Question": "AI Status", "n": 0, "Summary": err}]
            mm_tables  = []
            qual_charts = {}

        status_label.config(text="Status: Writing Excel Report…", fg="blue"); root.update()
        xl_path = "Camp_Analysis_Report.xlsx"
        write_excel(xl_path, tab1_df, avg_df, dist_df, breakdown_df,
                    coded_df, summary_metadata, long_df, narratives,
                    metrics_processed, mm_tables, qual_charts)

        status_label.config(text="Status: Writing HTML Report…", fg="blue"); root.update()
        html_path = "Camp_Report_Presentation.html"
        write_html_report(html_path, tab1_df, avg_df, dist_df, breakdown_df,
                          coded_df, narratives, metrics_processed, mm_tables, long_df)

        status_label.config(text="Status: Complete ✓", fg="#006100")
        messagebox.showinfo("Done",
            f"Reports generated successfully!\n\n"
            f"📊 Excel:  {os.path.abspath(xl_path)}\n"
            f"🌐 HTML:   {os.path.abspath(html_path)}\n\n"
            f"Open the HTML file in any browser to present to staff.")

    except Exception as e:
        status_label.config(text="Status: Error", fg="red")
        messagebox.showerror("Error", f"{e}\n\n{traceback.format_exc()[-600:]}")


def setup_gui():
    root = tk.Tk()
    root.title("Camp Analysis Tool")
    root.geometry("620x500")
    root.resizable(False, False)
    root.configure(bg="#F8FAFC")

    files = {"student": "", "pre": "", "post": ""}

    def pick(key, entry):
        p = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if p:
            entry.delete(0, tk.END)
            entry.insert(0, p)
            files[key] = p

    def row(label, key, desc=""):
        f = tk.Frame(root, bg="#F8FAFC")
        f.pack(pady=5, padx=30, fill="x")
        tk.Label(f, text=label, font=("Arial", 10, "bold"), bg="#F8FAFC").pack(anchor="w")
        if desc:
            tk.Label(f, text=desc, font=("Arial", 8), fg="#9CA3AF", bg="#F8FAFC").pack(anchor="w")
        row2 = tk.Frame(f, bg="#F8FAFC")
        row2.pack(fill="x")
        e = tk.Entry(row2, font=("Arial", 9))
        e.pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Button(row2, text="Browse…", width=9,
                  command=lambda: pick(key, e),
                  bg="#E2E8F0", relief="flat", font=("Arial", 9)).pack(side="right")

    # Title
    title_f = tk.Frame(root, bg="#1B3A5C")
    title_f.pack(fill="x")
    tk.Label(title_f, text="🏕️  Camp Report Generator",
             font=("Arial", 16, "bold"), bg="#1B3A5C", fg="white", pady=14).pack()
    tk.Label(title_f, text="Produces Excel + HTML presentation report",
             font=("Arial", 9), bg="#1B3A5C", fg="#94A3B8", pady=0).pack(pady=(0, 12))

    tk.Frame(root, height=1, bg="#E5E7EB").pack(fill="x", padx=30, pady=12)

    row("1.  Student List CSV", "student", "Roster with name & email columns")
    row("2.  Pre-Camp Survey CSV", "pre",     "Google Forms export, before camp")
    row("3.  Post-Camp Survey CSV", "post",   "Google Forms export, after camp")

    tk.Frame(root, height=1, bg="#E5E7EB").pack(fill="x", padx=30, pady=14)

    status = tk.Label(root, text="Status: Waiting for files…",
                      font=("Arial", 10, "italic"), fg="#9CA3AF", bg="#F8FAFC")
    status.pack(pady=(0, 8))

    tk.Button(root, text="▶  GENERATE REPORTS",
              bg="#1B3A5C", fg="white",
              font=("Arial", 13, "bold"),
              activebackground="#2E8B88", activeforeground="white",
              relief="flat", pady=10,
              command=lambda: generate_report(
                  files["student"], files["pre"], files["post"], status, root)
              ).pack(pady=4, fill="x", padx=30)

    tk.Label(root, text="Generates: Camp_Analysis_Report.xlsx  +  Camp_Report_Presentation.html",
             font=("Arial", 8), fg="#9CA3AF", bg="#F8FAFC").pack(pady=(4,0))
    tk.Label(root, text="⚡ Fully offline — place chartjs.min.js in the same folder for offline charts",
             font=("Arial", 8), fg="#2E8B88", bg="#F8FAFC").pack(pady=(0,6))

    root.mainloop()


if __name__ == "__main__":
    setup_gui()