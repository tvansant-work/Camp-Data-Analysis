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

import gc, hashlib, json, math, os, pickle, re, traceback
from datetime import datetime

import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
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


_GARBAGE_SET = frozenset({
    "yes", "no", "nope", "yep", "yup", "yeah", "na", "n/a", "none",
    "nothing", "idk", "not sure", "don't know", "dont know", "good",
    "fine", "ok", "okay", "not applicable", "nil", "null", "-", ".",
    "no response", "no comment", "unsure", "maybe", "dunno", "cool",
    "fun", "great", "nice", "good thanks", "not really", "no idea",
})

def is_valid_response(text):
    """Return True if the response is a genuine, substantive answer worth analysing."""
    text = str(text).strip()
    # Minimum length
    if len(text) < 10:
        return False
    # Must have at least 3 words containing letters
    words = [w for w in text.split() if re.search(r'[a-zA-Z]', w)]
    if len(words) < 3:
        return False
    # Blocklist of non-answers
    if text.lower() in _GARBAGE_SET:
        return False
    # Reject if fewer than 40% alphabetic characters (random chars, numbers-only, etc.)
    alpha = sum(1 for c in text if c.isalpha())
    if alpha / len(text) < 0.4:
        return False
    return True


# ═══════════════════════════════════════════════════════════════════
#  CACHE
# ═══════════════════════════════════════════════════════════════════

CACHE_DIR     = os.path.expanduser("~/Library/Application Support/Camp_Analysis")
CACHE_FILE    = os.path.join(CACHE_DIR, "coding_cache.pkl")
CACHE_VERSION = "v8-map-reduce"

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
#  NON-ATTENDER ANALYSIS CODEBOOK
# ═══════════════════════════════════════════════════════════════════

NON_ATTEND_CODEBOOK = [
    {
        "num": 1,
        "label": "Reason for Not Attending",
        "fields": {
            "NA_Reason": ["Health-or-Medical", "Family-Commitment", "Financial-Cost",
                          "Fear-or-Anxiety", "Another-School-Event", "Personal-Choice",
                          "Logistical", "Not-Specified"],
            "NA_Regret": ["Yes", "Partial", "No", "Unclear"],
            "NA_Sentiment": ["Regretful", "Neutral", "Relieved", "Unclear"],
        },
        "guide": (
            "Code the primary reason the student did not attend camp. "
            "NA_Regret: whether they express regret about missing it. "
            "NA_Sentiment: their overall emotional tone toward not attending."
        ),
    },
    {
        "num": 2,
        "label": "Constructive Feedback for Future Attendance",
        "fields": {
            "NA_Has_Constructive": ["Yes", "Partial", "No"],
            "NA_Support_Type": ["Financial-Aid", "Medical-or-Dietary", "More-Notice-or-Info",
                                "Flexible-Scheduling", "Anxiety-or-Social-Support",
                                "Gear-or-Equipment", "Not-Applicable", "Not-Specified"],
        },
        "guide": (
            "Code whether the response contains actionable suggestions for how the school "
            "could help this student attend in future. "
            "Yes = specific and actionable, Partial = vague or implied, No = no suggestion. "
            "NA_Support_Type: primary type of support that would help them attend."
        ),
    },
]

ALL_NON_ATTEND_FIELDS = [f for qc in NON_ATTEND_CODEBOOK for f in qc["fields"]]


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
    "Write in clear paragraphs using Australian English spelling and vocabulary. "
    "Be specific and evidence-based."
)

def generate_narrative(narrative_model, narrative_tokenizer, qc, responses, status_label, root):
    """Summarise all responses using a map-reduce approach.

    Responses are first filtered for quality, then split into 3 equal chunks.
    Each chunk is summarised independently (map), and the chunk summaries are
    synthesised into a final narrative + best quotes (reduce).
    For small cohorts (<=50 valid responses) a single pass is used instead.
    """
    # ── Filter garbage / non-answers
    valid_responses = [r for r in responses if is_valid_response(r)]
    total_n         = len(valid_responses)
    raw_n           = sum(1 for r in responses if len(str(r).strip()) > 3)
    filtered_out    = raw_n - total_n

    if not valid_responses:
        return "Insufficient responses to generate a summary."

    filtered_note = (
        f" ({filtered_out} low-quality responses filtered before analysis)"
        if filtered_out > 0 else ""
    )

    # ── Single pass for small cohorts
    if total_n <= 50:
        status_label.config(
            text=f"Status: Writing Q{qc['num']} narrative ({total_n} responses)...",
            fg="#D97706")
        root.update()
        resp_text = "\n".join(f"- {r}" for r in valid_responses)
        prompt = (
            f'Summarise these student responses about: "{qc['label']}"\n\n'
            f'Write exactly 2 short paragraphs:\n'
            f'  1. The most common themes and overall student sentiment.\n'
            f'  2. What these responses reveal about student development or growth.\n\n'
            f'Then include exactly 3 to 5 highly valuable, verbatim quotes that best '
            f'capture the essence of the responses. Choose the most insightful quotes '
            f'regardless of where they appear in the list — do not feel obligated to '
            f'spread selection evenly. For each quote, add a brief contextual note in '
            f'square brackets before it only if it genuinely aids understanding '
            f'(e.g. [on coasteering]). Format with bullet points and quotation marks.\n\n'
            f'Student responses (n={total_n}){filtered_note}:\n{resp_text}'
        )
        return call_mlx(narrative_model, narrative_tokenizer, SYSTEM_NARRATOR,
                        prompt, max_tokens=600, thinking=False).strip()

    # ── Map-reduce for larger cohorts: divide into 3 equal chunks
    chunk_size = math.ceil(total_n / 3)
    chunks     = [valid_responses[i:i + chunk_size]
                  for i in range(0, total_n, chunk_size)]
    n_chunks   = len(chunks)

    chunk_summaries = []
    for i, chunk in enumerate(chunks, 1):
        status_label.config(
            text=f"Status: Q{qc['num']} — summarising batch {i}/{n_chunks} ({len(chunk)} responses)...",
            fg="#D97706")
        root.update()
        resp_text = "\n".join(f"- {r}" for r in chunk)
        map_prompt = (
            f'You are reading a batch of student responses about: "{qc['label']}".\n'
            f'This is batch {i} of {n_chunks} (responses {(i-1)*chunk_size+1}–'
            f'{min(i*chunk_size, total_n)} of {total_n} total).\n\n'
            f'Identify the key themes in this batch and copy out 3–5 of the most '
            f'insightful verbatim quotes from these responses. Be concise.\n\n'
            f'Student responses:\n{resp_text}'
        )
        summary = call_mlx(narrative_model, narrative_tokenizer, SYSTEM_NARRATOR,
                           map_prompt, max_tokens=450, thinking=False).strip()
        chunk_summaries.append(summary)

    # ── Reduce: synthesise all chunk summaries into the final narrative
    status_label.config(
        text=f"Status: Q{qc['num']} — synthesising across {n_chunks} batches...",
        fg="#D97706")
    root.update()

    combined = "\n\n---\n\n".join(
        f"BATCH {i} ({len(chunks[i-1])} responses):\n{s}"
        for i, s in enumerate(chunk_summaries, 1)
    )

    reduce_prompt = (
        f'Below are summaries of all {n_chunks} batches covering {total_n} student '
        f'responses about: "{qc['label']}"{filtered_note}.\n\n'
        f'Write the final report section:\n'
        f'  1. Two short paragraphs: (a) the most common themes and overall sentiment '
        f'across ALL responses, (b) what these responses reveal about student development.\n'
        f'  2. Exactly 3 to 5 of the very best verbatim student quotes drawn from any '
        f'batch — choose purely on quality and insight, not to represent each batch '
        f'equally. For each quote, add a brief contextual note in square brackets '
        f'only if it genuinely aids understanding. Format with bullet points and '
        f'quotation marks.\n\n'
        f'BATCH SUMMARIES:\n{combined}'
    )
    return call_mlx(narrative_model, narrative_tokenizer, SYSTEM_NARRATOR,
                    reduce_prompt, max_tokens=650, thinking=False).strip()


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
        valid_n = sum(1 for r in raw_vals if is_valid_response(r))
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
#  NON-ATTENDER ANALYSIS PIPELINE
# ═══════════════════════════════════════════════════════════════════

def _find_non_attend_text_cols(df):
    """Return up to 6 free-text columns from non-attending students."""
    priority_kw = [["reason"], ["why"], ["unable"], ["could not"],
                   ["feedback"], ["comment"], ["suggest"], ["help"],
                   ["barrier"], ["support"], ["improve"]]
    found = []
    for kws in priority_kw:
        col = find_col(df, kws)
        if col and col not in found:
            found.append(col)
    # For small groups (n < 30) use a much lower fill threshold (10%);
    # for larger groups keep 30%.
    min_fill = 0.10 if len(df) < 30 else 0.30
    for col in df.columns:
        if col in found:
            continue
        col_l = col.lower()
        # Skip obviously non-text columns
        if any(k in col_l for k in ("timestamp", "email", "score", "rating",
                                     "scale", "attend", "id", "number")):
            continue
        vals = df[col].dropna().astype(str)
        if len(vals) == 0:
            continue
        text_vals = vals[vals.str.len() > 5]
        if len(text_vals) / max(len(df), 1) < min_fill:
            continue
        non_num = text_vals[~text_vals.str.fullmatch(r'[\d\.\+\-\/\:\s,]+')]
        if len(non_num) / max(len(text_vals), 1) > 0.5:
            found.append(col)
    return found[:6]


def _combined_row_text(row, text_cols):
    """Concatenate all non-empty text fields for a DataFrame row."""
    parts = []
    for col in text_cols:
        v = str(row.get(col, "")).strip()
        if len(v) > 3 and v.lower() not in ("nan", "none", "n/a", "-"):
            parts.append(v)
    return " | ".join(parts) if parts else ""


def code_non_attend_question(model, tokenizer, qc, combined_texts,
                              status_label, root, BATCH=40):
    """Code all non-attender combined texts for one NON_ATTEND_CODEBOOK entry."""
    fields    = list(qc["fields"].keys())
    items     = [{"id": i, "text": t, "empty": len(t.strip()) <= 3}
                 for i, t in enumerate(combined_texts)]
    coded     = {}
    to_code   = [it for it in items if not it["empty"]]
    total_bat = max(1, (len(to_code) + BATCH - 1) // BATCH)

    for bn, start in enumerate(range(0, len(to_code), BATCH), 1):
        batch = to_code[start:start + BATCH]
        status_label.config(
            text=f"Status: Coding non-attenders ({qc['label']}) — batch {bn}/{total_bat}…",
            fg="#D97706")
        root.update()
        prompt      = build_coding_prompt(qc, batch)
        dynamic_max = max(512, int(len(batch) * (len(fields) * 8 + 15) * 1.2) + 64)
        raw         = call_mlx(model, tokenizer, SYSTEM_CODER, prompt,
                               max_tokens=dynamic_max, thinking=False)
        parsed      = safe_json(raw)
        if isinstance(parsed, list):
            for row in parsed:
                if isinstance(row, dict) and "id" in row:
                    coded[row["id"]] = {f: str(row.get(f, "Parse-Error")).strip()
                                        for f in fields}
        for it in batch:
            if it["id"] not in coded:
                coded[it["id"]] = {f: "Parse-Error" for f in fields}

    for it in items:
        if it["empty"]:
            coded[it["id"]] = {f: "No-Response" for f in fields}
    return coded


def build_non_attend_analysis(non_attend_df, model, tokenizer, status_label, root):
    """Full pipeline: code non-attender responses + generate AI narratives."""
    if non_attend_df is None or non_attend_df.empty:
        return pd.DataFrame(), []

    name_col  = find_col(non_attend_df, ["first", "name"]) or find_col(non_attend_df, ["name"])
    class_col = find_col(non_attend_df, ["class"])
    names     = non_attend_df[name_col].astype(str).tolist() if name_col else ["Student"] * len(non_attend_df)
    classes   = non_attend_df[class_col].astype(str).tolist() if class_col else ["Unknown"] * len(non_attend_df)

    text_cols      = _find_non_attend_text_cols(non_attend_df)
    # If no text columns detected, fall back to every string column except email/timestamp
    if not text_cols:
        text_cols = [
            c for c in non_attend_df.columns
            if non_attend_df[c].dtype == object
            and not any(k in c.lower() for k in ("email", "timestamp", "id"))
        ]
    combined_texts = [_combined_row_text(row, text_cols)
                      for _, row in non_attend_df.iterrows()]

    all_coded  = {i: {} for i in range(len(non_attend_df))}
    narratives = []

    for qc in NON_ATTEND_CODEBOOK:
        q_coded = code_non_attend_question(model, tokenizer, qc,
                                           combined_texts, status_label, root)
        for idx in range(len(non_attend_df)):
            all_coded[idx].update(q_coded.get(idx,
                {f: "Parse-Error" for f in qc["fields"]}))

        status_label.config(
            text=f"Status: Writing non-attender narrative ({qc['label']})…", fg="#D97706")
        root.update()
        sample = [t for t in combined_texts if is_valid_response(t)]
        if sample:
            resp_text = "\n".join(f"- {r}" for r in sample)
            valid_n   = sum(1 for t in combined_texts if is_valid_response(t))
            prompt = (
                f'Using Australian English, write a brief, concise summary of these '
                f'responses from students who did NOT attend camp, specifically about: '
                f'"{qc["label"]}"\n\n'
                f'This is a small group (n={valid_n}), so keep it tight: write ONE short '
                f'paragraph covering the key themes and any patterns. '
                f'Then list 2 to 3 verbatim student quotes that best capture the themes. '
                f'For each quote, add a brief contextual note in square brackets before it '
                f'if it helps the reader understand the context (e.g. [on equipment concerns]). '
                f'Only add context where it genuinely aids understanding. '
                f'Format quotes with bullet points and quotation marks.\n\n'
                f'Student responses (n={valid_n}):\n{resp_text}'
            )
            narr_text = call_mlx(model, tokenizer, SYSTEM_NARRATOR,
                                 prompt, max_tokens=400, thinking=False).strip()
        else:
            narr_text = "No usable responses found for this question."
            valid_n   = 0
        narratives.append({
            "Question": f"Q{qc['num']}: {qc['label']}",
            "n":        valid_n,
            "Summary":  narr_text,
        })

    na_df = pd.DataFrame({
        "Student_Name":  names,
        "Class":         classes,
        "Combined_Text": combined_texts,
    })
    for qc in NON_ATTEND_CODEBOOK:
        for field in qc["fields"]:
            na_df[field] = [all_coded.get(i, {}).get(field, "No-Response")
                            for i in range(len(non_attend_df))]
    return na_df, narratives


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

# Post-only metrics: no pre-camp baseline — reported as standalone scores.
# Both keywords are searched in post_df only.
POST_ONLY_METRICS = {
    "Y9 Camp Readiness": ["7 day camp", "year 9"],   # post-camp col 28
}

# ── Category groupings ───────────────────────────────────────────────
# Groups all METRIC_MAP keys into four reporting categories.
# Camping Skills is a sub-category of skills but reported separately.
METRIC_CATEGORIES = {
    "Activity Skills": ["Mountain Biking", "Biking Comfort", "Coasteering",
                        "Swimming", "Sea Kayaking", "Overnight Hike", "Snorkelling"],
    "Camping Skills":  ["Camping Skills", "Sleeping Outdoors"],
    "Attitudes":       ["Teamwork", "Autonomy", "Drive"],
    "Knowledge":       ["Aboriginal Culture"],
    "Agency in Learning": ["Agency in Learning (Y7\u2192Y8, 1\u201310)"],
}
CATEGORY_ORDER = ["Activity Skills", "Camping Skills", "Attitudes", "Knowledge", "Agency in Learning"]
CATEGORY_COLOURS = {
    "Activity Skills": "#2E8B88",   # teal
    "Camping Skills":  "#2D7A4F",   # green
    "Attitudes":       "#D97706",   # amber
    "Knowledge":       "#5B21B6",   # purple
    "Agency in Learning": "#4338CA", # indigo
}
CATEGORY_LIGHT_COLOURS = {
    "Activity Skills": "#D4F0EF",
    "Camping Skills":  "#C8EDD9",
    "Attitudes":       "#FEF3C7",
    "Knowledge":       "#EDE9FE",
    "Agency in Learning": "#E0E7FF",
}
CATEGORY_ICONS = {
    "Activity Skills": "🏄",
    "Camping Skills":  "⛺",
    "Attitudes":       "💪",
    "Knowledge":       "📚",
    "Agency in Learning": "🎯",
}


# ═══════════════════════════════════════════════════════════════════
#  AGENCY IN LEARNING  (Ruby Data Converter integration)
#  Blends in the "Agency in Learning" progress data extracted by the
#  Student Report Extractor (Ruby Data Converter) — a 1-5 progression
#  scale assessed by teachers. Year 2026 = this Y8 camp cohort.
#  Year 2025 = the same cohort's Y7 Maria Camp baseline, included only
#  as a CAUTIOUS comparison (different camp, different rater/timing,
#  and the "Class" field in that export is not reliable).
# ═══════════════════════════════════════════════════════════════════

AGENCY_DEFINITION = "The capacity to produce learning of value to self or community."

AGENCY_LEVEL_DESCRIPTIONS = {
    1: "Learners at this level use guidance from others to support participation in learning.",
    2: "Learners at this level learn by interpreting and following instructions, looking for "
       "guidance on what they should learn and how they should learn it.",
    3: "Learners at this level are skilled achievers who aspire to reach standards, making "
       "informed and deliberate decisions about their learning.",
    4: "Learners at this level are motivated to learn independently and from others, engaging "
       "with ideas and challenges to deepen their own understandings and competence.",
    5: "Learners at this level apply themselves relentlessly to their learning and are creative "
       "producers of knowledge, seeking to deepen and expand what they know and can do in "
       "domains of interest.",
}

RUBY_CAMP_YEAR        = "2026"   # this Y8 camp cohort
RUBY_BASELINE_YEAR     = "2025"   # prior-year Y7 Maria Camp baseline (cautious comparison only)
RUBY_HIGH_CONF_CUTOFF  = 0.95     # match-confidence treated as a reliable, exact-style match


def _agency_confidence_flag(row):
    """Classify a Ruby Data Converter row by how trustworthy its name/ID match is."""
    sid  = str(row.get("Student ID", "")).strip()
    conf = pd.to_numeric(row.get("Match Confidence", 0), errors="coerce")
    conf = 0.0 if pd.isna(conf) else conf
    if sid in ("", "Unknown / Not Found", "nan", "None"):
        return "Unmatched"
    if conf >= RUBY_HIGH_CONF_CUTOFF:
        return "High"
    return "Fuzzy \u2013 Review"


def load_agency_data(ruby_path):
    """
    Load the combined CSV exported by the Ruby Data Converter (Student Report
    Extractor) and turn it into a clean, one-row-per-student-per-year table of
    Agency in Learning scores.

    Expected columns: Year, NM Element, Class, School, Rater, Student ID,
    Email, Match Confidence, Extracted Name, Current Level, Progress in
    Level, Source File, Page.
    """
    raw = pd.read_csv(ruby_path)
    raw.columns = raw.columns.str.strip()

    required = ["Year", "Student ID", "Email", "Match Confidence",
                "Current Level", "Progress in Level"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(
            f"Agency in Learning CSV is missing expected column(s): {', '.join(missing)}. "
            f"Make sure you're uploading the combined CSV exported from the Student "
            f"Report Extractor (Ruby Data Converter), with the Email column included."
        )

    if "NM Element" in raw.columns:
        raw = raw[raw["NM Element"].astype(str).str.contains("Agency in Learning", case=False, na=False)]

    raw["Email"] = raw["Email"].astype(str).str.strip().str.lower()
    raw = raw[raw["Email"].str.contains("@", na=False)]   # drop blank / unmatched-email rows

    raw["Year"]  = raw["Year"].astype(str).str.strip()
    raw["Level Num"] = pd.to_numeric(
        raw["Current Level"].astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    raw["Progress"] = pd.to_numeric(raw["Progress in Level"], errors="coerce").fillna(0.0)
    # Convert to a 1-10 scale so Agency in Learning aligns with the rest of the
    # survey data (which is also out of 10).  Each level spans 2 points:
    #   Level 1 → 1–2  |  Level 2 → 3–4  |  Level 3 → 5–6
    #   Level 4 → 7–8  |  Level 5 → 9–10
    # Formula: (Level - 1) × 2 + 1 + Progress  (Progress is 0.0–1.0)
    raw["Agency Score"] = (raw["Level Num"] - 1) * 2 + 1 + raw["Progress"]
    raw["Match Confidence"] = pd.to_numeric(raw["Match Confidence"], errors="coerce").fillna(0.0)
    raw["Confidence Flag"]  = raw.apply(_agency_confidence_flag, axis=1)
    raw = raw.dropna(subset=["Level Num"])

    # A student can appear more than once per year (multiple scans/pages/elements).
    # Keep the highest-confidence row per Email+Year; ties broken by keeping the first.
    raw = raw.sort_values("Match Confidence", ascending=False)
    raw = raw.drop_duplicates(subset=["Email", "Year"], keep="first")
    return raw.reset_index(drop=True)


def build_agency_df(agency_raw, roster_df):
    """
    One row per camp student: this year's (Y8 camp, 2026) Agency in Learning
    score, last year's (Y7 Maria Camp, 2025) baseline where available, and a
    cautiously-labelled year-on-year comparison.

    roster_df should be tab1_df (or similar) with First Name / Surname /
    Email / Class / Location columns already resolved.
    """
    cols_keep = ["Email", "Current Level", "Progress in Level", "Agency Score",
                "NM Element", "Confidence Flag", "Match Confidence", "Extracted Name"]
    cols_keep = [c for c in cols_keep if c in agency_raw.columns]

    cur  = agency_raw[agency_raw["Year"] == RUBY_CAMP_YEAR][cols_keep].copy()
    base = agency_raw[agency_raw["Year"] == RUBY_BASELINE_YEAR][cols_keep].copy()
    cur  = cur.rename(columns={c: f"{c} (Y8 Camp 2026)" for c in cols_keep if c != "Email"})
    base = base.rename(columns={c: f"{c} (Y7 Maria Camp 2025)" for c in cols_keep if c != "Email"})

    roster_cols = [c for c in ["First Name", "Surname", "Email", "Class", "Location"]
                  if c in roster_df.columns]
    roster = roster_df[roster_cols].drop_duplicates(subset=["Email"]).copy()
    roster["Email"] = roster["Email"].astype(str).str.strip().str.lower()

    df = pd.merge(roster, cur,  on="Email", how="left")
    df = pd.merge(df,      base, on="Email", how="left")

    has_cur  = df["Agency Score (Y8 Camp 2026)"].notna() if "Agency Score (Y8 Camp 2026)" in df.columns else pd.Series(False, index=df.index)
    has_base = df["Agency Score (Y7 Maria Camp 2025)"].notna() if "Agency Score (Y7 Maria Camp 2025)" in df.columns else pd.Series(False, index=df.index)
    both     = has_cur & has_base

    df["Agency Growth (Y7\u2192Y8)"] = np.where(
        both,
        df.get("Agency Score (Y8 Camp 2026)", np.nan) - df.get("Agency Score (Y7 Maria Camp 2025)", np.nan),
        np.nan)

    def _reliability(i):
        if not has_cur.loc[i]:
            return "No 2026 (camp) data"
        if not both.loc[i]:
            return "No 2025 baseline \u2014 single point-in-time only"
        flags = {df.at[i, "Confidence Flag (Y8 Camp 2026)"] if "Confidence Flag (Y8 Camp 2026)" in df.columns else None,
                df.at[i, "Confidence Flag (Y7 Maria Camp 2025)"] if "Confidence Flag (Y7 Maria Camp 2025)" in df.columns else None}
        if "Unmatched" in flags:
            return "Unmatched \u2014 exclude from comparison"
        if "Fuzzy \u2013 Review" in flags:
            return "Comparison available \u2014 verify name match"
        return "Comparison available \u2014 different camp/rater & ~1yr apart, interpret cautiously"

    df["Comparison Reliability"] = [_reliability(i) for i in df.index]
    return df


def build_agency_summary(agency_df):
    """Aggregate stats for the dedicated Agency in Learning section/sheet."""
    if agency_df is None or agency_df.empty or "Agency Score (Y8 Camp 2026)" not in agency_df.columns:
        return {}
    cur_scores = agency_df["Agency Score (Y8 Camp 2026)"].dropna()
    growth     = agency_df["Agency Growth (Y7\u2192Y8)"].dropna() if "Agency Growth (Y7\u2192Y8)" in agency_df.columns else pd.Series(dtype=float)
    level_dist = (agency_df["Current Level (Y8 Camp 2026)"].dropna().value_counts().sort_index().to_dict()
                 if "Current Level (Y8 Camp 2026)" in agency_df.columns else {})
    conf_dist  = (agency_df["Confidence Flag (Y8 Camp 2026)"].dropna().value_counts().to_dict()
                 if "Confidence Flag (Y8 Camp 2026)" in agency_df.columns else {})
    return {
        "n_2026":         int(len(cur_scores)),
        "n_both_years":   int(len(growth)),
        "avg_2026":       round(float(cur_scores.mean()), 2) if len(cur_scores) else None,
        "level_dist":     level_dist,
        "confidence_dist": conf_dist,
        "avg_growth":     round(float(growth.mean()), 2) if len(growth) else None,
        "pct_grew":       round(len(growth[growth > 0]) / len(growth) * 100, 1) if len(growth) else None,
    }

def get_metric_category(metric_name):
    """Return the reporting category for a given metric name."""
    for cat, metrics in METRIC_CATEGORIES.items():
        if metric_name in metrics:
            return cat
    return "Other"

def build_category_summary(avg_df, metrics_processed):
    """Compute per-category avg shift, median shift, and % improvers from avg_df."""
    rows = []
    for cat in CATEGORY_ORDER:
        cat_rows = avg_df[avg_df["Metric"].isin(METRIC_CATEGORIES.get(cat, []))]
        if cat_rows.empty:
            continue
        rows.append({
            "Category":       cat,
            "Avg Shift":      round(float(cat_rows["Avg Shift"].mean()), 2),
            "Median Shift":   round(float(cat_rows["Median Shift"].mean()), 2),
            "% Improvers":    round(float(cat_rows["% Improvers"].mean()), 1),
            "Metrics Tracked": len(cat_rows),
        })
    return pd.DataFrame(rows)

def process_quantitative(students_df, pre_df, post_df, ruby_path=None):
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
        mp_suffix = "_pre"
        if not pc:
            # Pre-camp keyword not in pre_df — check post_df (e.g. Aboriginal Culture uses
            # a retrospective "Thinking back…" question on the post-camp form as its baseline).
            pc = find_col(post_df, pk)
            mp_suffix = "_post"
        if not pc or not qc:
            continue
        mp = pc + mp_suffix if pc + mp_suffix in merged.columns else pc
        mq = qc + "_post"   if qc + "_post"   in merged.columns else qc
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
                     "Avg Shift": vq.mean() - vp.mean(),
                     "Median Shift": float(vd.median()),
                     "% Improvers": round(len(vd[vd > 0]) / len(vd) * 100, 1) if len(vd) else 0})
        dists.append({"Metric": name,
                      "Improved":    len(vd[vd > 0]) / len(vd) if len(vd) else 0,
                      "Stayed Same": len(vd[vd == 0]) / len(vd) if len(vd) else 0,
                      "Declined":    len(vd[vd < 0]) / len(vd) if len(vd) else 0})

    tab1_df = pd.DataFrame(tab1_data)
    core_metrics = list(metrics)   # snapshot before Agency (different timeframe) is folded in

    # ── Agency in Learning (Ruby Data Converter) ──────────────────────
    agency_df, agency_summary = pd.DataFrame(), {}
    if ruby_path:
        try:
            agency_raw = load_agency_data(ruby_path)
            agency_df  = build_agency_df(agency_raw, tab1_df)
            agency_summary = build_agency_summary(agency_df)

            growth_map = dict(zip(agency_df["Email"], agency_df["Agency Growth (Y7\u2192Y8)"]))
            agency_metric_name = "Agency in Learning (Y7\u2192Y8, 1\u201310)"
            tab1_df[agency_metric_name] = tab1_df["Email"].map(growth_map)
            metrics.append(agency_metric_name)

            gvals = tab1_df[agency_metric_name].dropna()
            if len(gvals):
                cur_map  = dict(zip(agency_df["Email"], agency_df.get("Agency Score (Y8 Camp 2026)", pd.Series(dtype=float))))
                base_map = dict(zip(agency_df["Email"], agency_df.get("Agency Score (Y7 Maria Camp 2025)", pd.Series(dtype=float))))
                cur_vals  = tab1_df["Email"].map(cur_map).dropna()
                base_vals = tab1_df["Email"].map(base_map).dropna()
                avgs.append({
                    "Metric": agency_metric_name,
                    "Pre-Camp Avg":  round(float(base_vals.mean()), 2) if len(base_vals) else float("nan"),
                    "Post-Camp Avg": round(float(cur_vals.mean()), 2) if len(cur_vals) else float("nan"),
                    "Avg Shift":     float(gvals.mean()),
                    "Median Shift":  float(gvals.median()),
                    "% Improvers":   round(len(gvals[gvals > 0]) / len(gvals) * 100, 1) if len(gvals) else 0,
                })
                dists.append({
                    "Metric": agency_metric_name,
                    "Improved":    len(gvals[gvals > 0]) / len(gvals) if len(gvals) else 0,
                    "Stayed Same": len(gvals[gvals == 0]) / len(gvals) if len(gvals) else 0,
                    "Declined":    len(gvals[gvals < 0]) / len(gvals) if len(gvals) else 0,
                })

        except Exception as e:
            print(f"Agency in Learning data warning: {e}")

    unique_locs    = [l for l in tab1_df["Location"].unique() if str(l).lower() not in ("unknown","nan")]
    unique_classes = [c for c in tab1_df["Class"].unique()    if str(c).lower() not in ("unknown","nan")]
    bd_rows = []
    for cls in sorted(unique_classes):
        sub = tab1_df[tab1_df["Class"] == cls]
        row = {"Class": cls, "Location": "ALL LOCATIONS",
               "Average Shift (All Areas)": sub[core_metrics].mean().mean()}
        for m in metrics:
            row[m] = sub[m].mean()
        bd_rows.append(row)
        for loc in sorted(unique_locs):
            sl = sub[sub["Location"] == loc]
            if not sl.empty:
                row2 = {"Class": f"  ↳ {loc}", "Location": loc,
                        "Average Shift (All Areas)": sl[core_metrics].mean().mean()}
                for m in metrics:
                    row2[m] = sl[m].mean()
                bd_rows.append(row2)

    post_only_scores = _compute_post_only_scores(post_df, merged)
    if agency_summary.get("avg_2026") is not None:
        post_only_scores["Agency in Learning \u2014 2026 Camp Snapshot (all matched students)"] = {
            "avg": agency_summary["avg_2026"], "n": agency_summary["n_2026"]}

    return (tab1_df, pd.DataFrame(avgs), pd.DataFrame(dists), pd.DataFrame(bd_rows), merged, metrics,
            post_only_scores, agency_df, agency_summary)


def _compute_post_only_scores(post_df, merged):
    """Compute averages for POST_ONLY_METRICS (no pre-camp baseline exists)."""
    scores = {}
    for name, keywords in POST_ONLY_METRICS.items():
        col = find_col(post_df, keywords)
        if not col:
            continue
        mc = col + "_post" if col + "_post" in merged.columns else col
        if mc not in merged.columns:
            mc = col if col in merged.columns else None
        if not mc:
            continue
        vals = pd.to_numeric(merged[mc], errors="coerce").dropna()
        if len(vals):
            scores[name] = {"avg": round(float(vals.mean()), 2), "n": int(len(vals))}
    return scores


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
                metrics_processed, mm_tables, qual_charts, post_only_scores=None,
                na_df=None, na_narratives=None, agency_df=None, agency_summary=None):

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
    ws_na   = workbook.add_worksheet("🚫 Non-Attenders")
    ws_cat  = workbook.add_worksheet("📂 Category Summary")

    coded_df.to_excel(writer, sheet_name="Qual Responses", index=False, na_rep="—")
    long_df.to_excel(writer,  sheet_name="Longitudinal Data", index=False, na_rep="—")

    ws_dash  = workbook.add_worksheet("_QualDash")
    ws_cdraw = workbook.add_worksheet("_ChartData")
    ws_qcd   = workbook.add_worksheet("_QualChartData")
    ws_cdraw.hide(); ws_qcd.hide(); ws_dash.hide()
    writer.sheets["Qual Summary"]        = ws_sum
    writer.sheets["Qual Highlights"]     = ws_high
    writer.sheets["Mixed Methods"]       = ws_mm
    writer.sheets["🚫 Non-Attenders"]   = ws_na
    writer.sheets["📂 Category Summary"] = ws_cat
    writer.sheets["Qual Responses"]      = writer.sheets["Qual Responses"]
    writer.sheets["Longitudinal Data"]   = writer.sheets["Longitudinal Data"]
    writer.sheets["_QualDash"]           = ws_dash
    writer.sheets["_ChartData"]          = ws_cdraw
    writer.sheets["_QualChartData"]      = ws_qcd

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

    # Section: Quantitative highlights — now grouped by category
    ws_exec.set_row(8, 20)
    sec_fmt = workbook.add_format({"bold": True, "font_size": 11, "font_name": "Arial",
                                   "bg_color": C["navy"], "font_color": C["white"]})
    ws_exec.merge_range("B9:I9", "  📐  Skill & Attitude Shifts by Category  (avg change, pre → post camp; Agency in Learning = Y7 → Y8, see note)", sec_fmt)

    # Mini table: metric shifts grouped by category
    if not avg_df.empty and "Metric" in avg_df.columns:
        ws_exec.set_row(9, 16)
        hdr_fmt = workbook.add_format({"bold": True, "bg_color": C["grey_light"],
                                        "font_name": "Arial", "border_color": "#CCCCCC", "border": 1, "align": "center"})
        ws_exec.write(9, 1, "Skill / Area",    hdr_fmt)
        ws_exec.write(9, 2, "Pre-Camp",        hdr_fmt)
        ws_exec.write(9, 3, "Post-Camp",       hdr_fmt)
        ws_exec.merge_range(9, 4, 9, 5, "Avg Shift", hdr_fmt)
        ws_exec.write(9, 6, "Median Shift",    hdr_fmt)
        ws_exec.write(9, 7, "% Improvers",     hdr_fmt)
        ws_exec.write(9, 8, "Direction",        hdr_fmt)

        er = 10  # current excel row
        for cat in CATEGORY_ORDER:
            cat_metrics_in_data = [r for _, r in avg_df.iterrows()
                                   if r.get("Metric","") in METRIC_CATEGORIES.get(cat, [])]
            if not cat_metrics_in_data:
                continue
            # ── Category header row
            cat_colour  = CATEGORY_COLOURS.get(cat, C["teal"])
            cat_light   = CATEGORY_LIGHT_COLOURS.get(cat, C["teal_light"])
            cat_hdr_fmt = workbook.add_format({
                "bold": True, "font_size": 10, "font_name": "Arial",
                "bg_color": cat_colour, "font_color": "#FFFFFF", "border": 1,
            })
            cat_avg_shift = sum(r.get("Avg Shift", 0) or 0 for r in cat_metrics_in_data) / len(cat_metrics_in_data)
            cat_dir = "▲ Better" if cat_avg_shift > 0.05 else ("▼ Declined" if cat_avg_shift < -0.05 else "— Same")
            ws_exec.set_row(er, 18)
            ws_exec.write(er, 1, f"{CATEGORY_ICONS.get(cat, '')}  {cat}", cat_hdr_fmt)
            for cc in range(2, 8):
                ws_exec.write(er, cc, "", cat_hdr_fmt)
            ws_exec.write(er, 8, cat_dir, cat_hdr_fmt)
            er += 1

            for ri, row_data in enumerate(cat_metrics_in_data):
                ws_exec.set_row(er, 16)
                row_bg  = cat_light if ri % 2 == 0 else C["white"]
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
                ws_exec.write(er, 1, "  " + row_data.get("Metric", ""), base_fmt)
                ws_exec.write(er, 2, row_data.get("Pre-Camp Avg", ""), num_fmt)
                ws_exec.write(er, 3, row_data.get("Post-Camp Avg", ""), num_fmt)
                ws_exec.merge_range(er, 4, er, 5, shift, sf)
                # Median Shift
                med = row_data.get("Median Shift", 0) or 0
                if med > 0.05:
                    med_sf = workbook.add_format({"bg_color": C["green_light"], "font_color": C["green"],
                                                   "bold": True, "num_format": "+0.0;-0.0;0.0",
                                                   "align": "center", "font_name": "Arial"})
                elif med < -0.05:
                    med_sf = workbook.add_format({"bg_color": C["red_light"], "font_color": C["red"],
                                                   "bold": True, "num_format": "+0.0;-0.0;0.0",
                                                   "align": "center", "font_name": "Arial"})
                else:
                    med_sf = num_fmt
                ws_exec.write(er, 6, med, med_sf)
                # % Improvers
                pct_imp = row_data.get("% Improvers", 0) or 0
                if pct_imp >= 55:
                    pct_sf = workbook.add_format({"bg_color": C["green_light"], "font_color": C["green"],
                                                   "bold": True, "num_format": "0.0\"%\"",
                                                   "align": "center", "font_name": "Arial"})
                elif pct_imp < 35:
                    pct_sf = workbook.add_format({"bg_color": C["red_light"], "font_color": C["red"],
                                                   "bold": True, "num_format": "0.0\"%\"",
                                                   "align": "center", "font_name": "Arial"})
                else:
                    pct_sf = workbook.add_format({"bg_color": C["amber_light"], "font_color": C["amber"],
                                                   "num_format": "0.0\"%\"",
                                                   "align": "center", "font_name": "Arial"})
                ws_exec.write(er, 7, pct_imp, pct_sf)
                ws_exec.write(er, 8, dir_str, dir_fmt)
                er += 1

    # ── Post-only metric rows (e.g. Y9 Camp Readiness — no pre-camp baseline)
    # er is already set by the category loop above; fall back if avg_df was empty
    if avg_df.empty or "Metric" not in avg_df.columns:
        er = 10
    po_extra = 0
    if post_only_scores:
        for po_name, po_data in post_only_scores.items():
            er_po = er + po_extra
            ws_exec.set_row(er_po, 16)
            po_base_fmt = workbook.add_format({
                "bg_color": C["teal_light"], "font_name": "Arial", "font_size": 10,
                "border_color": "#E5E7EB", "border": 1
            })
            po_num_fmt = workbook.add_format({
                "bg_color": C["teal_light"], "font_name": "Arial", "font_size": 10,
                "num_format": "0.0", "align": "center",
                "border_color": "#E5E7EB", "border": 1
            })
            po_tag_fmt = workbook.add_format({
                "bg_color": C["teal_light"], "font_color": C["teal"], "font_name": "Arial",
                "font_size": 9, "italic": True, "align": "center",
                "border_color": "#E5E7EB", "border": 1
            })
            ws_exec.write(er_po, 1, f"★  {po_name}", po_base_fmt)
            ws_exec.write(er_po, 2, "—", po_tag_fmt)
            ws_exec.write(er_po, 3, po_data["avg"], po_num_fmt)
            ws_exec.merge_range(er_po, 4, er_po, 7,
                                "Post-camp score only  (no pre-camp baseline)", po_tag_fmt)
            ws_exec.write(er_po, 8, f"n = {po_data['n']}", po_tag_fmt)
            po_extra += 1

    # Section: Qualitative snapshot
    q_sec_row = er + po_extra + 2
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
    #  📂 CATEGORY SUMMARY SHEET
    # ────────────────────────────────────────────────────────────────
    ws_cat.hide_gridlines(2)
    ws_cat.set_column("A:A", 3)
    ws_cat.set_column("B:B", 26)
    ws_cat.set_column("C:F", 18)
    ws_cat.set_column("G:G", 22)

    cat_title_fmt = workbook.add_format({
        "bold": True, "font_size": 16, "font_name": "Arial",
        "bg_color": C["navy"], "font_color": C["white"],
        "align": "center", "valign": "vcenter",
    })
    ws_cat.merge_range("B1:G2", "📂  Category Summary — Avg Shift by Domain", cat_title_fmt)
    ws_cat.set_row(0, 30); ws_cat.set_row(1, 30)
    ws_cat.merge_range("B3:G3",
        "Grouped view: Activity Skills · Camping Skills · Attitudes · Knowledge",
        workbook.add_format({"font_size": 10, "font_name": "Arial",
                             "bg_color": C["teal"], "font_color": C["white"], "align": "center"}))
    ws_cat.set_row(2, 18)

    cat_col_hdr_fmt = workbook.add_format({
        "bold": True, "bg_color": C["grey_light"], "border": 1,
        "font_name": "Arial", "align": "center", "font_size": 10,
    })
    cat_row_start = 4
    ws_cat.set_row(cat_row_start, 16)
    ws_cat.write(cat_row_start, 1, "Category",         cat_col_hdr_fmt)
    ws_cat.write(cat_row_start, 2, "Avg Shift",         cat_col_hdr_fmt)
    ws_cat.write(cat_row_start, 3, "Median Shift",      cat_col_hdr_fmt)
    ws_cat.write(cat_row_start, 4, "% Improvers",       cat_col_hdr_fmt)
    ws_cat.write(cat_row_start, 5, "Metrics Tracked",   cat_col_hdr_fmt)
    ws_cat.write(cat_row_start, 6, "Direction",         cat_col_hdr_fmt)

    # Hidden data sheet for chart
    ws_cat_data = workbook.add_worksheet("_CatSummaryData")
    ws_cat_data.hide()
    writer.sheets["_CatSummaryData"] = ws_cat_data

    cat_data_row = 0
    ws_cat_data.write(cat_data_row, 0, "Category")
    ws_cat_data.write(cat_data_row, 1, "Avg Shift")
    cat_data_row += 1

    cat_sr = cat_row_start + 1
    cat_chart_n = 0
    for cat in CATEGORY_ORDER:
        cat_rows = avg_df[avg_df["Metric"].isin(METRIC_CATEGORIES.get(cat, []))] if not avg_df.empty else pd.DataFrame()
        if cat_rows.empty:
            continue
        cat_colour = CATEGORY_COLOURS.get(cat, C["teal"])
        cat_light  = CATEGORY_LIGHT_COLOURS.get(cat, C["teal_light"])
        cat_avg    = float(cat_rows["Avg Shift"].mean())
        cat_med    = float(cat_rows["Median Shift"].mean()) if "Median Shift" in cat_rows.columns else 0
        cat_pct    = float(cat_rows["% Improvers"].mean()) if "% Improvers" in cat_rows.columns else 0
        cat_n      = len(cat_rows)
        cat_dir    = "▲ Improved" if cat_avg > 0.05 else ("▼ Declined" if cat_avg < -0.05 else "— Same")

        row_fmt = workbook.add_format({"bg_color": cat_light, "font_name": "Arial",
                                        "font_size": 11, "border": 1, "bold": True,
                                        "font_color": cat_colour})
        num_fmt_c = workbook.add_format({"bg_color": cat_light, "font_name": "Arial",
                                          "font_size": 11, "border": 1,
                                          "num_format": "+0.00;-0.00;0.00", "align": "center",
                                          "font_color": C["green"] if cat_avg > 0.05 else (C["red"] if cat_avg < -0.05 else C["grey"])})
        pct_fmt_c = workbook.add_format({"bg_color": cat_light, "font_name": "Arial",
                                          "font_size": 11, "border": 1,
                                          "num_format": "0.0\"%\"", "align": "center"})
        ctr_fmt_c = workbook.add_format({"bg_color": cat_light, "font_name": "Arial",
                                          "font_size": 11, "border": 1, "align": "center"})
        dir_fmt_c = workbook.add_format({"bg_color": cat_light, "font_name": "Arial",
                                          "font_size": 11, "border": 1, "align": "center",
                                          "font_color": C["green"] if cat_avg > 0.05 else (C["red"] if cat_avg < -0.05 else C["grey"]),
                                          "bold": True})
        ws_cat.set_row(cat_sr, 22)
        ws_cat.write(cat_sr, 1, f"{CATEGORY_ICONS.get(cat,'')}  {cat}", row_fmt)
        ws_cat.write(cat_sr, 2, cat_avg, num_fmt_c)
        ws_cat.write(cat_sr, 3, cat_med, num_fmt_c)
        ws_cat.write(cat_sr, 4, cat_pct, pct_fmt_c)
        ws_cat.write(cat_sr, 5, cat_n,   ctr_fmt_c)
        ws_cat.write(cat_sr, 6, cat_dir, dir_fmt_c)
        cat_sr += 1

        # Write hidden data for chart
        ws_cat_data.write(cat_data_row, 0, cat)
        ws_cat_data.write(cat_data_row, 1, cat_avg)
        cat_data_row += 1
        cat_chart_n += 1

        # Per-category metric detail rows below the summary row
        for mi, (_, mrow) in enumerate(cat_rows.iterrows()):
            ws_cat.set_row(cat_sr, 16)
            sub_bg  = C["offwhite"] if mi % 2 == 0 else C["white"]
            sub_fmt = workbook.add_format({"bg_color": sub_bg, "font_name": "Arial",
                                            "font_size": 10, "border": 1})
            sub_num = workbook.add_format({"bg_color": sub_bg, "font_name": "Arial",
                                            "font_size": 10, "border": 1,
                                            "num_format": "+0.00;-0.00;0.00", "align": "center"})
            sub_pct = workbook.add_format({"bg_color": sub_bg, "font_name": "Arial",
                                            "font_size": 10, "border": 1,
                                            "num_format": "0.0\"%\"", "align": "center"})
            sub_ctr = workbook.add_format({"bg_color": sub_bg, "font_name": "Arial",
                                            "font_size": 10, "border": 1, "align": "center"})
            mv = float(mrow.get("Avg Shift", 0) or 0)
            ws_cat.write(cat_sr, 1, "    → " + str(mrow.get("Metric","")), sub_fmt)
            ws_cat.write(cat_sr, 2, mv, sub_num)
            ws_cat.write(cat_sr, 3, float(mrow.get("Median Shift", 0) or 0), sub_num)
            ws_cat.write(cat_sr, 4, float(mrow.get("% Improvers", 0) or 0), sub_pct)
            ws_cat.write(cat_sr, 5, 1, sub_ctr)
            ws_cat.write(cat_sr, 6, "▲" if mv > 0.05 else ("▼" if mv < -0.05 else "—"), sub_ctr)
            cat_sr += 1

        ws_cat.set_row(cat_sr, 6)   # spacer between categories
        cat_sr += 1

    # Category bar chart
    if cat_chart_n > 0:
        cat_chart = workbook.add_chart({"type": "bar"})
        colours_list = [CATEGORY_COLOURS.get(c, C["teal"]) for c in CATEGORY_ORDER
                        if not avg_df[avg_df["Metric"].isin(METRIC_CATEGORIES.get(c, []))].empty
                        if not avg_df.empty]
        cat_chart.add_series({
            "name":       "Avg Shift",
            "categories": ["_CatSummaryData", 1, 0, cat_chart_n, 0],
            "values":     ["_CatSummaryData", 1, 1, cat_chart_n, 1],
            "fill":       {"color": "#2E8B88"},
            "data_labels": {"value": True, "num_format": "+0.00;-0.00;0.00"},
        })
        cat_chart.set_title({"name": "Average Shift by Category (pre → post camp)",
                             "name_font": {"size": 14, "bold": True}})
        cat_chart.set_x_axis({"name": "Avg Shift", "crossing": 0})
        cat_chart.set_y_axis({"reverse": True})
        cat_chart.set_legend({"none": True})
        cat_chart.set_size({"width": 500, "height": 280})
        ws_cat.insert_chart(f"B{cat_sr + 2}", cat_chart)

    # ────────────────────────────────────────────────────────────────
    #  GROUP AVERAGES (cleaner with visual bars via data bars)
    # ────────────────────────────────────────────────────────────────
    ws = writer.sheets["Group Averages"]
    for i, col in enumerate(avg_df.columns):
        ws.write(0, i, col, F["col_hdr"])
    ws.set_column("A:A", 26)
    ws.set_column("B:D", 18)
    ws.set_column("E:F", 18)

    # Write metrics grouped by category with coloured section headers
    ga_r = 1
    for cat in CATEGORY_ORDER:
        cat_rows = [(i, r) for i, r in avg_df.iterrows()
                    if r.get("Metric","") in METRIC_CATEGORIES.get(cat, [])]
        if not cat_rows:
            continue
        cat_colour = CATEGORY_COLOURS.get(cat, C["teal"])
        cat_hdr_f  = workbook.add_format({
            "bold": True, "font_size": 10, "font_name": "Arial",
            "bg_color": cat_colour, "font_color": "#FFFFFF", "border": 1,
        })
        # Category header spanning all columns
        n_cols = len(avg_df.columns)
        ws.set_row(ga_r, 18)
        ws.write(ga_r, 0, f"{CATEGORY_ICONS.get(cat,'')}  {cat}", cat_hdr_f)
        for cc in range(1, n_cols):
            ws.write(ga_r, cc, "", cat_hdr_f)
        ga_r += 1
        for local_ri, (_, row) in enumerate(cat_rows):
            cat_light = CATEGORY_LIGHT_COLOURS.get(cat, C["teal_light"])
            row_bg = cat_light if local_ri % 2 == 0 else C["white"]
            base = workbook.add_format({"bg_color": row_bg, "font_name": "Arial",
                                         "font_size": 10, "num_format": "0.0", "align": "center"})
            label_fmt = workbook.add_format({"bg_color": row_bg, "font_name": "Arial", "font_size": 10})
            ws.write(ga_r, 0, row.iloc[0], label_fmt)
            ws.write_number(ga_r, 1, row.iloc[1], base)
            ws.write_number(ga_r, 2, row.iloc[2], base)
            avg_shift = row.iloc[3]
            ws.write_number(ga_r, 3, avg_shift, shift_fmt(avg_shift))
            med_shift = row.iloc[4] if len(avg_df.columns) > 4 else 0
            ws.write_number(ga_r, 4, med_shift, shift_fmt(med_shift))
            pct_imp = row.iloc[5] if len(avg_df.columns) > 5 else 0
            if pct_imp >= 55:
                pct_cell_fmt = workbook.add_format({"bg_color": C["green_light"], "font_color": C["green"],
                                                    "bold": True, "num_format": "0.0\"%\"", "align": "center",
                                                    "font_name": "Arial", "font_size": 10})
            elif pct_imp < 35:
                pct_cell_fmt = workbook.add_format({"bg_color": C["red_light"], "font_color": C["red"],
                                                    "bold": True, "num_format": "0.0\"%\"", "align": "center",
                                                    "font_name": "Arial", "font_size": 10})
            else:
                pct_cell_fmt = workbook.add_format({"bg_color": C["amber_light"], "font_color": C["amber"],
                                                    "num_format": "0.0\"%\"", "align": "center",
                                                    "font_name": "Arial", "font_size": 10})
            ws.write_number(ga_r, 5, pct_imp, pct_cell_fmt)
            ga_r += 1
        # Category average row
        cat_all_shifts = [r.iloc[3] for _, r in cat_rows]
        cat_all_pre    = [r.iloc[1] for _, r in cat_rows]
        cat_all_post   = [r.iloc[2] for _, r in cat_rows]
        cat_avg_shift  = float(np.nanmean(cat_all_shifts)) if cat_all_shifts else 0
        cat_avg_pre    = float(np.nanmean(cat_all_pre))    if cat_all_pre    else 0
        cat_avg_post   = float(np.nanmean(cat_all_post))   if cat_all_post   else 0
        cat_sum_fmt = workbook.add_format({
            "bold": True, "bg_color": cat_colour, "font_color": "#FFFFFF",
            "font_name": "Arial", "font_size": 10,
            "num_format": "0.0", "align": "center", "italic": True,
        })
        cat_sum_lbl = workbook.add_format({
            "bold": True, "bg_color": cat_colour, "font_color": "#FFFFFF",
            "font_name": "Arial", "font_size": 10, "italic": True,
        })
        cat_sum_shift = workbook.add_format({
            "bold": True, "bg_color": cat_colour, "font_color": "#FFFFFF",
            "font_name": "Arial", "font_size": 10,
            "num_format": "+0.0;-0.0;0.0", "align": "center", "italic": True,
        })
        ws.set_row(ga_r, 16)
        ws.write(ga_r, 0, f"  ↳ {cat} Average", cat_sum_lbl)
        ws.write_number(ga_r, 1, cat_avg_pre,   cat_sum_fmt)
        ws.write_number(ga_r, 2, cat_avg_post,  cat_sum_fmt)
        ws.write_number(ga_r, 3, cat_avg_shift, cat_sum_shift)
        for cc in range(4, len(avg_df.columns)):
            ws.write(ga_r, cc, "", cat_sum_fmt)
        ga_r += 2  # spacer row after each category

    ws.conditional_format(f"B2:C{ga_r}",
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

    # ────────────────────────────────────────────────────────────────
    #  🚫 NON-ATTENDERS SHEET
    # ────────────────────────────────────────────────────────────────
    ws_na.hide_gridlines(2)
    ws_na.set_column("A:A", 3)
    ws_na.set_column("B:B", 22)
    ws_na.set_column("C:C", 22)
    ws_na.set_column("D:D", 60)
    ws_na.set_column("E:Z", 24)

    # Title banner
    na_title_fmt = workbook.add_format({
        "bold": True, "font_size": 16, "font_name": "Arial",
        "bg_color": "#7C3AED", "font_color": C["white"],
        "align": "center", "valign": "vcenter",
    })
    na_sub_fmt = workbook.add_format({
        "font_size": 10, "font_name": "Arial",
        "bg_color": "#DDD6FE", "font_color": "#4C1D95",
        "align": "center", "valign": "vcenter", "italic": True,
    })
    ws_na.merge_range("B1:I2", "🚫  Non-Attender Analysis", na_title_fmt)
    ws_na.set_row(0, 28); ws_na.set_row(1, 28)
    n_na = len(na_df) if na_df is not None and not na_df.empty else 0
    ws_na.merge_range("B3:I3",
        f"Students who did not attend camp  |  n = {n_na}  |  "
        "Reasons & Constructive Feedback", na_sub_fmt)
    ws_na.set_row(2, 18)

    na_sec_fmt = workbook.add_format({
        "bold": True, "font_size": 11, "font_name": "Arial",
        "bg_color": "#7C3AED", "font_color": C["white"],
    })
    na_field_fmt = workbook.add_format({
        "bold": True, "font_size": 10, "font_name": "Arial",
        "bg_color": "#EDE9FE", "font_color": "#4C1D95", "italic": True,
    })
    na_wrap_fmt = workbook.add_format({
        "text_wrap": True, "valign": "top", "font_name": "Arial", "font_size": 10,
    })

    # ── Section 1: AI Narratives
    row_idx = 4
    ws_na.set_row(row_idx, 20)
    ws_na.merge_range(row_idx, 1, row_idx, 8,
                      "  💬  AI Narrative Summaries", na_sec_fmt)
    row_idx += 1

    if na_narratives:
        for nar in na_narratives:
            ws_na.set_row(row_idx, 22)
            ws_na.write(row_idx, 1, nar.get("Question", ""), F["section"])
            ws_na.write(row_idx, 2, f"n = {nar.get('n', 0)}", F["section"])
            for c in range(3, 9):
                ws_na.write(row_idx, c, "", F["section"])
            row_idx += 1

            summary_text = str(nar.get("Summary", ""))
            lines = len(summary_text.split('\n')) + (len(summary_text) // 90)
            ws_na.set_row(row_idx, max(80, lines * 15))
            ws_na.write(row_idx, 1, "Analysis & Quotes", na_field_fmt)
            ws_na.write(row_idx, 2, summary_text, na_wrap_fmt)
            row_idx += 2
            ws_na.set_row(row_idx, 6)   # spacer
            row_idx += 2
    else:
        ws_na.write(row_idx, 1,
                    "No non-attender responses found, or AI analysis was not run.",
                    F["skip"])
        row_idx += 2

    # ── Section 2: Coded responses table
    ws_na.set_row(row_idx, 20)
    ws_na.merge_range(row_idx, 1, row_idx, 8,
                      "  📋  Coded Responses — Individual Students", na_sec_fmt)
    row_idx += 1

    if na_df is not None and not na_df.empty:
        # Header row
        ws_na.set_row(row_idx, 16)
        na_col_hdr = workbook.add_format({
            "bold": True, "bg_color": "#EDE9FE", "font_color": "#4C1D95",
            "font_name": "Arial", "font_size": 10, "border": 1,
            "text_wrap": True, "align": "center",
        })
        for ci, col_name in enumerate(na_df.columns):
            ws_na.write(row_idx, 1 + ci, col_name, na_col_hdr)
        row_idx += 1

        # Value colour maps for NA fields
        na_reason_colours = {
            "Health-or-Medical":    ("#FEF3C7", "#D97706"),
            "Family-Commitment":    ("#EDE9FE", "#5B21B6"),
            "Financial-Cost":       ("#FEE2E2", "#C0392B"),
            "Fear-or-Anxiety":      ("#FEE2E2", "#C0392B"),
            "Another-School-Event": ("#D1FAE5", "#065F46"),
            "Personal-Choice":      ("#F3F4F6", "#6B7280"),
            "Logistical":           ("#FEF3C7", "#D97706"),
            "Not-Specified":        ("#F9FAFB", "#9CA3AF"),
        }
        na_yes_fmt = workbook.add_format({
            "bg_color": C["green_light"], "font_color": C["green"],
            "bold": True, "font_name": "Arial", "font_size": 10,
        })
        na_par_fmt = workbook.add_format({
            "bg_color": C["amber_light"], "font_color": C["amber"],
            "font_name": "Arial", "font_size": 10,
        })
        na_no_fmt = workbook.add_format({
            "bg_color": C["grey_light"], "font_color": C["grey"],
            "font_name": "Arial", "font_size": 10,
        })

        for r in range(len(na_df)):
            row_bg = C["offwhite"] if r % 2 == 0 else C["white"]
            ws_na.set_row(row_idx, 15)
            for ci, col_name in enumerate(na_df.columns):
                val     = na_df.iloc[r, ci]
                val_str = "" if pd.isna(val) else str(val)
                base_f  = workbook.add_format({
                    "bg_color": row_bg, "font_name": "Arial", "font_size": 10,
                    "text_wrap": col_name == "Combined_Text",
                    "valign": "top" if col_name == "Combined_Text" else "vcenter",
                })
                if val_str in ("No-Response", "Column-Not-Found", ""):
                    ws_na.write(row_idx, 1 + ci, val_str, F["skip"])
                elif val_str == "Parse-Error":
                    ws_na.write(row_idx, 1 + ci, val_str, F["err"])
                elif col_name == "NA_Has_Constructive":
                    f = {"Yes": na_yes_fmt, "Partial": na_par_fmt}.get(val_str, na_no_fmt)
                    ws_na.write(row_idx, 1 + ci, val_str, f)
                elif col_name == "NA_Regret":
                    f = {"Yes": na_yes_fmt, "Partial": na_par_fmt}.get(val_str, na_no_fmt)
                    ws_na.write(row_idx, 1 + ci, val_str, f)
                elif col_name == "NA_Reason" and val_str in na_reason_colours:
                    bg, fg = na_reason_colours[val_str]
                    reason_f = workbook.add_format({
                        "bg_color": bg, "font_color": fg,
                        "font_name": "Arial", "font_size": 10,
                    })
                    ws_na.write(row_idx, 1 + ci, val_str, reason_f)
                else:
                    ws_na.write(row_idx, 1 + ci, val_str, base_f)
            row_idx += 1

        # Freeze header row (relative to sheet start)
        ws_na.freeze_panes(row_idx - len(na_df), 2)
    else:
        ws_na.write(row_idx, 1, "No non-attender data to display.", F["skip"])

    # ────────────────────────────────────────────────────────────────
    #  🎯 AGENCY IN LEARNING  (dedicated sheet)
    # ────────────────────────────────────────────────────────────────
    if agency_df is not None and not agency_df.empty:
        ws_ag = workbook.add_worksheet("🎯 Agency in Learning")
        writer.sheets["🎯 Agency in Learning"] = ws_ag
        ws_ag.hide_gridlines(2)
        ws_ag.set_column("A:A", 3)
        ws_ag.set_column("B:B", 20)
        ws_ag.set_column("C:C", 14)
        ws_ag.set_column("D:D", 24)
        ws_ag.set_column("E:E", 14)
        ws_ag.set_column("F:F", 24)
        ws_ag.set_column("G:G", 14)
        ws_ag.set_column("H:H", 16)
        ws_ag.set_column("I:I", 40)

        ag_title_fmt = workbook.add_format({
            "bold": True, "font_size": 18, "font_name": "Arial",
            "bg_color": "#4338CA", "font_color": C["white"], "align": "center", "valign": "vcenter",
        })
        ws_ag.merge_range("B1:I2", "🎯  Agency in Learning", ag_title_fmt)
        ws_ag.set_row(0, 30); ws_ag.set_row(1, 30)
        ag_sub_fmt = workbook.add_format({
            "font_size": 10, "font_name": "Arial", "bg_color": "#6D5BD0",
            "font_color": C["white"], "align": "center", "valign": "vcenter", "italic": True,
        })
        ws_ag.merge_range("B3:I3",
            f"{AGENCY_DEFINITION}  |  Scored 1–10 to align with camp survey data "
            "(each level spans 2 pts: L1=1–2, L2=3–4, L3=5–6, L4=7–8, L5=9–10)",
            ag_sub_fmt)
        ws_ag.set_row(2, 18)

        ag_sec_fmt = workbook.add_format({
            "bold": True, "font_size": 11, "font_name": "Arial",
            "bg_color": "#4338CA", "font_color": C["white"],
        })
        ag_note_fmt = workbook.add_format({
            "italic": True, "font_size": 9, "font_name": "Arial", "font_color": C["grey"],
            "text_wrap": True, "valign": "top",
        })
        ag_field_fmt = workbook.add_format({
            "bold": True, "font_size": 10, "font_name": "Arial",
            "bg_color": "#E0E7FF", "font_color": "#4338CA",
        })

        r = 4
        ws_ag.set_row(r, 18)
        ws_ag.merge_range(r, 1, r, 8, "  ⚠️  Reading this data: Year 2026 = this Y8 camp. Year 2025 = the "
                                       "same cohort's Y7 Maria Camp, a different camp with a different rater "
                                       "and ~1 year apart \u2014 treat any year-on-year comparison as indicative, "
                                       "not definitive. The \u201cClass\u201d field in the source export is not "
                                       "reliable and is not used for matching (students are matched by email).",
                          ag_note_fmt)
        ws_ag.set_row(r, 32)
        r += 2

        # Level reference table
        ws_ag.merge_range(r, 1, r, 8, "  📖  The Five Levels of Agency in Learning", ag_sec_fmt)
        r += 1
        for lvl in range(1, 6):
            ws_ag.set_row(r, 34)
            ws_ag.write(r, 1, f"Level {lvl}  (score {lvl*2-1}–{lvl*2}/10)", ag_field_fmt)
            lvl_bg = workbook.add_format({"text_wrap": True, "valign": "vcenter", "font_name": "Arial",
                                          "font_size": 10, "bg_color": "#F5F4FF" if lvl % 2 else C["white"]})
            ws_ag.merge_range(r, 2, r, 8, AGENCY_LEVEL_DESCRIPTIONS[lvl], lvl_bg)
            r += 1
        r += 1

        # Snapshot stat cards
        ws_ag.merge_range(r, 1, r, 8, "  📊  2026 Camp Snapshot", ag_sec_fmt)
        r += 1
        s = agency_summary or {}
        ws_ag.set_row(r, 22); ws_ag.set_row(r+1, 32); ws_ag.set_row(r+2, 18)
        write_stat_card(ws_ag, r, 1, "Students Matched", str(s.get("n_2026", 0)), "Agency score, 2026 camp",
                        bg="#E0E7FF", font_col="#4338CA")
        write_stat_card(ws_ag, r, 4, "Avg Score (2026)", str(s.get("avg_2026", "—")), "out of 10",
                        bg="#E0E7FF", font_col="#4338CA")
        grow_n = s.get("n_both_years", 0)
        grow_v = s.get("avg_growth", None)
        write_stat_card(ws_ag, r, 7, "Avg Growth (Y7→Y8)",
                        f"{grow_v:+.2f}" if grow_v is not None else "—",
                        f"n={grow_n}, see caution above", bg="#FEF3C7", font_col=C["amber"])
        r += 4

        # Per-student table
        ws_ag.merge_range(r, 1, r, 8, "  📋  Student-Level Detail", ag_sec_fmt)
        r += 1
        hdr_fmt = workbook.add_format({
            "bold": True, "bg_color": "#E0E7FF", "font_color": "#4338CA",
            "font_name": "Arial", "font_size": 10, "border": 1, "text_wrap": True, "align": "center",
        })
        headers = ["Name", "Class", "2026 Level", "2026 Confidence", "2025 Level (Y7 Maria Camp)",
                  "2025 Confidence", "Growth (Y7→Y8)", "Comparison Reliability"]
        for ci, h in enumerate(headers):
            ws_ag.write(r, 1 + ci, h, hdr_fmt)
        r += 1

        conf_fmts = {
            "High":           workbook.add_format({"bg_color": C["green_light"], "font_color": C["green"], "bold": True, "font_name": "Arial", "font_size": 10, "align": "center"}),
            "Fuzzy \u2013 Review": workbook.add_format({"bg_color": C["amber_light"], "font_color": C["amber"], "font_name": "Arial", "font_size": 10, "align": "center"}),
            "Unmatched":      workbook.add_format({"bg_color": C["red_light"], "font_color": C["red"], "font_name": "Arial", "font_size": 10, "align": "center"}),
        }
        blank_fmt = workbook.add_format({"font_color": "#BBBBBB", "italic": True, "font_name": "Arial", "font_size": 10, "align": "center"})

        ag_sorted = agency_df.sort_values(
            by="Agency Score (Y8 Camp 2026)" if "Agency Score (Y8 Camp 2026)" in agency_df.columns else "Email",
            ascending=False, na_position="last")
        for i, (_, row_d) in enumerate(ag_sorted.iterrows()):
            row_bg = C["offwhite"] if i % 2 == 0 else C["white"]
            base_f = workbook.add_format({"bg_color": row_bg, "font_name": "Arial", "font_size": 10})
            ctr_f  = workbook.add_format({"bg_color": row_bg, "font_name": "Arial", "font_size": 10, "align": "center"})
            name = f"{row_d.get('First Name','')} {row_d.get('Surname','')}".strip()
            ws_ag.write(r, 1, name or "—", base_f)
            ws_ag.write(r, 2, str(row_d.get("Class", "—")), ctr_f)
            ws_ag.write(r, 3, str(row_d.get("Current Level (Y8 Camp 2026)", "—")) if pd.notna(row_d.get("Current Level (Y8 Camp 2026)")) else "—", ctr_f)
            c1 = row_d.get("Confidence Flag (Y8 Camp 2026)")
            ws_ag.write(r, 4, c1 if pd.notna(c1) else "No data", conf_fmts.get(c1, blank_fmt))
            ws_ag.write(r, 5, str(row_d.get("Current Level (Y7 Maria Camp 2025)", "—")) if pd.notna(row_d.get("Current Level (Y7 Maria Camp 2025)")) else "—", ctr_f)
            c2 = row_d.get("Confidence Flag (Y7 Maria Camp 2025)")
            ws_ag.write(r, 6, c2 if pd.notna(c2) else "No data", conf_fmts.get(c2, blank_fmt))
            growth_v = row_d.get("Agency Growth (Y7\u2192Y8)")
            if pd.notna(growth_v):
                gf = F["g_dec"] if growth_v > 0.05 else (F["r_dec"] if growth_v < -0.05 else F["z_dec"])
                ws_ag.write(r, 7, growth_v, gf)
            else:
                ws_ag.write(r, 7, "—", blank_fmt)
            ws_ag.write(r, 8, str(row_d.get("Comparison Reliability", "")), workbook.add_format(
                {"bg_color": row_bg, "font_name": "Arial", "font_size": 9, "italic": True, "text_wrap": True}))
            r += 1

        ws_ag.freeze_panes(r - len(ag_sorted), 2)

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
    return '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>'


def write_html_report(html_path, tab1_df, avg_df, dist_df, breakdown_df,
                      coded_df, narratives, metrics_processed, mm_tables, long_df,
                      post_only_scores=None, na_df=None, na_narratives=None,
                      agency_df=None, agency_summary=None):
    """Generate a fully self-contained, offline HTML presentation report."""
    import json as _json

    year = datetime.now().year

    def safe_float(v):
        try:
            f = float(str(v).replace("+","").strip())
            return round(f, 2) if f == f else 0
        except Exception:
            return 0

    # ── Real stat card percentages from coded_df
    def pct_val(field, positive_vals=("Positive",)):
        if field not in coded_df.columns:
            return 0
        vals = coded_df[field]
        valid = vals[~vals.isin(["No-Response","Column-Not-Found","Parse-Error",""])]
        return round(len(valid[valid.isin(positive_vals)]) / len(valid) * 100) if len(valid) else 0

    peer_pct   = pct_val("Q1_Sentiment")
    fav_pct    = pct_val("Q5_Sentiment")
    growth_pct = 0
    for gf in ["Q2_Growth","Q8_Growth"]:
        if gf in coded_df.columns:
            v = coded_df[gf]
            v = v[~v.isin(["No-Response","Column-Not-Found","Parse-Error",""])]
            if len(v):
                growth_pct = round(len(v[v.isin(["Yes","Partial"])]) / len(v) * 100)
                break

    # ── Category summary for overview cards
    cat_summary_cards = []
    for cat in CATEGORY_ORDER:
        cat_rows = avg_df[avg_df["Metric"].isin(METRIC_CATEGORIES.get(cat, []))] if not avg_df.empty else pd.DataFrame()
        if cat_rows.empty:
            continue
        cat_avg_shift = float(cat_rows["Avg Shift"].mean())
        cat_pct_imp   = float(cat_rows["% Improvers"].mean()) if "% Improvers" in cat_rows.columns else 0
        cat_colour_cls = {"Activity Skills": "teal", "Camping Skills": "green",
                          "Attitudes": "amber", "Knowledge": "purple",
                          "Agency in Learning": "indigo"}.get(cat, "")
        cat_summary_cards.append({
            "cat": cat, "icon": CATEGORY_ICONS.get(cat, ""),
            "avg_shift": cat_avg_shift, "pct_imp": cat_pct_imp,
            "colour_cls": cat_colour_cls,
        })

    # ── Metric chart data
    metric_labels, pre_vals, post_vals, shift_vals, median_shift_vals, pct_improvers_vals = [], [], [], [], [], []
    if not avg_df.empty and "Metric" in avg_df.columns:
        for _, row in avg_df.iterrows():
            metric_labels.append(str(row.get("Metric","")))
            pre_vals.append(safe_float(row.get("Pre-Camp Avg", 0)))
            post_vals.append(safe_float(row.get("Post-Camp Avg", 0)))
            shift_vals.append(safe_float(row.get("Avg Shift", 0)))
            median_shift_vals.append(safe_float(row.get("Median Shift", 0)))
            pct_improvers_vals.append(safe_float(row.get("% Improvers", 0)))

    # ── Distribution chart data
    dist_labels, improved_vals, same_vals, declined_vals = [], [], [], []
    if not dist_df.empty:
        for _, row in dist_df.iterrows():
            dist_labels.append(str(row.get("Metric","")))
            improved_vals.append(round(safe_float(row.get("Improved",0))*100, 1))
            same_vals.append(round(safe_float(row.get("Stayed Same",0))*100, 1))
            declined_vals.append(round(safe_float(row.get("Declined",0))*100, 1))

    # ── Class chart data
    class_labels, class_shift_vals = [], []
    if not breakdown_df.empty:
        for _, row in breakdown_df.iterrows():
            cls = str(row.get("Class",""))
            if not cls.startswith("  ↳") and cls not in ("","nan"):
                class_labels.append(cls.strip())
                class_shift_vals.append(safe_float(row.get("Average Shift (All Areas)",0)))

    # ── Agency in Learning chart + table data
    ag_summary = agency_summary or {}
    # Level labels show both the level number and its 1-10 score band
    ag_level_labels = [f"L{i} ({i*2-1}–{i*2})" for i in range(1, 6)]
    ag_level_dist   = ag_summary.get("level_dist", {})
    ag_level_counts = []
    for i in range(1, 6):
        cnt = 0
        for k, v in ag_level_dist.items():
            if str(k).strip().endswith(str(i)):
                cnt += v
        ag_level_counts.append(int(cnt))

    ag_growth_labels, ag_growth_vals = [], []
    ag_table_rows_html = ""
    if agency_df is not None and not agency_df.empty:
        ag_sorted = agency_df.sort_values(
            by="Agency Score (Y8 Camp 2026)" if "Agency Score (Y8 Camp 2026)" in agency_df.columns else "Email",
            ascending=False, na_position="last")
        ag_conf_cls = {"High": "pos", "Fuzzy \u2013 Review": "zero", "Unmatched": "neg"}
        for _, rrow in ag_sorted.iterrows():
            name = f"{rrow.get('First Name','')} {rrow.get('Surname','')}".strip() or "\u2014"
            lvl26 = rrow.get("Current Level (Y8 Camp 2026)")
            lvl26 = lvl26 if pd.notna(lvl26) else "\u2014"
            lvl25 = rrow.get("Current Level (Y7 Maria Camp 2025)")
            lvl25 = lvl25 if pd.notna(lvl25) else "\u2014"
            conf26 = rrow.get("Confidence Flag (Y8 Camp 2026)")
            conf26 = conf26 if pd.notna(conf26) else "No data"
            growth = rrow.get("Agency Growth (Y7\u2192Y8)")
            if pd.notna(growth):
                ag_growth_labels.append(name)
                ag_growth_vals.append(round(float(growth), 2))
                g_str   = f"{growth:+.2f}"
                g_class = "pos" if growth > 0.05 else ("neg" if growth < -0.05 else "zero")
            else:
                g_str, g_class = "\u2014", "zero"
            conf_class = ag_conf_cls.get(conf26, "zero")
            reliability = str(rrow.get("Comparison Reliability", ""))
            ag_table_rows_html += (
                f"<tr><td>{name}</td><td style='text-align:center'>{lvl26}</td>"
                f"<td style='text-align:center' class='{conf_class}'>{conf26}</td>"
                f"<td style='text-align:center'>{lvl25}</td>"
                f"<td style='text-align:center' class='{g_class}'>{g_str}</td>"
                f"<td style='font-size:11px;color:#6B7280'>{reliability}</td></tr>"
            ).replace("'", '"')

    ag_level_desc_html = "".join(
        f"<div class='ag-level-row'>"
        f"<span class='ag-level-badge'>Level {i}<br><span style='font-weight:400;font-size:10px'>"
        f"score {i*2-1}\u2013{i*2}</span></span>"
        f"<span class='ag-level-text'>{AGENCY_LEVEL_DESCRIPTIONS[i]}</span></div>"
        for i in range(1, 6)
    )

    agency_section_html = ""
    if agency_df is not None and not agency_df.empty:
        n_2026 = ag_summary.get("n_2026", 0)
        avg_2026 = ag_summary.get("avg_2026")
        n_both = ag_summary.get("n_both_years", 0)
        avg_growth = ag_summary.get("avg_growth")
        pct_grew = ag_summary.get("pct_grew")
        agency_section_html = (
            "<h2 class=\"section-title\" id=\"agency\" style=\"border-color:#4338CA;color:#4338CA\">"
            "&#127919; Agency in Learning</h2>"
            f"<p style='color:#6B7280;margin-bottom:8px;font-size:13px'><em>{AGENCY_DEFINITION}</em></p>"
            "<p style='color:#6B7280;margin-bottom:12px;font-size:12.5px'>"
            "Scored on a <strong>1–10 scale</strong> to align with other camp survey metrics. "
            "Each of the 5 descriptor levels spans 2 points: Level&nbsp;1&nbsp;=&nbsp;1\u20132, "
            "Level&nbsp;2&nbsp;=&nbsp;3\u20134, Level&nbsp;3&nbsp;=&nbsp;5\u20136, "
            "Level&nbsp;4&nbsp;=&nbsp;7\u20138, Level&nbsp;5&nbsp;=&nbsp;9\u201310. "
            "The full level descriptors are shown in the table below.</p>"
            "<div class='ag-warning'>&#9888;&#65039; <strong>Reading this data:</strong> "
            "2026 = this Y8 camp cohort. 2025 = the same students&rsquo; Y7 Maria Camp baseline &mdash; "
            "a different camp, different rater, and roughly a year apart. Year-on-year comparisons are "
            "shown for interest but should be treated as indicative, not definitive. Students are matched "
            "by email; the source data&rsquo;s &ldquo;Class&rdquo; field is not used because it is unreliable.</div>"
            "<div class='ag-grid'>"
            "<div class='ag-levels-card'><h4>The Five Levels</h4>" + ag_level_desc_html + "</div>"
            "<div class='ag-stats-col'>"
            "<div class='cards' style='margin-bottom:16px'>"
            f"<div class='card indigo'><div class='card-val'>{n_2026}</div>"
            "<div class='card-lbl'>Students matched (2026)</div></div>"
            f"<div class='card indigo'><div class='card-val'>{avg_2026 if avg_2026 is not None else '\u2014'}</div>"
            "<div class='card-lbl'>Avg score (2026), out of 10</div></div>"
            f"<div class='card amber'><div class='card-val'>{(f'{avg_growth:+.2f}' if avg_growth is not None else '\u2014')}</div>"
            f"<div class='card-lbl'>Avg growth Y7&rarr;Y8 (n={n_both})</div>"
            f"<div class='card-sub'>{(f'{pct_grew:.0f}% grew' if pct_grew is not None else 'No baseline overlap')}</div></div>"
            "</div>"
            "<div class='chart-box' style='height:260px'><h3>2026 Level Distribution</h3>"
            "<canvas id='agency-level-chart' style='height:200px;width:100%;display:block'></canvas></div>"
            "</div></div>"
            + (f"<div class='chart-box' style='height:{min(460, max(280, len(ag_growth_vals)*24+60))}px;margin-top:16px'>"
               "<h3>Growth, Y7 Maria Camp &rarr; Y8 Camp (students with both data points)</h3>"
               f"<div style='overflow-y:auto;height:{min(400, max(220, len(ag_growth_vals)*24))}px'>"
               f"<canvas id='agency-growth-chart' style='height:{max(220, len(ag_growth_vals)*24)}px;width:100%;display:block'></canvas></div></div>"
               if ag_growth_vals else
               "<p style='color:#9CA3AF;font-size:12px;margin-top:8px'>No students currently have both a 2025 and 2026 score to compare.</p>")

            + "<h3 style='margin-top:20px;font-size:15px;color:#1B3A5C'>Student-Level Detail</h3>"
            "<div class='breakdown-wrap'><table class='breakdown-table' style='white-space:normal'>"
            "<thead><tr><th>Name</th><th>2026 Level</th><th>2026 Match</th>"
            "<th>2025 Level (Y7)</th><th>Growth</th><th>Comparison Reliability</th></tr></thead>"
            f"<tbody>{ag_table_rows_html}</tbody></table></div>"
        )

    # ── Qual chart data (counts per coded value, per field)
    qual_charts_js = {}
    for qc in QUAL_QUESTIONS:
        for field in qc["fields"]:
            if field not in coded_df.columns:
                continue
            counts = coded_df[field].value_counts()
            counts = counts[~counts.index.isin(["No-Response","Column-Not-Found","Parse-Error",""])]
            if len(counts) == 0:
                continue
            qual_charts_js[field] = {
                "labels": [k.replace("-"," ") for k in counts.index.tolist()],
                "data":   counts.values.tolist(),
            }

    # ── Clean mm_tables: remove spurious header rows
    clean_mm = []
    for df in mm_tables:
        if df.empty:
            continue
        df2 = df[df.get("Analysis","") != "Analysis"] if "Analysis" in df.columns else df
        df2 = df2.dropna(subset=["Analysis"]) if "Analysis" in df2.columns else df2
        df2 = df2[df2["Analysis"].astype(str).str.strip() != ""] if "Analysis" in df2.columns else df2
        if not df2.empty:
            clean_mm.append(df2)

    # ── Mixed methods HTML blocks
    mm_html = ""
    for df in clean_mm:
        title = str(df["Analysis"].iloc[0]) if "Analysis" in df.columns else "Analysis"
        cols  = [c for c in df.columns if c != "Analysis"]
        rows_html = ""
        for _, row in df.iterrows():
            cells = ""
            for col in cols:
                val = row[col]
                cls_attr = ""
                if col == "Avg Shift":
                    try:
                        fv = float(str(val).replace("+","").replace("−","-"))
                        cls_attr = " class=\'pos\'" if fv > 0 else " class=\'neg\'" if fv < 0 else ""
                        val = f"{fv:+.2f}" if fv != 0 else "0"
                    except Exception:
                        pass
                cells += f"<td{cls_attr}>{val}</td>"
            rows_html += f"<tr>{cells}</tr>"
        headers = "".join(f"<th>{c}</th>" for c in cols)
        mm_html += (
            f"<div class=\'mm-block\'><h4 class=\'mm-title\'>{title}</h4>"
            f"<table class=\'mm-table\'><thead><tr>{headers}</tr></thead>"
            f"<tbody>{rows_html}</tbody></table></div>"
        )
    mm_html = mm_html.replace("\'", '"')

    # ── Parse narrative: split on *** to separate analysis from quotes
    def parse_narrative(text):
        """Split AI narrative into (analysis_text, [quotes]).

        Detects quotes by looking for lines that start with a bullet character
        followed by a quotation mark, or a contextual bracket tag then a quote.
        Falls back to the legacy *** separator if present.
        Quote lines are stripped from the analysis block so they don't appear twice.
        """
        if not text or str(text).strip() in ("","nan"):
            return "", []

        text = str(text).replace("\\n", "\n").strip()

        # ── Legacy separator support
        if "***" in text:
            parts = text.split("***", 1)
            analysis_raw = parts[0].strip()
            quote_block  = parts[1]
        else:
            analysis_raw = text
            quote_block  = None

        # ── Detect quote lines anywhere in the text
        # A quote line is one whose meaningful content (after stripping bullets/spaces)
        # starts with a quotation mark or an optional [context tag] then a quotation mark.
        QUOTE_OPENERS = ('"', '\u201c', '\u2018', "'")
        BULLET_CHARS  = ("*", "•", "-", "·", "–")

        def looks_like_quote(line):
            s = line.strip()
            # Strip leading bullet + space
            for b in BULLET_CHARS:
                if s.startswith(b):
                    s = s[len(b):].strip()
                    break
            # Allow optional [context tag]
            if s.startswith("["):
                end = s.find("]")
                if end != -1:
                    s = s[end + 1:].strip()
            return s.startswith(QUOTE_OPENERS) and len(s) > 12

        lines          = text.split("\n")
        analysis_lines = []
        raw_quote_lines = []

        if quote_block is not None:
            # Use the explicit separator split
            analysis_lines = analysis_raw.split("\n")
            raw_quote_lines = quote_block.split("\n")
        else:
            # Auto-detect: once we see the first quote line, everything after is quotes
            in_quotes = False
            for line in lines:
                if not in_quotes and looks_like_quote(line):
                    in_quotes = True
                if in_quotes:
                    raw_quote_lines.append(line)
                else:
                    analysis_lines.append(line)

        # Clean the analysis block
        analysis = "\n".join(analysis_lines).strip()

        # Extract clean quote strings
        quotes = []
        for line in raw_quote_lines:
            s = line.strip()
            if not s:
                continue
            # Strip leading bullet
            for b in BULLET_CHARS:
                if s.startswith(b):
                    s = s[len(b):].strip()
                    break
            # Preserve [context tag] if present
            context_tag = ""
            if s.startswith("["):
                end = s.find("]")
                if end != -1:
                    context_tag = s[:end + 1] + " "
                    s = s[end + 1:].strip()
            # Strip outer quote marks
            for q in ('\u201c', '\u201d', '"'): s = s.replace(q, '"')
            s = s.strip('"').strip("'").strip()
            if len(s) > 8:
                quotes.append(context_tag + s)

        return analysis, quotes

    # ── Pre-compute Aboriginal Culture metric index (needed in qual loop for Q4)
    abcult_idx = next((i for i, m in enumerate(metric_labels) if m == "Aboriginal Culture"), None)
    abcult_stats_html = ""
    if abcult_idx is not None:
        ac_pre   = pre_vals[abcult_idx]
        ac_post  = post_vals[abcult_idx]
        ac_shift = shift_vals[abcult_idx]
        ac_pct   = pct_improvers_vals[abcult_idx] if abcult_idx < len(pct_improvers_vals) else 0
        ac_arrow = "▲" if ac_shift > 0.05 else ("▼" if ac_shift < -0.05 else "—")
        ac_cls   = "pos" if ac_shift > 0.05 else ("neg" if ac_shift < -0.05 else "zero")
        abcult_stats_html = (
            f'<div class="abcult-stats">'
            f'<div class="abcult-stat"><div class="abcult-val">{ac_pre:.1f}</div>'
            f'<div class="abcult-lbl">Pre-camp avg</div></div>'
            f'<div class="abcult-stat"><div class="abcult-val">{ac_post:.1f}</div>'
            f'<div class="abcult-lbl">Post-camp avg</div></div>'
            f'<div class="abcult-stat"><div class="abcult-val {ac_cls}">{ac_arrow} {ac_shift:+.2f}</div>'
            f'<div class="abcult-lbl">Avg shift</div></div>'
            f'<div class="abcult-stat"><div class="abcult-val">{ac_pct:.0f}%</div>'
            f'<div class="abcult-lbl">Students improved</div></div>'
            f'</div>'
        )

    # ── Filter Aboriginal Culture from main metric chart data (shown in Q4 instead)
    chart_filter  = [i for i, m in enumerate(metric_labels) if m != "Aboriginal Culture"]
    chart_labels  = [metric_labels[i] for i in chart_filter]
    chart_pre     = [pre_vals[i]       for i in chart_filter]
    chart_post    = [post_vals[i]      for i in chart_filter]
    chart_shift   = [shift_vals[i]     for i in chart_filter]
    chart_dist_lbl  = [dist_labels[i]   for i in chart_filter]
    chart_improved  = [improved_vals[i] for i in chart_filter]
    chart_same      = [same_vals[i]     for i in chart_filter]
    chart_declined  = [declined_vals[i] for i in chart_filter]
    chart_shift_cols = [("#2D7A4F" if v > 0.05 else "#C0392B" if v < -0.05 else "#9CA3AF")
                        for v in chart_shift]
    # Dynamic canvas height: 28px per metric, hard-capped at 480px
    chart_h = min(480, max(240, len(chart_labels) * 28))

    # ── Build qual sections HTML
    qual_section_html = ""
    for nar in narratives:
        q_label  = nar.get("Question","")
        n_val    = nar.get("n", 0)
        summary  = nar.get("Summary","")
        analysis, quotes = parse_narrative(summary)

        m = re.match(r"Q(\d+)", q_label)
        qnum = int(m.group(1)) if m else 0
        qlabel_display = re.sub(r"^Q\d+:\s*","", q_label)

        qc = next((q for q in QUAL_QUESTIONS if q["num"] == qnum), None)
        fields = list(qc["fields"].keys()) if qc else []

        # Field charts sidebar
        field_charts = ""
        for field in fields:
            if field in qual_charts_js:
                cid = "chart-" + field.replace("_","-").lower()
                field_charts += (
                    "<div class=\'field-chart-wrap\'>"
                    "<h5 class=\"field-chart-title\">" + field.replace("_"," ") + "</h5>"
                    f"<canvas id=\'{cid}\' style=\'height:130px;width:100%;display:block\'></canvas>"
                    "</div>"
                )
        field_charts = field_charts.replace("\'", '"')

        # Analysis paragraphs
        analysis_html = "".join(
            f"<p>{para.strip()}</p>"
            for para in analysis.split("\n\n") if para.strip()
        )

        # Student quotes
        quotes_html = ""
        if quotes:
            items = "".join(
                f"<blockquote class=\'quote-item\'>&ldquo;{q}&rdquo;</blockquote>"
                for q in quotes[:5]
            ).replace("\'",'"')
            quotes_html = f"<div class=\'quotes-block\'>{items}</div>".replace("\'",'"')

        no_chart = "" if field_charts else "<p class=\'no-chart\'>Chart data populates after AI coding run.</p>".replace("\'",'"')

        # ── Q4: inject Aboriginal Culture quantitative stats alongside qualitative findings
        abcult_inject = ""
        if qnum == 4 and abcult_stats_html:
            abcult_inject = (
                f'<div class="abcult-banner">'
                f'<div class="abcult-title">&#128218; Aboriginal Culture — Quantitative Shift</div>'
                f'{abcult_stats_html}'
                f'</div>'
            )

        qual_section_html += f"""
        <div class="qual-section">
            <div class="qual-header">
                <span class="q-num">Q{qnum}</span>
                <span class="q-label">{qlabel_display}</span>
                <span class="q-n">n&nbsp;=&nbsp;{n_val}</span>
            </div>
            {abcult_inject}
            <div class="qual-body">
                <div class="qual-left">
                    <div class="qual-analysis">{analysis_html}</div>
                    {quotes_html}
                </div>
                <div class="qual-charts">{field_charts}{no_chart}</div>
            </div>
        </div>"""

    # ── Build non-attender section HTML
    na_section_html = ""
    n_na_students = len(na_df) if na_df is not None and not na_df.empty else 0
    if na_narratives:
        for nar in na_narratives:
            nar_label   = nar.get("Question", "Non-Attender Analysis")
            nar_n       = nar.get("n", 0)
            nar_summary = nar.get("Summary", "")
            nar_analysis, nar_quotes = parse_narrative(nar_summary)
            nar_analysis_html = "".join(
                f"<p>{para.strip()}</p>"
                for para in nar_analysis.split("\n\n") if para.strip()
            )
            nar_quotes_html = ""
            if nar_quotes:
                q_items = "".join(
                    f'<blockquote class="quote-item na-quote">&ldquo;{q}&rdquo;</blockquote>'
                    for q in nar_quotes[:5]
                )
                nar_quotes_html = f'<div class="quotes-block">{q_items}</div>'
            na_section_html += f"""
        <div class="qual-section na-section">
            <div class="qual-header na-header">
                <span class="q-num na-badge">🚫</span>
                <span class="q-label">{nar_label}</span>
                <span class="q-n">n&nbsp;=&nbsp;{nar_n}</span>
            </div>
            <div class="qual-body">
                <div class="qual-left">
                    <div class="qual-analysis">{nar_analysis_html}</div>
                    {nar_quotes_html}
                </div>
                <div class="qual-charts">
                    <p class="no-chart" style="color:#7C3AED;font-style:normal;font-size:12px">
                        Non-attender responses coded for: reasons, regret, sentiment,
                        and constructive feedback for future attendance.
                    </p>
                </div>
            </div>
        </div>"""
    # ── Coded table for non-attenders (summary counts)
    na_table_html = ""
    if na_df is not None and not na_df.empty:
        # Summarise each coded field as a count table
        field_summaries = ""
        for field in ALL_NON_ATTEND_FIELDS:
            if field not in na_df.columns:
                continue
            counts = na_df[field].value_counts()
            counts = counts[~counts.index.isin(["No-Response","Parse-Error",""])]
            if counts.empty:
                continue
            rows_h = "".join(
                f"<tr><td>{k.replace('-',' ')}</td>"
                f"<td style='text-align:center;font-weight:700'>{v}</td>"
                f"<td style='text-align:center;color:#6B7280'>"
                f"{v/len(na_df)*100:.0f}%</td></tr>"
                for k, v in counts.items()
            )
            field_summaries += (
                f"<div class='na-count-block'>"
                f"<h5 class='na-count-title'>{field.replace('_',' ')}</h5>"
                f"<table class='mm-table'>"
                f"<thead><tr><th>Category</th><th>Count</th><th>%</th></tr></thead>"
                f"<tbody>{rows_h}</tbody></table></div>"
            )
        if field_summaries:
            na_table_html = (
                f'<div class="na-count-grid">{field_summaries}</div>'
            )

    # ── Grouped metric table: Activity Skills, Camping Skills, Attitudes only
    # Knowledge (Aboriginal Culture) moves to the Q4 section below.
    metric_rows_html = ""
    for cat in CATEGORY_ORDER:
        if cat == "Knowledge":
            continue   # Aboriginal Culture lives with Q4 First Nations section
        cat_metrics_in_avg = [(i, row) for i, row in enumerate(metric_labels)
                              if row in METRIC_CATEGORIES.get(cat, [])]
        if not cat_metrics_in_avg:
            continue
        cat_colour_hex = CATEGORY_COLOURS.get(cat, "#2E8B88")
        cat_light_hex  = CATEGORY_LIGHT_COLOURS.get(cat, "#D4F0EF")
        cat_icon       = CATEGORY_ICONS.get(cat, "")
        # ── Category header: background on the <td> so CSS tr:nth-child can't override it
        metric_rows_html += (
            f'<tr>'
            f'<td colspan="7" style="background:{cat_colour_hex};color:#ffffff;'
            f'font-weight:700;font-size:13px;padding:10px 16px;'
            f'letter-spacing:.04em;border-bottom:2px solid rgba(0,0,0,.1)">'
            f'{cat_icon}&nbsp;&nbsp;{cat}</td></tr>'
        )
        for i, m in cat_metrics_in_avg:
            pre  = pre_vals[i]
            post = post_vals[i]
            sh   = shift_vals[i]
            med  = median_shift_vals[i] if i < len(median_shift_vals) else 0
            pct  = pct_improvers_vals[i] if i < len(pct_improvers_vals) else 0
            arrow     = "▲" if sh > 0.05 else ("▼" if sh < -0.05 else "—")
            cls       = "pos" if sh > 0.05 else ("neg" if sh < -0.05 else "zero")
            med_arrow = "▲" if med > 0.05 else ("▼" if med < -0.05 else "—")
            med_cls   = "pos" if med > 0.05 else ("neg" if med < -0.05 else "zero")
            # % Improvers: colour-coded text only, row background handles cell bg
            pct_color = "#2D7A4F" if pct >= 55 else ("#C0392B" if pct < 35 else "#D97706")
            bar_w     = min(abs(sh) / 3.0 * 100, 100)
            # Apply category light colour directly to every <td> so nth-child can't override
            bg = f"background:{cat_light_hex};"
            metric_rows_html += (
                f'<tr>'
                f'<td class="metric-name" style="{bg}padding-left:28px">{m}</td>'
                f'<td class="num-cell" style="{bg}">{pre:.1f}</td>'
                f'<td class="num-cell" style="{bg}">{post:.1f}</td>'
                f'<td class="shift-cell {cls}" style="{bg}">{arrow} {sh:+.2f}</td>'
                f'<td class="shift-cell {med_cls}" style="{bg}">{med_arrow} {med:+.2f}</td>'
                f'<td class="num-cell" style="{bg}font-weight:700;color:{pct_color}">{pct:.1f}%</td>'
                f'<td class="bar-cell" style="{bg}">'
                f'<div class="shift-bar {cls}" style="width:{bar_w:.0f}%;min-width:2px"></div></td>'
                f'</tr>'
            ).replace("'", '"')

    # ── Category summary stat cards (3 domains, Aboriginal Culture shown separately in Q4)
    cat_cards_html = ""
    for cs in cat_summary_cards:
        if cs["cat"] == "Knowledge":
            continue   # shown in Q4 section
        sh = cs["avg_shift"]
        arrow = "▲" if sh > 0.05 else ("▼" if sh < -0.05 else "—")
        sign_cls = "pos" if sh > 0.05 else ("neg" if sh < -0.05 else "zero")
        cat_cards_html += (
            f'<div class="card {cs["colour_cls"]}">'
            f'<div class="card-val"><span class="{sign_cls}">{arrow} {sh:+.2f}</span></div>'
            f'<div class="card-lbl">{cs["icon"]} {cs["cat"]} &mdash; avg shift</div>'
            f'<div class="card-sub">{cs["pct_imp"]:.0f}% of students improved</div>'
            f'</div>'
        )

    # ── Class breakdown rows
    class_rows_html = ""
    class_header_cells = "<th>Class</th><th>Avg Shift</th>" + "".join(f"<th>{m}</th>" for m in metrics_processed)
    if not breakdown_df.empty:
        for _, row in breakdown_df.iterrows():
            cn  = str(row.get("Class",""))
            sub = cn.startswith("  ↳")
            rc  = "sub-row" if sub else "class-row"
            av  = safe_float(row.get("Average Shift (All Areas)", 0))
            sc  = "pos" if av > 0 else ("neg" if av < 0 else "zero")
            cells = f"<td>{cn.strip()}</td><td class=\'num-cell {sc}\'>{av:+.2f}</td>"
            for m in metrics_processed:
                v  = safe_float(row.get(m, 0))
                mc = "pos" if v > 0 else ("neg" if v < 0 else "zero")
                cells += f"<td class=\'num-cell {mc}\'>{v:+.1f}</td>"
            cells = cells.replace("\'",'"')
            class_rows_html += f"<tr class=\'{rc}\'>{cells}</tr>".replace("\'",'"')

    # ── Counts for hero section
    n_students    = len(tab1_df)
    n_improved    = sum(1 for v in shift_vals if v > 0)
    n_metrics     = len(shift_vals)
    overall_shift = f"{sum(shift_vals)/n_metrics:+.2f}" if n_metrics else "—"

    # ── JS: map each qual field name to its canvas id
    field_canvas_map_js = ""
    for qc in QUAL_QUESTIONS:
        for field in qc["fields"]:
            cid = "chart-" + field.replace("_","-").lower()
            field_canvas_map_js += f"  \"{field}\": \"{cid}\",\n"
    field_canvas_map_js = field_canvas_map_js.replace("\'","'")

    shift_colours = _json.dumps(["#2D7A4F" if v > 0.05 else "#C0392B" if v < -0.05 else "#9CA3AF" for v in shift_vals])
    qual_js       = _json.dumps(qual_charts_js, indent=2)

    # Load Chart.js (embedded offline or CDN fallback)
    chartjs_tag = _load_chartjs()

    # ── Y9 Camp Readiness: teal stat card + post-only metric table row
    po_card_html = ""
    po_metric_row_html = ""
    if post_only_scores:
        for po_name, po_data in post_only_scores.items():
            po_avg = po_data["avg"]
            po_n   = po_data["n"]
            po_val_str = f"{po_avg:.1f}/10"
            po_card_html += (
                f'<div class="card teal">'
                f'<div class="card-val">{po_val_str}</div>'
                f'<div class="card-lbl">{po_name} &mdash; post-camp (n={po_n})</div>'
                f'</div>'
            )
            bar_w = min(po_avg / 10.0 * 100, 100)
            po_metric_row_html += (
                f'<tr class="po-row">'
                f'<td class="metric-name">&#9733; {po_name}</td>'
                f'<td class="num-cell" style="color:#9CA3AF">&#8212;</td>'
                f'<td class="num-cell">{po_avg:.1f}</td>'
                f'<td class="shift-cell" style="color:#2E8B88;font-size:12px;font-weight:600">'
                f'Post-only</td>'
                f'<td class="shift-cell" style="color:#9CA3AF">&#8212;</td>'
                f'<td class="shift-cell" style="color:#9CA3AF">&#8212;</td>'
                f'<td class="bar-cell"><div class="shift-bar" '
                f'style="background:#2E8B88;width:{bar_w:.0f}%;min-width:2px"></div></td>'
                f'</tr>'
            )

    html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Camp Report """ + str(year) + """</title>
""" + chartjs_tag + """
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#F0F4F8;color:#1a2332;font-size:14px;line-height:1.6}
.nav{background:#1B3A5C;color:white;padding:0 32px;display:flex;align-items:center;height:56px;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.2)}
.nav-title{font-size:18px;font-weight:700;flex:1}
.nav-links{display:flex;gap:24px}
.nav-links a{color:rgba(255,255,255,.8);text-decoration:none;font-size:13px;font-weight:500;padding:4px 0;border-bottom:2px solid transparent;transition:all .2s}
.nav-links a:hover{color:white;border-color:#2E8B88}
.page{max-width:1200px;margin:0 auto;padding:32px 24px}
.hero{background:linear-gradient(135deg,#1B3A5C 0%,#2E8B88 100%);color:white;border-radius:16px;padding:40px 48px;margin-bottom:32px;display:flex;align-items:center;gap:48px}
.hero-text h1{font-size:28px;font-weight:800;margin-bottom:8px}
.hero-text p{opacity:.85;font-size:15px}
.hero-stats{display:flex;gap:32px;flex-shrink:0}
.hero-stat{text-align:center}
.hero-stat .big{font-size:42px;font-weight:800;line-height:1}
.hero-stat .lbl{font-size:12px;opacity:.8;text-transform:uppercase;letter-spacing:.05em;margin-top:4px}
.section-title{font-size:20px;font-weight:700;color:#1B3A5C;margin:40px 0 20px;padding-bottom:10px;border-bottom:3px solid #2E8B88;display:flex;align-items:center;gap:10px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:24px}
.card{background:white;border-radius:12px;padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,.08);border-top:4px solid #2E8B88;transition:transform .15s,box-shadow .15s}
.card:hover{transform:translateY(-2px);box-shadow:0 4px 12px rgba(0,0,0,.12)}
.card-val{font-size:36px;font-weight:800;color:#1B3A5C;line-height:1}
.card-lbl{font-size:13px;color:#6B7280;margin-top:4px}
.card.green{border-top-color:#2D7A4F}.card.green .card-val{color:#2D7A4F}
.card.amber{border-top-color:#D97706}.card.amber .card-val{color:#D97706}
.card.purple{border-top-color:#5B21B6}.card.purple .card-val{color:#5B21B6}
.card.indigo{border-top-color:#4338CA}.card.indigo .card-val{color:#4338CA}
.ag-warning{background:#FEF3C7;border-left:4px solid #D97706;border-radius:6px;padding:10px 14px;font-size:12.5px;color:#92400E;margin-bottom:18px;line-height:1.5}
.ag-grid{display:grid;grid-template-columns:1.1fr 1fr;gap:20px;align-items:start}
@media(max-width:900px){.ag-grid{grid-template-columns:1fr}}
.ag-levels-card{background:white;border-radius:12px;padding:18px 20px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.ag-levels-card h4{margin:0 0 10px;color:#4338CA;font-size:13px;text-transform:uppercase;letter-spacing:.04em}
.ag-level-row{display:flex;gap:10px;align-items:flex-start;padding:7px 0;border-bottom:1px solid #F0F0F0}
.ag-level-row:last-child{border-bottom:none}
.ag-level-badge{flex:0 0 auto;background:#E0E7FF;color:#4338CA;font-weight:700;font-size:11px;padding:3px 8px;border-radius:6px}
.ag-level-text{font-size:12.5px;color:#374151;line-height:1.45}
.ag-stats-col{display:flex;flex-direction:column;gap:0}
.card.teal{border-top-color:#2E8B88}.card.teal .card-val{color:#2E8B88}
.po-row td{background:#D4F0EF!important;font-style:italic}
.metric-table{width:100%;border-collapse:collapse;background:white;border-radius:10px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:24px}
.metric-table th{background:#1B3A5C;color:white;padding:12px 16px;text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:.05em}
.metric-table td{padding:10px 16px;border-bottom:1px solid #F0F0F0}
.metric-name{font-weight:600}
.num-cell{text-align:center;font-variant-numeric:tabular-nums}
.shift-cell{text-align:center;font-weight:700;font-size:15px;min-width:80px}
.bar-cell{width:120px}
.shift-bar{height:8px;border-radius:4px}
.pos{color:#2D7A4F}.pos .shift-bar,.shift-bar.pos{background:#2D7A4F}
.neg{color:#C0392B}.neg .shift-bar,.shift-bar.neg{background:#C0392B}
.zero{color:#6B7280}.zero .shift-bar,.shift-bar.zero{background:#D3D3D3}
.amber{color:#D97706}
.chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-bottom:24px}
.chart-box{background:white;border-radius:12px;padding:20px 24px 16px;box-shadow:0 1px 4px rgba(0,0,0,.08);overflow:hidden}
.chart-box h3{font-size:15px;font-weight:700;color:#1B3A5C;margin:0 0 12px}
.chart-box canvas{display:block;width:100%;height:100%}
.abcult-banner{background:linear-gradient(135deg,#EDE9FE,#DDD6FE);border-left:4px solid #5B21B6;padding:16px 24px;margin:0}
.abcult-title{font-size:13px;font-weight:700;color:#5B21B6;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px}
.abcult-stats{display:flex;gap:24px;flex-wrap:wrap}
.abcult-stat{text-align:center;background:white;border-radius:8px;padding:10px 18px;box-shadow:0 1px 3px rgba(91,33,182,.15)}
.abcult-val{font-size:22px;font-weight:800;color:#5B21B6;line-height:1}
.abcult-lbl{font-size:11px;color:#6B7280;margin-top:3px;text-transform:uppercase;letter-spacing:.04em}
.breakdown-wrap{overflow-x:auto;margin-bottom:24px}
.breakdown-table{width:100%;border-collapse:collapse;background:white;border-radius:10px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08);font-size:12px;white-space:nowrap}
.breakdown-table th{background:#1B3A5C;color:white;padding:9px 12px;text-align:center;font-size:11px;text-transform:uppercase;letter-spacing:.04em}
.breakdown-table td{padding:8px 12px;border-bottom:1px solid #F0F0F0;text-align:center}
.class-row td{background:#EFF6FF;font-weight:700;color:#1B3A5C;text-align:left}
.sub-row td:first-child{text-align:left;color:#374151;padding-left:20px}
.qual-section{background:white;border-radius:12px;margin-bottom:24px;box-shadow:0 1px 4px rgba(0,0,0,.08);overflow:hidden}
.qual-header{background:#1B3A5C;color:white;padding:16px 24px;display:flex;align-items:center;gap:16px}
.q-num{background:#2E8B88;color:white;border-radius:6px;padding:4px 12px;font-weight:800;font-size:15px;flex-shrink:0}
.q-label{font-size:17px;font-weight:700;flex:1}
.q-n{opacity:.75;font-size:13px;flex-shrink:0}
.qual-body{padding:24px;display:grid;grid-template-columns:1fr 380px;gap:32px}
.qual-analysis p{color:#374151;line-height:1.75;font-size:14px;margin-bottom:12px}
.quotes-block{margin-top:16px;border-top:2px solid #E5E7EB;padding-top:16px}
blockquote.quote-item{background:#F0F9FF;border-left:4px solid #2E8B88;padding:10px 14px;margin-bottom:10px;border-radius:0 8px 8px 0;font-style:italic;color:#1B3A5C;font-size:13.5px;line-height:1.6}
.qual-charts{display:flex;flex-direction:column;gap:14px;border-left:1px solid #F0F0F0;padding-left:20px;min-width:0}
.field-chart-wrap{background:#F9FAFB;border-radius:8px;padding:12px;position:relative}
.field-chart-title{font-size:10px;font-weight:700;color:#6B7280;text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px}
.no-chart{color:#9CA3AF;font-size:12px;font-style:italic;padding:12px}
.mm-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(340px,1fr));gap:20px}
.mm-block{background:white;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.mm-title{font-size:13px;font-weight:700;color:#1B3A5C;margin-bottom:12px;border-bottom:2px solid #2E8B88;padding-bottom:6px}
.mm-table{width:100%;border-collapse:collapse;font-size:13px}
.mm-table th{background:#EFF6FF;padding:8px 10px;text-align:left;font-size:11px;color:#1B3A5C;border-bottom:2px solid #2E8B88}
.mm-table td{padding:7px 10px;border-bottom:1px solid #F0F0F0}
.mm-table .pos{color:#2D7A4F;font-weight:700}.mm-table .neg{color:#C0392B;font-weight:700}
.na-section{border:2px solid #7C3AED}
.na-header{background:#7C3AED!important}
.na-badge{background:#5B21B6!important}
blockquote.na-quote{background:#F5F3FF;border-left-color:#7C3AED}
.na-count-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px;margin-bottom:24px}
.na-count-block{background:white;border-radius:10px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.08);border-top:3px solid #7C3AED}
.na-count-title{font-size:11px;font-weight:700;color:#7C3AED;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px}
.cat-header-row td{font-size:13px!important;letter-spacing:.04em}
.card-sub{font-size:11px;color:#9CA3AF;margin-top:2px}
.footer{text-align:center;color:#9CA3AF;font-size:12px;margin-top:48px;padding:24px}
@media(max-width:768px){
  .hero{flex-direction:column;padding:28px 24px;gap:24px}
  .hero-stats{flex-wrap:wrap;justify-content:center}
  .chart-grid{grid-template-columns:1fr}
  .qual-body{grid-template-columns:1fr}
  .qual-charts{display:none}
  .nav-links{display:none}
}
</style>
</head>
<body>
<nav class="nav">
  <div class="nav-title">&#127958; Camp Report """ + str(year) + """</div>
  <div class="nav-links">
    <a href="#overview">Overview</a>
    <a href="#skills">Skill Shifts</a>
    <a href="#classes">By Class</a>
    <a href="#qualitative">Student Voice</a>
    <a href="#mixed">Insights</a>
    <a href="#agency">Agency in Learning</a>
    <a href="#non-attenders">Non-Attenders</a>
  </div>
</nav>
<div class="page">
  <div class="hero" id="overview">
    <div class="hero-text">
      <h1>Camp Report """ + str(year) + """</h1>
      <p>Comprehensive student outcomes: quantitative skill assessments and qualitative survey analysis.</p>
    </div>
    <div class="hero-stats">
      <div class="hero-stat"><div class="big">""" + str(n_students) + """</div><div class="lbl">Students</div></div>
      <div class="hero-stat"><div class="big">""" + str(n_metrics) + """</div><div class="lbl">Skills Tracked</div></div>
      <div class="hero-stat"><div class="big">""" + str(len(QUAL_QUESTIONS)) + """</div><div class="lbl">Survey Qs</div></div>
    </div>
  </div>

  <div class="cards">
    <div class="card green">
      <div class="card-val">""" + str(n_improved) + """/""" + str(n_metrics) + """</div>
      <div class="card-lbl">Skills with positive shift</div>
    </div>
    <div class="card">
      <div class="card-val">""" + overall_shift + """</div>
      <div class="card-lbl">Average shift across all skills</div>
    </div>
    <div class="card amber">
      <div class="card-val">""" + str(peer_pct) + """%</div>
      <div class="card-lbl">Positive sentiment on peer support</div>
    </div>
    <div class="card purple">
      <div class="card-val">""" + str(fav_pct) + """%</div>
      <div class="card-lbl">Positive on favourite part</div>
    </div>
  """ + po_card_html + """
  </div>

  <h2 class="section-title">&#128202; Results by Category</h2>
  <div class="cards">""" + cat_cards_html + """</div>

  <h2 class="section-title" id="skills">&#128208; Skill &amp; Attitude Shifts by Category</h2>
  <table class="metric-table">
    <thead><tr>
      <th>Skill / Area</th>
      <th style="text-align:center">Pre-Camp Avg</th>
      <th style="text-align:center">Post-Camp Avg</th>
      <th style="text-align:center">Mean Shift</th>
      <th style="text-align:center">Median Shift</th>
      <th style="text-align:center">% Improvers</th>
      <th>Change</th>
    </tr></thead>
    <tbody>""" + metric_rows_html + po_metric_row_html + """</tbody>
  </table>

  <div class="chart-grid">
    <div class="chart-box" style="height:""" + str(chart_h + 64) + """px">
      <h3>Score Shift by Skill (post &#8722; pre)</h3>
      <canvas id="shift-chart" style="height:""" + str(chart_h) + """px;width:100%;display:block"></canvas>
    </div>
    <div class="chart-box" style="height:""" + str(chart_h + 64) + """px">
      <h3>Who Improved, Stayed Same, or Declined?</h3>
      <canvas id="dist-chart" style="height:""" + str(chart_h) + """px;width:100%;display:block"></canvas>
    </div>
  </div>

  <div class="chart-grid">
    <div class="chart-box" style="height:""" + str(chart_h + 64) + """px">
      <h3>Pre-Camp vs Post-Camp Average Scores</h3>
      <canvas id="prepost-chart" style="height:""" + str(chart_h) + """px;width:100%;display:block"></canvas>
    </div>
    <div class="chart-box" style="height:300px">
      <h3>Average Shift by Class</h3>
      <canvas id="class-chart" style="height:240px;width:100%;display:block"></canvas>
    </div>
  </div>

  <h2 class="section-title" id="classes">&#128101; Results by Class</h2>
  <div class="breakdown-wrap">
    <table class="breakdown-table">
      <thead><tr>""" + class_header_cells + """</tr></thead>
      <tbody>""" + class_rows_html + """</tbody>
    </table>
  </div>

  <h2 class="section-title" id="qualitative">&#128172; Student Voice &#8212; Qualitative Analysis</h2>
  """ + qual_section_html + """

  <h2 class="section-title" id="mixed">&#128300; Mixed Methods Insights</h2>
  <p style="color:#6B7280;margin-bottom:20px;font-size:13px">
    These tables cross-reference students&#8217; written responses with their quantitative score shifts.
  </p>
  <div class="mm-grid">""" + mm_html + """</div>

  """ + agency_section_html + """

  <h2 class="section-title" id="non-attenders" style="border-color:#7C3AED;color:#7C3AED">
    &#128683; Non-Attender Analysis
  </h2>
  <p style="color:#6B7280;margin-bottom:20px;font-size:13px">
    Responses from students who did <strong>not</strong> attend camp &mdash;
    including reasons for non-attendance and constructive feedback on how the school
    could support future participation. Total non-attenders: <strong>""" + str(n_na_students) + """</strong>.
  </p>
  """ + na_section_html + """
  """ + na_table_html + """
</div>
<div class="footer">Camp Report """ + str(year) + """ &middot; Generated by Camp Analysis Tool &middot; Student data anonymised.</div>

<script>
const PALETTE=['#2E8B88','#1B3A5C','#2D7A4F','#D97706','#5B21B6','#C0392B','#2980B9','#8E44AD','#16A085','#E67E22'];
const TICK_FONT={size:11,family:"'Segoe UI',Arial,sans-serif"};

new Chart(document.getElementById('shift-chart'),{
  type:'bar',
  data:{
    labels:""" + _json.dumps(chart_labels) + """,
    datasets:[{label:'Score Shift',data:""" + _json.dumps(chart_shift) + """,
      backgroundColor:""" + _json.dumps(chart_shift_cols) + """,borderRadius:4}]
  },
  options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
    layout:{padding:{right:8}},
    plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${c.parsed.x>=0?'+':''}${c.parsed.x.toFixed(2)}`}}},
    scales:{
      x:{title:{display:true,text:'Change in score (post minus pre)',font:TICK_FONT},grid:{color:'#F0F0F0'},ticks:{font:TICK_FONT}},
      y:{grid:{display:false},ticks:{font:TICK_FONT}}
    }
  }
});

new Chart(document.getElementById('dist-chart'),{
  type:'bar',
  data:{
    labels:""" + _json.dumps(chart_dist_lbl) + """,
    datasets:[
      {label:'Improved',data:""" + _json.dumps(chart_improved) + """,backgroundColor:'#4CAF50',borderRadius:2},
      {label:'No Change',data:""" + _json.dumps(chart_same) + """,backgroundColor:'#D3D3D3',borderRadius:2},
      {label:'Declined',data:""" + _json.dumps(chart_declined) + """,backgroundColor:'#FF6B6B',borderRadius:2}
    ]
  },
  options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
    layout:{padding:{right:8}},
    plugins:{legend:{position:'top',labels:{font:TICK_FONT}}},
    scales:{
      x:{stacked:true,max:100,title:{display:true,text:'% of students',font:TICK_FONT},ticks:{font:TICK_FONT}},
      y:{stacked:true,grid:{display:false},ticks:{font:TICK_FONT}}
    }
  }
});

new Chart(document.getElementById('prepost-chart'),{
  type:'bar',
  data:{
    labels:""" + _json.dumps(chart_labels) + """,
    datasets:[
      {label:'Pre-Camp',data:""" + _json.dumps(chart_pre) + """,backgroundColor:'#8FAADC',borderRadius:3},
      {label:'Post-Camp',data:""" + _json.dumps(chart_post) + """,backgroundColor:'#70AD47',borderRadius:3}
    ]
  },
  options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
    layout:{padding:{right:8}},
    plugins:{legend:{position:'top',labels:{font:TICK_FONT}}},
    scales:{
      x:{min:0,max:10,title:{display:true,text:'Score (out of 10)',font:TICK_FONT},ticks:{font:TICK_FONT}},
      y:{grid:{display:false},ticks:{font:TICK_FONT}}
    }
  }
});

const cShifts=""" + _json.dumps(class_shift_vals) + """;
new Chart(document.getElementById('class-chart'),{
  type:'bar',
  data:{
    labels:""" + _json.dumps(class_labels) + """,
    datasets:[{label:'Avg Shift',data:cShifts,
      backgroundColor:cShifts.map(v=>v>0?'#2D7A4F':v<0?'#C0392B':'#D3D3D3'),borderRadius:4}]
  },
  options:{responsive:true,maintainAspectRatio:false,
    layout:{padding:{top:4}},
    plugins:{legend:{display:false}},
    scales:{
      y:{title:{display:true,text:'Avg Shift',font:TICK_FONT},grid:{color:'#F0F0F0'},ticks:{font:TICK_FONT}},
      x:{grid:{display:false},ticks:{font:TICK_FONT}}
    }
  }
});

const agLevelEl=document.getElementById('agency-level-chart');
if(agLevelEl){
  new Chart(agLevelEl,{
    type:'bar',
    data:{
      labels:""" + _json.dumps(ag_level_labels) + """,
      datasets:[{label:'Students',data:""" + _json.dumps(ag_level_counts) + """,
        backgroundColor:'#4338CA',borderRadius:4}]
    },
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${c.parsed.y} student(s)`}}},
      scales:{
        y:{beginAtZero:true,ticks:{precision:0,font:TICK_FONT},grid:{color:'#F0F0F0'},
           title:{display:true,text:'Students',font:TICK_FONT}},
        x:{grid:{display:false},ticks:{font:TICK_FONT}}
      }
    }
  });
}

const agGrowthEl=document.getElementById('agency-growth-chart');
if(agGrowthEl){
  const gVals=""" + _json.dumps(ag_growth_vals) + """;
  new Chart(agGrowthEl,{
    type:'bar',
    data:{
      labels:""" + _json.dumps(ag_growth_labels) + """,
      datasets:[{label:'Growth',data:gVals,
        backgroundColor:gVals.map(v=>v>0?'#2D7A4F':v<0?'#C0392B':'#D3D3D3'),borderRadius:3}]
    },
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>{
        const v=c.parsed.x;return ` ${v>=0?'+':''}${v.toFixed(2)} pts (on 10-pt scale)`;
      }}}},
      scales:{
        x:{title:{display:true,text:'Change in Agency score (1–10 scale)',font:TICK_FONT},
           grid:{color:'#F0F0F0'},ticks:{font:TICK_FONT}},
        y:{grid:{display:false},ticks:{font:{size:10}}}
      }
    }
  });
}

const qualData=""" + qual_js + """;
const fieldCanvas={
""" + field_canvas_map_js + """};
Object.entries(qualData).forEach(([field,d])=>{
  const cid=fieldCanvas[field];
  if(!cid) return;
  const el=document.getElementById(cid);
  if(!el) return;
  new Chart(el,{
    type:'bar',
    data:{labels:d.labels,datasets:[{data:d.data,backgroundColor:PALETTE.slice(0,d.data.length),borderRadius:3}]},
    options:{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${c.parsed.x} students`}}},
      scales:{
        x:{grid:{color:'#F0F0F0'},ticks:{font:{size:10}}},
        y:{grid:{display:false},ticks:{font:{size:10},
          callback:function(value){const lbl=this.getLabelForValue(value);return lbl.length>22?lbl.substring(0,20)+'…':lbl;}}}
      }
    }
  });
});
</script>
</body>
</html>"""

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML report written: {html_path}")


def generate_report(student_path, pre_path, post_path, status_label, root, ruby_path=None):
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

        attend_col    = find_col(post_df, ["attend", "camp"])
        non_attend_df = pd.DataFrame()
        if attend_col:
            yes_mask      = post_df[attend_col].astype(str).str.lower().str.contains("yes", na=False)
            non_attend_df = post_df[~yes_mask].copy().reset_index(drop=True)
            post_df       = post_df[yes_mask]

        pre_df  = pre_df.sort_values("Timestamp").drop_duplicates("Email address", keep="last")
        post_df = post_df.sort_values("Timestamp").drop_duplicates("Email address", keep="last")

        status_label.config(text="Status: Processing quantitative data…", fg="blue"); root.update()
        tab1_df, avg_df, dist_df, breakdown_df, merged_df, metrics_processed, post_only_scores, agency_df, agency_summary = process_quantitative(
            students, pre_df, post_df, ruby_path=ruby_path)

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

            # ── Non-attender analysis (while model is still loaded)
            na_df, na_narratives = build_non_attend_analysis(
                non_attend_df, model, tokenizer, status_label, root)

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
            narratives    = [{"Question": "AI Status", "n": 0, "Summary": err}]
            na_df         = pd.DataFrame()
            na_narratives = []
            mm_tables     = []
            qual_charts   = {}

        status_label.config(text="Status: Writing Excel Report…", fg="blue"); root.update()
        docs_dir = os.path.join(os.path.expanduser("~"), "Documents")
        os.makedirs(docs_dir, exist_ok=True)
        xl_path = os.path.join(docs_dir, "Camp_Analysis_Report.xlsx")
        write_excel(xl_path, tab1_df, avg_df, dist_df, breakdown_df,
                    coded_df, summary_metadata, long_df, narratives,
                    metrics_processed, mm_tables, qual_charts, post_only_scores,
                    na_df=na_df, na_narratives=na_narratives,
                    agency_df=agency_df, agency_summary=agency_summary)

        status_label.config(text="Status: Writing HTML Report…", fg="blue"); root.update()
        html_path = os.path.join(docs_dir, "Camp_Report_Presentation.html")
        write_html_report(html_path, tab1_df, avg_df, dist_df, breakdown_df,
                          coded_df, narratives, metrics_processed, mm_tables, long_df,
                          post_only_scores,
                          na_df=na_df, na_narratives=na_narratives,
                          agency_df=agency_df, agency_summary=agency_summary)

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

    # 'clam' is a cross-platform ttk theme that actually honours custom
    # colours (unlike the default macOS 'aqua' theme, which ignores
    # bg/fg on plain tk.Button widgets and made GENERATE REPORTS hard to see)
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Generate.TButton",
                     background="#1B3A5C", foreground="white",
                     font=("Arial", 13, "bold"), padding=10,
                     borderwidth=0, relief="flat")
    style.map("Generate.TButton",
              background=[("active", "#2E8B88"), ("pressed", "#2E8B88")],
              foreground=[("active", "white"), ("pressed", "white")])

    files = {"student": "", "pre": "", "post": "", "ruby": ""}

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
    row("4.  Agency in Learning CSV (optional)", "ruby",
        "Combined CSV from the Student Report Extractor (Ruby Data Converter), incl. Email column")

    tk.Frame(root, height=1, bg="#E5E7EB").pack(fill="x", padx=30, pady=14)

    status = tk.Label(root, text="Status: Waiting for files…",
                      font=("Arial", 10, "italic"), fg="#9CA3AF", bg="#F8FAFC")
    status.pack(pady=(0, 8))

    ttk.Button(root, text="▶  GENERATE REPORTS",
               style="Generate.TButton",
               command=lambda: generate_report(
                   files["student"], files["pre"], files["post"], status, root,
                   ruby_path=files["ruby"] or None)
               ).pack(pady=4, fill="x", padx=30)

    tk.Label(root, text="Generates: ~/Documents/Camp_Analysis_Report.xlsx  +  Camp_Report_Presentation.html",
             font=("Arial", 8), fg="#9CA3AF", bg="#F8FAFC").pack(pady=(4,0))
    tk.Label(root, text="⚡ Fully offline — place chartjs.min.js in the same folder for offline charts",
             font=("Arial", 8), fg="#2E8B88", bg="#F8FAFC").pack(pady=(0,6))

    root.mainloop()


if __name__ == "__main__":
    setup_gui()