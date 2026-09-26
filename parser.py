"""
Power Grid Project Extractor
-----------------------------
Upload a construction-project PDF for a power grid company. Gemini reads it
and pulls out each project's details as JSON. Coordinates are left null on
purpose — a separate Overpass-based process fills those in later, matched
back to each project by its "id".

Run with:
    streamlit run app.py
"""

import base64
import json
import os
import re
import time
import traceback
import uuid
from datetime import datetime

import streamlit as st
from dotenv import load_dotenv
from google import genai

load_dotenv()

DEFAULT_MODEL = "gemini-3.8-flash"
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

EXTRACTION_PROMPT = """You are analyzing a document that lists construction projects for a \
power grid / electric utility company named "{company}".

Read the ENTIRE document (including any tables) and identify every distinct \
construction project mentioned (e.g. new transmission lines, substations, \
generation facilities, upgrades, interconnections).

Do NOT attempt to determine or estimate GPS coordinates — a separate process \
handles geocoding. Just extract the location as text, as precisely and \
completely as the document states it.

For EACH project, extract the following fields as a JSON object:

- "project_name": short name/title of the project as written in the document.
- "description": 1-2 sentence plain-language summary of what the project is.
- "location_text": the location exactly as described in the document, verbatim \
where possible (full address, mile markers, route description, etc). Use "" \
if nothing is given.
- "street_address": street address only, if stated. Use null if not given.
- "city": city name, if stated or clearly implied. Use null if not given.
- "county": county name, if stated. Use null if not given.
- "state": state/province, if stated. Use null if not given.
- "zip_code": ZIP/postal code, if stated. Use null if not given.
- "start_date": start or "in service" date if stated, in "YYYY-MM-DD" format. \
If only a year or year-quarter is given, use "YYYY-01-01" and note the \
imprecision in "description". Use null if not stated.
- "end_date": completion/energization date, same rules as start_date. Use \
null if not stated or ongoing/planned with no end date.
- "status": one of "proposed", "planned", "permitting", "under_construction", \
"completed", "in_service", or "unknown".
- "capacity_mw": numeric capacity in megawatts if stated. Use null if not \
applicable/stated.
- "voltage_kv": numeric voltage in kV if stated. Use null if not stated.

Respond with ONLY a raw JSON array of these objects — no markdown code fences, \
no commentary, no surrounding text. If you find no projects, respond with [].
"""


def call_gemini(client, model, contents, timeout_s):
    """One straightforward, synchronous call. Returns the raw output text."""
    interaction = client.interactions.create(
        model=model,
        input=contents,
        generation_config={
            "thinking_level": "low",   # this is extraction, not reasoning — keep it fast
            "max_output_tokens": 24000,
        },
        timeout=timeout_s,
    )
    return (interaction.output_text or "").strip()


def extract_projects_from_pdf(client, model, pdf_bytes, company, timeout_s):
    prompt = EXTRACTION_PROMPT.format(company=company or "the company")
    raw = call_gemini(
        client,
        model,
        contents=[
            {
                "type": "document",
                "data": base64.b64encode(pdf_bytes).decode("utf-8"),
                "mime_type": "application/pdf",
            },
            {"type": "text", "text": prompt},
        ],
        timeout_s=timeout_s,
    )

    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    if not raw:
        raise ValueError("Gemini returned an empty response.")

    data = json.loads(raw)  # let this raise json.JSONDecodeError naturally if malformed
    if not isinstance(data, list):
        raise ValueError(f"Expected a JSON array, got {type(data).__name__}.")
    return data


def projects_to_records(projects, company, source_filename):
    records = []
    for p in projects:
        records.append({
            "id": str(uuid.uuid4()),
            "company": company,
            "project_name": p.get("project_name", "Unnamed project"),
            "description": p.get("description", ""),
            "location_text": p.get("location_text", ""),
            "street_address": p.get("street_address"),
            "city": p.get("city"),
            "county": p.get("county"),
            "state": p.get("state"),
            "zip_code": p.get("zip_code"),
            "latitude": None,   # filled in later by the Overpass geocoding step
            "longitude": None,  # filled in later by the Overpass geocoding step
            "start_date": p.get("start_date"),
            "end_date": p.get("end_date"),
            "status": p.get("status", "unknown"),
            "capacity_mw": p.get("capacity_mw"),
            "voltage_kv": p.get("voltage_kv"),
            "source_document": source_filename,
            "extracted_at": datetime.utcnow().isoformat() + "Z",
        })
    return records


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------

st.set_page_config(page_title="Grid Project Extractor", layout="wide")
st.title("⚡ Power Grid Project Extractor")
st.caption("Upload a project PDF → Gemini extracts project details as JSON, with coordinates left blank for Overpass to fill in later.")

with st.sidebar:
    st.header("Settings")
    if GEMINI_API_KEY:
        st.success("Gemini API key loaded from .env")
    else:
        st.error("GEMINI_API_KEY not found in .env")
    model_name = st.text_input("Model", value=DEFAULT_MODEL)
    timeout_s = st.slider("Request timeout (seconds)", 30, 300, 120)

    st.divider()
    st.subheader("Debug: test connection")
    st.caption("Sends a tiny text-only prompt (no PDF) so you can tell if slowness is the API itself or the document processing.")
    if st.button("Run connection test"):
        t0 = time.time()
        try:
            client = genai.Client(api_key=GEMINI_API_KEY, http_options={"timeout": 30_000})
            raw = call_gemini(client, model_name, contents="Reply with exactly: ok", timeout_s=30)
            st.success(f"Got a response in {time.time() - t0:.1f}s: {raw!r}")
        except Exception as e:
            st.error(f"Failed after {time.time() - t0:.1f}s: {type(e).__name__}: {e}")
            st.code(traceback.format_exc())

company = st.text_input("Company name", placeholder="e.g. Sunrise Power & Light")
pdf_file = st.file_uploader("Project document (PDF)", type=["pdf"])

run_button = st.button("Extract projects", type="primary", disabled=not (GEMINI_API_KEY and company and pdf_file))

if "records" not in st.session_state:
    st.session_state.records = None

if run_button:
    t0 = time.time()
    try:
        with st.spinner(f"Calling Gemini (timeout set to {timeout_s}s)..."):
            client = genai.Client(api_key=GEMINI_API_KEY, http_options={"timeout": timeout_s * 1000})
            pdf_bytes = pdf_file.getvalue()
            projects = extract_projects_from_pdf(client, model_name, pdf_bytes, company, timeout_s)
            st.session_state.records = projects_to_records(projects, company, pdf_file.name)
        st.toast(f"Done in {time.time() - t0:.1f}s")
    except Exception as e:
        st.error(f"Extraction failed after {time.time() - t0:.1f}s: {type(e).__name__}: {e}")
        with st.expander("Full error details"):
            st.code(traceback.format_exc())
        st.session_state.records = None

if st.session_state.records is not None:
    records = st.session_state.records
    st.success(f"Extracted {len(records)} project(s).")
    st.dataframe(records, use_container_width=True)

    st.download_button(
        "Download JSON",
        data=json.dumps(records, indent=2),
        file_name=f"{company.replace(' ', '_').lower()}_projects.json",
        mime="application/json",
    )

    with st.expander("Raw JSON preview"):
        st.json(records)