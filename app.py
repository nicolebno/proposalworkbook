import logging
import traceback
import streamlit as st
import pandas as pd
import pdfplumber
import subprocess
import json
import generate_workbook
from io import BytesIO
from build_benefit_summary import BENEFIT_ROWS


logging.basicConfig(
    filename="app.log",
    filemode="a",             # append to the file
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s %(message)s"
)

def safe_generate_workbook(census, rates, other_params):
    try:
        return generate_workbook(census, rates, other_params)
    except Exception:
        logging.error("Workbook generation failed", exc_info=True)
        # optional separate traceback file
        with open("traceback.log", "a") as f:
            traceback.print_exc(file=f)
        st.error("Something went wrong generating the workbook—see app.log for details.")
        return None




# --- Page Setup ---
st.set_page_config(page_title="Proposal App - Upload & Parse", layout="wide")
st.title("📂 Proposal App: Upload & Parse Files")

# Load config
try:
    with open("config.json", "r") as f:
        config = json.load(f)
except FileNotFoundError:
    st.warning("config.json not found. Using default styles.")
    config = {
        "branding": {"primary_color": "#266B8E", "font": "Calibri"},
        "formats": {"currency": "$#,##0.00", "date": "MM/DD/YYYY"},
        "defaults": {}
    }

# Streamlit inputs for contribution modeling
contrib_type = st.selectbox("Contribution Type", ["flat", "percent"])
contrib_employee = st.number_input("Employer Share for Employee", min_value=0.0, value=500.0)
contrib_dependent = st.number_input("Employer Share for Dependents", min_value=0.0, value=250.0)

# Access branding values
BRAND = config["branding"]
FORMATS = config["formats"]
PRIMARY_COLOR = BRAND.get("primary_color", "#266B8E")
CURRENCY_FORMAT = FORMATS.get("currency", "$#,##0.00")
DATE_FORMAT = FORMATS.get("date", "MM/DD/YYYY")
FONT = BRAND.get("font", "Calibri")

st.markdown(f"<style>body {{ font-family: {FONT}; }}</style>", unsafe_allow_html=True)

# --- File Upload Slots ---
census_file = st.file_uploader("Upload Census File (.xlsx or .csv)", type=["xlsx", "csv"])
rate_file = st.file_uploader("Upload Current Rates (.xlsx only)", type=["xlsx"])
renewal_rate_file = st.file_uploader("Upload Renewal Rates (.xlsx only)", type=["xlsx"])
benefit_summary = st.file_uploader("Upload Benefit Summary (.pdf)", type=["pdf"])

# --- Dynamic Plan Inputs ---
st.subheader("🩺 Enter Plan Design Summaries")
if "plans" not in st.session_state:
    st.session_state.plans = []

new_plan_name = st.text_input("Plan Name")
plan_data = {}
for item in BENEFIT_ROWS:
    plan_data[item] = st.text_input(f"{item}", key=f"{item}_{new_plan_name}")

if st.button("➕ Add Plan"):
    st.session_state.plans.append({
        "name": new_plan_name,
        "benefits": plan_data.copy()
    })
    
if st.session_state.plans:
    st.markdown("### ✅ Plans Added")
    for p in st.session_state.plans:
        st.markdown(f"- **{p['name']}**")

# --- Preview Outputs ---
def parse_census(file):
    try:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        return df.head()
    except Exception as e:
        return f"❌ Error parsing census: {e}"

def parse_rate_sheet(file):
    try:
        df = pd.read_excel(file, header=None)
        return df.head(10)
    except Exception as e:
        return f"❌ Error parsing rate sheet: {e}"

def parse_benefit_summary(file):
    try:
        with pdfplumber.open(file) as pdf:
            return f"✅ Parsed PDF: {len(pdf.pages)} pages\nSample text: {pdf.pages[0].extract_text()[:200]}"
    except Exception as e:
        return f"❌ Error parsing PDF: {e}"

if census_file:
    st.subheader("👥 Census Preview")
    st.write(parse_census(census_file))

if rate_file:
    st.subheader("📊 Rate Sheet Preview")
    st.write(parse_rate_sheet(rate_file))

if benefit_summary:
    st.subheader("📄 Benefit Summary Info")
    st.write(parse_benefit_summary(benefit_summary))

# --- Trigger Workbook Generation ---
if census_file and rate_file and renewal_rate_file and benefit_summary:
    if st.button("🚀 Generate Proposal Workbook"):
        try:
            with open("plans.json", "w") as f:
                json.dump(st.session_state.plans, f)

            subprocess.run([
                "python", "generate_workbook.py",
                contrib_type,
                str(contrib_employee),
                str(contrib_dependent)
            ], check=True)

            st.success("Workbook generated successfully!")
            with open("financial_summary_output.xlsx", "rb") as f:
                st.download_button(
                    "📤 Download Finished Proposal Workbook",
                    f,
                    file_name="Proposal.xlsx",
                    key="proposal_download"
                )
        except subprocess.CalledProcessError:
            st.error("Something went wrong while generating the workbook.")
