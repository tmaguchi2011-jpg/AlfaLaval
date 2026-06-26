# =========================================================
# STREAMLIT ALFA LAVAL HMI IMAGE READER APP
# =========================================================

import os
import json
import base64
import requests
from io import BytesIO
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st
from openai import OpenAI

# =========================================================
# STREAMLIT PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Alfa Laval HMI Image Reader",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# PASSWORD PROTECTION
# =========================================================

PASSWORD = ""

pwd = st.text_input("Enter Password", type="password")

if pwd != PASSWORD:
    st.stop()

# =========================================================
# CONFIGURATION
# =========================================================

EXCEL_FILE = "alfa_laval_hmi_reading_backup.xlsx"
SHEET_NAME = "Alfa_Laval_HMI_Readings"
SESSION_STATE_FILE = "alfa_laval_current_session_state.json"

MODEL = "gpt-4.1-mini"
SESSION_VALID_HOURS = 10

try:
    OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]
except Exception:
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

try:
    POWER_AUTOMATE_URL = st.secrets["POWER_AUTOMATE_URL"]
except Exception:
    POWER_AUTOMATE_URL = os.getenv("POWER_AUTOMATE_URL")

# =========================================================
# HMI JSON SCHEMA
# =========================================================

HMI_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "Timestamp": {"type": "string"},
        "SourceFile": {"type": "string"},

        "43FT60": {"type": ["number", "null"]},
        "43CT60": {"type": ["number", "null"]},
        "41PT60": {"type": ["number", "null"]},
        "41TT60": {"type": ["number", "null"]},
        "41TC60": {"type": ["number", "null"]},
        "42PT60": {"type": ["number", "null"]},
        "42PC60": {"type": ["number", "null"]},
        "42DPT60": {"type": ["number", "null"]},
        "42DPC60": {"type": ["number", "null"]},
        "44LT60": {"type": ["number", "null"]},
        "53LT60": {"type": ["number", "null"]},
        "44LC60": {"type": ["number", "null"]},
        "53LC60": {"type": ["number", "null"]},

        "42VC60": {"type": ["string", "null"]},
        "42VC60_Position": {"type": ["number", "null"]},
        "72VC60": {"type": ["string", "null"]},
        "72VC60_Position": {"type": ["number", "null"]},
        "44VC60": {"type": ["string", "null"]},
        "44VC60_Position": {"type": ["number", "null"]},
        "53VC60": {"type": ["string", "null"]},
        "53VC60_Position": {"type": ["number", "null"]},

        "43VA40": {"type": ["string", "null"]},
        "42VA40": {"type": ["string", "null"]},
        "41VA40": {"type": ["string", "null"]},

        "41PF30": {"type": ["string", "null"]},
        "41PF30_Speed": {"type": ["number", "null"]},
        "Cont01": {"type": ["string", "null"]},

        "Emergency_Stop": {"type": ["string", "null"]},
        "74PIC60": {"type": ["string", "null"]}
    },
    "required": [
        "Timestamp", "SourceFile",
        "43FT60", "43CT60", "41PT60", "41TT60", "41TC60",
        "42PT60", "42PC60", "42DPT60", "42DPC60",
        "44LT60", "53LT60", "44LC60", "53LC60",
        "42VC60", "42VC60_Position",
        "72VC60", "72VC60_Position",
        "44VC60", "44VC60_Position",
        "53VC60", "53VC60_Position",
        "43VA40", "42VA40", "41VA40",
        "41PF30", "41PF30_Speed", "Cont01",
        "Emergency_Stop", "74PIC60"
    ]
}

# =========================================================
# FUNCTIONS
# =========================================================

def image_bytes_to_base64(image_bytes):
    return base64.b64encode(image_bytes).decode("utf-8")


def load_session_state():
    if not os.path.exists(SESSION_STATE_FILE):
        return None

    try:
        with open(SESSION_STATE_FILE, "r") as f:
            session_state = json.load(f)

        lot_number = session_state.get("LotNumber")
        operator_name = session_state.get("OperatorName")
        start_time_text = session_state.get("StartTime")

        if not lot_number or not operator_name or not start_time_text:
            return None

        start_time = datetime.strptime(start_time_text, "%Y-%m-%d %H:%M:%S")
        expiration_time = start_time + timedelta(hours=SESSION_VALID_HOURS)

        if datetime.now() <= expiration_time:
            return {
                "LotNumber": lot_number,
                "OperatorName": operator_name,
                "StartTime": start_time,
                "ExpirationTime": expiration_time
            }

        return None

    except Exception:
        return None


def save_session_state(lot_number, operator_name):
    session_state = {
        "LotNumber": lot_number,
        "OperatorName": operator_name,
        "StartTime": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    with open(SESSION_STATE_FILE, "w") as f:
        json.dump(session_state, f, indent=4)


def clear_session_state():
    if os.path.exists(SESSION_STATE_FILE):
        os.remove(SESSION_STATE_FILE)


def extract_hmi_values(image_bytes, filename, lot_number, operator_name):
    if not OPENAI_API_KEY:
        raise ValueError("OPENAI_API_KEY is missing.")

    client = OpenAI(api_key=OPENAI_API_KEY)
    image_b64 = image_bytes_to_base64(image_bytes)

    prompt = """
You are an industrial HMI vision extraction assistant.

Analyze this Alfa Laval HMI screen and extract ONLY process data.

Return a flat JSON object using exactly these field names:

43FT60, 43CT60, 41PT60, 41TT60, 41TC60,
42PT60, 42PC60, 42DPT60, 42DPC60,
44LT60, 53LT60, 44LC60, 53LC60,
42VC60, 42VC60_Position,
72VC60, 72VC60_Position,
44VC60, 44VC60_Position,
53VC60, 53VC60_Position,
43VA40, 42VA40, 41VA40,
41PF30, 41PF30_Speed, Cont01,
Emergency_Stop, 74PIC60.

Rules:

1. For process values:
- Read the variable tag immediately above the value box.
- Extract only the numerical value.
- Remove units such as bar, %, °C, mS/cm, us/cm.
- Return numeric values only.
- If a numeric value cannot be confidently read, return 0.

Never return null for numeric values.

Always return a numeric value.

Examples:

Unreadable pressure = 0
Unreadable temperature = 0
Unreadable flow = 0
Unreadable level = 0
Unreadable valve position = 0
Unreadable pump speed = 0

2. For valves:
- Return OPEN, CLOSED, or PARTIALLY OPEN.
- If a percentage is visible, place the number in the matching _Position field.
- Example: 42VC60 = "OPEN", 42VC60_Position = 100.0.

3. For pumps:
- Return RUNNING or STOPPED.
- If speed percentage is visible, put it in 41PF30_Speed.

4. For green square indicators:
- Return ON if the indicator is green.
- Return OFF if not green or clearly inactive.

5. Ignore:
- Logo
- Menus
- Time and date
- Alarm tabs
- Decorative piping
- Non-process text

6. Never invent values.
7. Never invent variable names.
8. Return only valid JSON.
"""

    response = client.responses.create(
        model=MODEL,
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": prompt},
                    {
                        "type": "input_image",
                        "image_url": f"data:image/jpeg;base64,{image_b64}"
                    }
                ]
            }
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "alfa_laval_hmi_reading",
                "schema": HMI_SCHEMA,
                "strict": True
            }
        }
    )

    data = json.loads(response.output_text)
# ----------------------------------------------------
# Replace unreadable numeric values (None) with 0
# ----------------------------------------------------

numeric_fields = [
    "43FT60",
    "43CT60",
    "41PT60",
    "41TT60",
    "41TC60",
    "42PT60",
    "42PC60",
    "42DPT60",
    "42DPC60",
    "44LT60",
    "53LT60",
    "44LC60",
    "53LC60",
    "42VC60_Position",
    "72VC60_Position",
    "44VC60_Position",
    "53VC60_Position",
    "41PF30_Speed"
]

for field in numeric_fields:
    if data.get(field) is None:
        data[field] = 0
    data["OperatorName"] = operator_name
    data["LotNumber"] = lot_number
    data["Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data["SourceFile"] = filename

    return data


def upload_row_to_sharepoint(data):
    if not POWER_AUTOMATE_URL:
        raise ValueError("POWER_AUTOMATE_URL is missing from Streamlit secrets.")

    response = requests.post(
        POWER_AUTOMATE_URL,
        json=data,
        timeout=30
    )

    if response.status_code not in [200, 201, 202]:
        raise Exception(
            f"SharePoint upload failed: {response.status_code} {response.text}"
        )

    return True


def order_columns(df):
    preferred_order = [
        "OperatorName", "LotNumber", "Timestamp", "SourceFile",
        "43FT60", "43CT60", "41PT60", "41TT60", "41TC60",
        "42PT60", "42PC60", "42DPT60", "42DPC60",
        "44LT60", "53LT60", "44LC60", "53LC60",
        "42VC60", "42VC60_Position",
        "72VC60", "72VC60_Position",
        "44VC60", "44VC60_Position",
        "53VC60", "53VC60_Position",
        "43VA40", "42VA40", "41VA40",
        "41PF30", "41PF30_Speed", "Cont01",
        "Emergency_Stop", "74PIC60"
    ]

    existing_columns = [col for col in preferred_order if col in df.columns]
    remaining_columns = [col for col in df.columns if col not in existing_columns]

    return df[existing_columns + remaining_columns]


def load_existing_excel():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
            return order_columns(df)
        except Exception:
            return pd.DataFrame()

    return pd.DataFrame()


def save_rows_to_excel(rows):
    new_df = pd.DataFrame(rows)
    new_df = order_columns(new_df)

    existing_df = load_existing_excel()

    if not existing_df.empty:
        final_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        final_df = new_df

    final_df = order_columns(final_df)

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        final_df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    return final_df


def dataframe_to_excel_bytes(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    output.seek(0)
    return output


def clear_excel_file():
    if os.path.exists(EXCEL_FILE):
        os.remove(EXCEL_FILE)
        return True

    return False


# =========================================================
# HEADER
# =========================================================

st.markdown(
    """
    <h1 style='color:#FFC107; font-size:42px; font-weight:700; margin-bottom:0px;'>
    UF-1 HMI Image Data-Converter
    </h1>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style='font-style:italic; color:gray; font-size:18px;'>
    Earthrise Engineering Center | Smart Agriculture Research and Investigation
    </div>
    """,
    unsafe_allow_html=True
)

st.caption(
    "This application converts Alfa Laval HMI screenshots into structured process data. "
    "It extracts numerical process values, valve status, pump status, and digital indicators, "
    "then stores the result in Excel and uploads the row to the SharePoint List."
)

st.write(
    "Upload one or more Alfa Laval HMI JPG/JPEG images. "
    "The app will extract the process values, generate an Excel backup, "
    "and upload the readings to SharePoint."
)

st.markdown("---")

# =========================================================
# OPERATOR + LOT SESSION CONTROL
# =========================================================

session_state = load_session_state()

st.subheader("Production Session Information")

if session_state is None:
    st.warning(
        "Operator name and lot number are required before uploading HMI images. "
        "Once confirmed, both will remain active for 10 hours."
    )

    operator_name_input = st.text_input(
        "Enter Operator Name",
        placeholder="Example: John Smith"
    )

    lot_number_input = st.text_input(
        "Enter Lot Number",
        placeholder="Example: LB-2026-001"
    )

    confirm_session_button = st.button("Confirm Operator and Lot Number")

    if confirm_session_button:
        if not operator_name_input.strip():
            st.error("Please enter a valid operator name.")
            st.stop()

        if not lot_number_input.strip():
            st.error("Please enter a valid lot number.")
            st.stop()

        save_session_state(
            lot_number=lot_number_input.strip(),
            operator_name=operator_name_input.strip()
        )

        st.success(
            f"Operator '{operator_name_input.strip()}' and lot number "
            f"'{lot_number_input.strip()}' have been confirmed for the next "
            f"{SESSION_VALID_HOURS} hours."
        )
        st.rerun()

    st.stop()

else:
    current_operator_name = session_state["OperatorName"]
    current_lot_number = session_state["LotNumber"]
    expiration_time = session_state["ExpirationTime"]

    st.success(f"Current operator: {current_operator_name}")
    st.success(f"Current active lot number: {current_lot_number}")

    st.info(
        "The current operator name and lot number will be automatically applied "
        f"to all new HMI readings until: {expiration_time.strftime('%Y-%m-%d %H:%M:%S')}"
    )

    with st.expander("Change or reset operator / lot number"):
        replacement_operator_name = st.text_input(
            "Enter new operator name",
            placeholder="Example: Jane Smith"
        )

        replacement_lot_number = st.text_input(
            "Enter new lot number",
            placeholder="Example: LB-2026-002"
        )

        col_session1, col_session2 = st.columns([1, 1])

        with col_session1:
            if st.button("Confirm New Operator and Lot Number"):
                if not replacement_operator_name.strip():
                    st.error("Please enter a valid new operator name.")
                elif not replacement_lot_number.strip():
                    st.error("Please enter a valid new lot number.")
                else:
                    save_session_state(
                        lot_number=replacement_lot_number.strip(),
                        operator_name=replacement_operator_name.strip()
                    )
                    st.success(
                        f"New operator '{replacement_operator_name.strip()}' "
                        f"and lot number '{replacement_lot_number.strip()}' "
                        f"have been confirmed for the next {SESSION_VALID_HOURS} hours."
                    )
                    st.rerun()

        with col_session2:
            if st.button("Clear Current Operator and Lot Number"):
                clear_session_state()
                st.warning("Current operator and lot number were cleared.")
                st.rerun()

st.markdown("---")

# =========================================================
# FILE UPLOAD
# =========================================================

uploaded_files = st.file_uploader(
    "Select Alfa Laval HMI image files",
    type=["jpg", "jpeg"],
    accept_multiple_files=True
)

process_button = st.button("Process Images")

# =========================================================
# PROCESS IMAGES
# =========================================================

if process_button:
    session_state = load_session_state()

    if session_state is None:
        st.warning(
            "Operator name and lot number expired or missing. "
            "Please confirm them again."
        )
        st.stop()

    active_operator_name = session_state["OperatorName"]
    active_lot_number = session_state["LotNumber"]

    if not uploaded_files:
        st.warning("Please upload at least one JPG or JPEG image.")

    else:
        rows = []

        progress_bar = st.progress(0)
        status_text = st.empty()

        for index, uploaded_file in enumerate(uploaded_files):
            filename = uploaded_file.name
            status_text.write(f"Processing: {filename}")

            try:
                image_bytes = uploaded_file.read()

                result = extract_hmi_values(
                    image_bytes=image_bytes,
                    filename=filename,
                    lot_number=active_lot_number,
                    operator_name=active_operator_name
                )

                rows.append(result)

                try:
                    upload_row_to_sharepoint(result)
                    st.success(f"Completed and uploaded to SharePoint: {filename}")

                except Exception as upload_error:
                    st.warning(
                        f"Completed extraction, but SharePoint upload failed for "
                        f"{filename}: {upload_error}"
                    )

            except Exception as e:
                st.error(f"Failed to process {filename}: {e}")

            progress_bar.progress((index + 1) / len(uploaded_files))

        if rows:
            final_df = save_rows_to_excel(rows)

            st.success(
                f"Finished. {len(rows)} image(s) processed under operator "
                f"'{active_operator_name}' and lot number '{active_lot_number}'."
            )

            st.subheader("Extracted Alfa Laval HMI Data")
            st.dataframe(final_df, use_container_width=True)

            excel_bytes = dataframe_to_excel_bytes(final_df)

            st.download_button(
                label="Download Excel File",
                data=excel_bytes,
                file_name=EXCEL_FILE,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.warning("No valid image data was processed.")

# =========================================================
# DISPLAY EXISTING DATA
# =========================================================

existing_df = load_existing_excel()

if not existing_df.empty:
    st.subheader("Existing Excel Data")
    st.dataframe(existing_df, use_container_width=True)

# =========================================================
# PASSWORD-PROTECTED CLEAR EXCEL SECTION
# =========================================================

with st.expander("Clear Stored Excel Data"):
    st.warning(
        "This will permanently delete the stored Excel file from the app. "
        "Download the Excel file before clearing if you need to save the data."
    )

    clear_password = st.text_input(
        "Enter password to clear stored Excel data",
        type="password",
        key="clear_excel_password"
    )

    confirm_clear_checkbox = st.checkbox(
        "I understand that this will erase the stored Excel data.",
        key="confirm_clear_checkbox"
    )

    clear_button = st.button("Clear Stored Excel Data")

    if clear_button:
        if clear_password != PASSWORD:
            st.error("Incorrect password. Excel data was not cleared.")

        elif not confirm_clear_checkbox:
            st.error("Please confirm that you understand the data will be erased.")

        else:
            deleted = clear_excel_file()

            if deleted:
                st.success("Stored Excel data was cleared successfully.")
            else:
                st.info("No stored Excel file was found.")
