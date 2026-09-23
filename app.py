import streamlit as st
import pandas as pd
import os
import time
from datetime import datetime
from twilio.rest import Client
from twilio.base.exceptions import TwilioRestException

st.set_page_config(page_title="Landowner Messaging Platform", page_icon="📩", layout="wide")

# =========================================================================
# CONFIGURATION — all from environment variables set in your hosting
# provider's dashboard. Nothing is hardcoded here.
# =========================================================================
APP_PASSWORD = os.environ.get("APP_PASSWORD")
TWILIO_ACCOUNT_SID = os.environ.get("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = os.environ.get("TWILIO_AUTH_TOKEN")
TWILIO_PHONE_NUMBER = os.environ.get("TWILIO_PHONE_NUMBER")
CONFIG_OK = all([TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, TWILIO_PHONE_NUMBER])

# DATA_DIR should point at your persistent disk's mount path in production
# (e.g. /var/data on Render). Defaults to a local "data" folder for testing.
DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "data"))
os.makedirs(DATA_DIR, exist_ok=True)

contacts_file = os.path.join(DATA_DIR, "Landowner_Contact_List.xlsx")
history_file = os.path.join(DATA_DIR, "Logged_Messages_History.xlsx")
templates_file = os.path.join(DATA_DIR, "Message_Templates.xlsx")

CONTACTS_COLUMNS = ["Landowner", "Phone #", "Tract No.", "Group", "Email"]
HISTORY_COLUMNS = ["Timestamp", "Landowner", "Phone", "Direction", "Message", "Status"]
TEMPLATES_COLUMNS = ["Name", "Message"]


# ---------------------------------------------------------------------------
# Login gate
# ---------------------------------------------------------------------------
def check_password():
    def password_entered():
        if st.session_state.get("password_input") == APP_PASSWORD:
            st.session_state["password_correct"] = True
            st.session_state.pop("password_input", None)
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct"):
        return True

    st.title("📩 Landowner Messaging Platform")
    if not APP_PASSWORD:
        st.error("APP_PASSWORD environment variable isn't set. The app can't verify a login yet.")
        return False
    st.text_input("Password", type="password", key="password_input", on_change=password_entered)
    if st.session_state.get("password_correct") is False:
        st.error("Incorrect password.")
    return False


if not check_password():
    st.stop()


# ---------------------------------------------------------------------------
# File helpers (all live on the persistent disk at DATA_DIR)
# ---------------------------------------------------------------------------
def load_excel_safe(path, expected_columns):
    if not os.path.exists(path):
        df = pd.DataFrame(columns=expected_columns)
        df.to_excel(path, index=False)
        return df
    try:
        df = pd.read_excel(path)
    except Exception as e:
        st.error(f"Error reading {os.path.basename(path)}: {e}")
        return pd.DataFrame(columns=expected_columns)

    for col in expected_columns:
        if col not in df.columns:
            df[col] = ""
    return df[expected_columns]


def save_excel(path, df, columns):
    df.to_excel(path, index=False, columns=columns)


def resolve_phone(raw_phone):
    raw_phone = str(raw_phone).strip()
    clean_phone = "".join(filter(str.isdigit, raw_phone))
    if len(clean_phone) == 10:
        return f"+1{clean_phone}"
    if clean_phone.startswith("1") and len(clean_phone) == 11:
        return f"+{clean_phone}"
    return None


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
st.title("📩 Landowner Messaging Platform")

if not CONFIG_OK:
    st.error(
        "Twilio credentials aren't set as environment variables (TWILIO_ACCOUNT_SID, "
        "TWILIO_AUTH_TOKEN, TWILIO_PHONE_NUMBER). Sending is disabled until this is fixed."
    )

df_contacts = load_excel_safe(contacts_file, CONTACTS_COLUMNS)
df_history = load_excel_safe(history_file, HISTORY_COLUMNS)
df_templates = load_excel_safe(templates_file, TEMPLATES_COLUMNS)

page = st.sidebar.radio("Menu", ["Send Messages", "Manage Landowners", "Message History"])

if st.sidebar.button("Log out"):
    st.session_state["password_correct"] = False
    st.rerun()

landowner_options = (
    sorted(df_contacts["Landowner"].dropna().astype(str).replace("", pd.NA).dropna().unique().tolist())
    if not df_contacts.empty
    else []
)
group_options = (
    sorted([g for g in df_contacts["Group"].dropna().astype(str).unique().tolist() if g.strip()])
    if "Group" in df_contacts.columns and not df_contacts.empty
    else []
)

# ---------------------------------------------------------------------------
# PAGE: Send Messages
# ---------------------------------------------------------------------------
if page == "Send Messages":
    c1, c2 = st.columns(2)
    c1.metric("Landowners Loaded", len(df_contacts))
    c2.metric("Messages Sent (all time)", len(df_history))

    st.write("---")
    st.subheader("Send a Message")

    if st.session_state.pop("_reset_selection", False):
        st.session_state.selected_landowners = []
        st.session_state.select_all_toggle = False

    if len(landowner_options) == 0:
        st.info("No landowners yet. Go to 'Manage Landowners' to upload or add your list.")
    else:
        filter_group = "All"
        if group_options:
            filter_group = st.selectbox("Filter by group:", ["All"] + group_options)

        if filter_group == "All":
            filtered_options = landowner_options
        else:
            filtered_options = sorted(
                df_contacts.loc[df_contacts["Group"].astype(str) == filter_group, "Landowner"].dropna().unique().tolist()
            )

        if "selected_landowners" not in st.session_state:
            st.session_state.selected_landowners = []
        st.session_state.selected_landowners = [
            n for n in st.session_state.selected_landowners if n in filtered_options
        ]

        def _toggle_select_all():
            if st.session_state.get("select_all_toggle"):
                st.session_state.selected_landowners = list(filtered_options)
            else:
                st.session_state.selected_landowners = []

        st.checkbox(
            f"Select all ({len(filtered_options)}){' in this group' if filter_group != 'All' else ''}",
            key="select_all_toggle",
            on_change=_toggle_select_all,
        )

        selected = st.multiselect(
            "Or pick specific landowner(s):",
            options=filtered_options,
            key="selected_landowners",
        )

        template_names = ["(none)"] + df_templates["Name"].dropna().tolist() if not df_templates.empty else ["(none)"]
        chosen_template = st.selectbox("Start from a saved template:", template_names)
        default_message = ""
        if chosen_template != "(none)":
            match = df_templates[df_templates["Name"] == chosen_template]
            if not match.empty:
                default_message = match.iloc[0]["Message"]

        message_text = st.text_area("Message to send:", value=default_message, height=120)

        with st.expander("💾 Save this message as a new template"):
            new_template_name = st.text_input("Template name:")
            if st.button("Save Template"):
                if new_template_name.strip() and message_text.strip():
                    new_row = pd.DataFrame([{"Name": new_template_name.strip(), "Message": message_text}])
                    updated_templates = pd.concat([df_templates, new_row], ignore_index=True)
                    save_excel(templates_file, updated_templates, TEMPLATES_COLUMNS)
                    st.success(f"Saved template '{new_template_name.strip()}'.")
                    st.rerun()
                else:
                    st.warning("Give the template a name and a message first.")

        if selected:
            st.caption(f"This will be sent to **{len(selected)}** landowner(s).")

        ready_to_send = bool(selected) and bool(message_text.strip()) and CONFIG_OK

        confirm = False
        if ready_to_send:
            confirm = st.checkbox(f"I confirm I want to send this message to {len(selected)} landowner(s) now.")

        send_clicked = st.button("Send Message", disabled=not (ready_to_send and confirm), type="primary")

        if send_clicked:
            results = []
            successes = 0
            failures = 0
            total = len(selected)
            progress = st.progress(0.0, text=f"Sending 0 of {total}...")

            for i, recipient_name in enumerate(selected, start=1):
                row = df_contacts[df_contacts["Landowner"] == recipient_name]
                raw_phone = ""
                target_number = None
                status = "Not Attempted"

                if row.empty or "Phone #" not in df_contacts.columns:
                    status = "No phone number on file"
                    failures += 1
                else:
                    raw_phone = str(row.iloc[0]["Phone #"]).strip()
                    target_number = resolve_phone(raw_phone)

                    if not target_number:
                        status = f"Invalid phone number: {raw_phone}"
                        failures += 1
                    else:
                        try:
                            client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
                            message = client.messages.create(
                                body=(message_text.strip() if "reply stop" in message_text.lower() else message_text.strip() + chr(10) + chr(10) + "Central Land Consulting project update. Msg & data rates may apply. Reply STOP to opt out, HELP for help."),
                                from_=TWILIO_PHONE_NUMBER,
                                to=target_number,
                            )
                            status = f"Submitted to Twilio (SID: {message.sid})"
                            successes += 1
                        except TwilioRestException as api_err:
                            status = f"Twilio Error {api_err.code}: {api_err.msg}"
                            failures += 1
                        except Exception as api_err:
                            status = f"Errored Out: {api_err}"
                            failures += 1

                results.append({
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Landowner": recipient_name,
                    "Phone": target_number or raw_phone,
                    "Direction": "Outbound (Twilio Network)",
                    "Message": (message_text.strip() if "reply stop" in message_text.lower() else message_text.strip() + chr(10) + chr(10) + "Central Land Consulting project update. Msg & data rates may apply. Reply STOP to opt out, HELP for help."),
                    "Status": status,
                })

                progress.progress(i / total, text=f"Sent {i} of {total}...")
                time.sleep(0.3)

            results_df = pd.DataFrame(results)
            updated_history = pd.concat([df_history, results_df], ignore_index=True)
            try:
                save_excel(history_file, updated_history, HISTORY_COLUMNS)
            except Exception as e:
                st.error(f"Messages were processed but the log FAILED TO SAVE: {e}")

            st.session_state["last_send_results"] = results_df
            st.session_state["_reset_selection"] = True
            st.rerun()

    if "last_send_results" in st.session_state:
        results_df = st.session_state.pop("last_send_results")
        num_success = results_df["Status"].str.startswith("Submitted").sum()
        num_failure = len(results_df) - num_success
        if num_failure == 0:
            st.success(f"Done — all {num_success} message(s) submitted to Twilio successfully.")
        else:
            st.warning(f"Done — {num_success} succeeded, {num_failure} failed out of {len(results_df)}. See details below.")
        st.dataframe(results_df, width="stretch")

# ---------------------------------------------------------------------------
# PAGE: Manage Landowners
# ---------------------------------------------------------------------------
elif page == "Manage Landowners":
    st.subheader("📇 Manage Landowners")

    with st.expander("📤 Upload / replace the full landowner list from a spreadsheet"):
        st.caption(
            "Upload an .xlsx file with columns: Landowner, Phone #, Tract No. "
            "(Group and Email are optional — they'll be added automatically if missing). "
            "This REPLACES the entire current list."
        )
        uploaded = st.file_uploader("Choose a .xlsx file", type=["xlsx"])
        if uploaded is not None:
            try:
                new_df = pd.read_excel(uploaded)
                missing_required = [c for c in ["Landowner", "Phone #", "Tract No."] if c not in new_df.columns]
                if missing_required:
                    st.error(f"Missing required column(s): {', '.join(missing_required)}. Upload not applied.")
                else:
                    for col in CONTACTS_COLUMNS:
                        if col not in new_df.columns:
                            new_df[col] = ""
                    st.write(f"Preview — {len(new_df)} row(s):")
                    st.dataframe(new_df[CONTACTS_COLUMNS], width="stretch")
                    if st.button("Confirm: Replace landowner list with this upload"):
                        save_excel(contacts_file, new_df, CONTACTS_COLUMNS)
                        st.success(f"Replaced landowner list — {len(new_df)} loaded.")
                        st.rerun()
            except Exception as e:
                st.error(f"Couldn't read that file: {e}")

    st.write("---")
    st.caption(
        "Or edit directly below — click a cell to change it, use the blank bottom row to add a new "
        "landowner, or select a row and press the trash icon to delete it. Click 'Save Changes' when done."
    )

    edited_df = st.data_editor(
        df_contacts,
        num_rows="dynamic",
        width="stretch",
        key="contacts_editor",
    )

    if st.button("Save Changes", type="primary"):
        cleaned = edited_df.dropna(how="all")
        cleaned = cleaned[cleaned["Landowner"].astype(str).str.strip() != ""]
        save_excel(contacts_file, cleaned, CONTACTS_COLUMNS)
        st.success(f"Saved. {len(cleaned)} landowner(s) on file.")
        st.rerun()

# ---------------------------------------------------------------------------
# PAGE: Message History
# ---------------------------------------------------------------------------
elif page == "Message History":
    st.subheader(f"📜 Message History ({len(df_history)} records)")
    if not df_history.empty:
        search = st.text_input("Search by landowner name:")
        display_df = df_history
        if search.strip():
            display_df = df_history[df_history["Landowner"].astype(str).str.contains(search, case=False, na=False)]
        st.dataframe(display_df.sort_values("Timestamp", ascending=False), width="stretch")

        with open(history_file, "rb") as f:
            st.download_button(
                "⬇️ Download full ledger as Excel",
                data=f,
                file_name="Logged_Messages_History.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("No messages logged yet.")
